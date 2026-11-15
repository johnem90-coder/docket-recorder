/**
 * Docket Recorder: Google Apps Script backend.
 *
 * The surveyor talks to the app. Claude works out whether he's describing a NEW visit to a
 * job or a CHANGE to a visit already recorded, and proposes the record. The app reads the
 * proposal back to him (or shows it), he confirms or corrects it, and only then is it saved.
 *
 * The Google Sheet has three tabs:
 *   Visits    one row per visit to a job (the main record)
 *   Dockets   one row per docket: its client and site address (new visits default to these)
 *   Settings  the home address that trips start and end at
 * Driving distances (start → site → end) come from Google Maps unless he says the km himself.
 *
 * This file is the back end. The phone app itself is docs/index.html, hosted on GitHub Pages;
 * it reaches this script through doPost() below, protected by a passcode.
 *
 * ───────────────────────────────────────────────────────────────────────────────
 *  TO CUSTOMISE: edit the COLUMNS list below. That's usually all you need.
 *   - Rename a column: change its `header`.
 *   - Add a column Claude should fill in: add a line with a new `key`, a `header`
 *     and an `ask` (a plain-English instruction telling Claude what to put there).
 *   - `required: true` means the app tells him when that detail is missing.
 *   - Reorder columns: move the lines up or down.
 *  After changing columns, run setupSheet() again so the header row is rewritten.
 *  (Keep the docket, date, client, site, address, time and km columns: the app uses them.)
 * ───────────────────────────────────────────────────────────────────────────────
 */

// Which Claude model reads the notes. Change this one line to switch models.
const MODEL = 'claude-sonnet-5-5';

// Tab names.
const VISITS_TAB = 'Visits';
const DOCKETS_TAB = 'Dockets';
const SETTINGS_TAB = 'Settings';

// Optional: a two-letter country code (e.g. 'nz', 'au', 'gb') to help Google Maps pick the
// right place when an address is ambiguous. Leave empty to let Google decide.
const MAPS_REGION = '';

// The columns of the Visits tab, left to right. Each row is one visit to one docket.
// Columns WITH an `ask` are filled in by Claude; the others by this script.
const COLUMNS = [
  { key: 'recorded_at',   header: 'Recorded at' },
  { key: 'docket_number', header: 'Docket #', required: true,
    ask: 'The docket or job number exactly as spoken (e.g. "4521" or "J-2024-17"). Empty if not mentioned. Never make one up.' },
  { key: 'job_date',      header: 'Job date', required: true,
    ask: 'Date the work was done, as YYYY-MM-DD. Work out words like "today" or "yesterday" from the date you are given. If no date is mentioned, use today.' },
  { key: 'client',        header: 'Client', required: true,
    ask: 'Client or customer name. Empty if not mentioned.' },
  { key: 'site',          header: 'Site address', required: true,
    ask: 'Address of the job site he drove to for this visit. Empty if not mentioned (the docket\'s address is used).' },
  { key: 'start_address', header: 'Start address',
    ask: 'Where he drove from to get to this visit, ONLY if he says it wasn\'t home, e.g. "came straight from the Olsen job" (use that job\'s site address). Empty means he started from home.' },
  { key: 'end_address',   header: 'End address',
    ask: 'Where he drove to after this visit, ONLY if he says it wasn\'t home, e.g. "went on to the Patel job" (use that job\'s site address). Empty means he went home.' },
  { key: 'arrived',       header: 'Arrived', required: true,
    ask: 'Time he arrived on site, as 24-hour HH:mm (e.g. "quarter past nine" → "09:15"). Empty if not mentioned.' },
  { key: 'departed',      header: 'Departed', required: true,
    ask: 'Time he left the site, as 24-hour HH:mm. Empty if not mentioned.' },
  { key: 'hours',         header: 'Hours on site' },
  { key: 'work_done',     header: 'Work done', required: true,
    ask: 'A short, clear summary of the work done on site, in his words where possible. Use "; " between separate tasks.' },
  { key: 'equipment',     header: 'Equipment / materials',
    ask: 'Equipment used or materials used/left on site (e.g. "GPS rover; 4 pegs"). Empty if not mentioned.' },
  { key: 'travel_km',     header: 'Travel km', required: true, auto: true,
    ask: 'Kilometres driven, digits only, ONLY if he says the distance himself (e.g. "it was 42 k"). Otherwise empty: the app works it out from Google Maps.' },
  { key: 'km_source',     header: 'Km from' },
  { key: 'follow_up',     header: 'Follow-up needed',
    ask: 'Anything that still needs doing, a return visit, or something to tell the office/client. Empty if none.' },
  { key: 'notes',         header: 'Other notes',
    ask: 'Anything else he said that does not fit the other fields. Empty if nothing.' },
  { key: 'transcript',    header: 'Raw transcript' },
  { key: 'status',        header: 'Status' },
  { key: 'reviewed',      header: 'Reviewed ✓' },
];

// The columns of the Dockets tab: one row per docket.
const DOCKET_COLUMNS = [
  { key: 'docket_number', header: 'Docket #' },
  { key: 'client',        header: 'Client' },
  { key: 'site',          header: 'Site address' },
  { key: 'created',       header: 'Created' },
];

// The Settings tab: one row per setting (label in column A, value in column B).
const SETTINGS = [
  { key: 'home_address', label: 'Home address' },
];

// The fields Claude fills in (every column that has an `ask`).
const AI_COLUMNS = COLUMNS.filter(function (c) { return c.ask; });
const AI_KEYS = AI_COLUMNS.map(function (c) { return c.key; });
const ROUTE_KEYS = ['start_address', 'site', 'end_address'];

// Columns that can be changed by hand on the phone's Edit screen: everything Claude fills in,
// plus the office's Reviewed tick. (Hours and km are re-worked out on save.)
const EDITABLE_KEYS = AI_KEYS.concat(['reviewed']);


/* ═════════════════════════════  SHEET SETUP  ═════════════════════════════ */

/**
 * Run this once from the Apps Script editor (and again after changing COLUMNS).
 * Creates the Visits, Dockets and Settings tabs if needed and writes tidy header rows.
 * (A tab called "Dockets" from an earlier version that holds visits is renamed to "Visits".)
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let visits = ss.getSheetByName(VISITS_TAB);
  const old = ss.getSheetByName(DOCKETS_TAB);
  if (!visits && old && old.getRange(1, 1).getValue() === 'Recorded at') { old.setName(VISITS_TAB); visits = old; }
  visits = visits || ss.insertSheet(VISITS_TAB, 0);
  writeHeader(visits, COLUMNS.map(function (c) { return c.header; }));
  visits.getRange(2, colIndex('docket_number'), visits.getMaxRows() - 1, 1).setNumberFormat('@');
  COLUMNS.forEach(function (c, i) {
    const wide = ['work_done', 'transcript', 'follow_up', 'notes', 'site', 'start_address', 'end_address'].indexOf(c.key) !== -1;
    visits.setColumnWidth(i + 1, wide ? 240 : 105);
  });

  const dockets = ss.getSheetByName(DOCKETS_TAB) || ss.insertSheet(DOCKETS_TAB, 1);
  writeHeader(dockets, DOCKET_COLUMNS.map(function (c) { return c.header; }));
  dockets.getRange(2, 1, dockets.getMaxRows() - 1, 1).setNumberFormat('@');
  dockets.setColumnWidth(1, 105); dockets.setColumnWidth(2, 200); dockets.setColumnWidth(3, 300); dockets.setColumnWidth(4, 140);

  const settings = ss.getSheetByName(SETTINGS_TAB) || ss.insertSheet(SETTINGS_TAB, 2);
  writeHeader(settings, ['Setting', 'Value']);
  SETTINGS.forEach(function (s, i) { settings.getRange(i + 2, 1).setValue(s.label).setFontWeight('bold'); });
  settings.setColumnWidth(1, 160); settings.setColumnWidth(2, 360);

  Logger.log('Tabs ready: %s, %s, %s. Put the home address in the Settings tab (or in the app).', VISITS_TAB, DOCKETS_TAB, SETTINGS_TAB);
}

function writeHeader(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8eef7').setWrap(true);
  sheet.setFrozenRows(1);
}


/* ═════════════════════════════  WEB APP  ═════════════════════════════ */

/*
 * The phone app is a web page hosted on GitHub Pages (docs/index.html). It talks to this script
 * by sending requests to the web-app link (the one ending in /exec). Because that link has to be
 * open to "Anyone" for the page to reach it, every request must include the passcode stored in
 * Project Settings → Script properties → APP_PASSCODE. Without it, nothing is read or saved.
 */

// The only functions the phone app is allowed to call.
const API_ACTIONS = {
  getAppData: getAppData,
  proposeEntry: proposeEntry,
  commitEntry: commitEntry,
  updateVisit: updateVisit,
  saveDocket: saveDocket,
  saveSettings: saveSettings,
};

/** Handles a request from the phone app: { passcode, action, args } → { ok, result } or { ok: false, error }. */
function doPost(e) {
  let out;
  try {
    const request = JSON.parse(e.postData.contents);
    const passcode = PropertiesService.getScriptProperties().getProperty('APP_PASSCODE');
    if (!passcode) throw coded('setup', 'APP_PASSCODE is not set in Script properties.');
    if (String(request.passcode || '') !== passcode) throw coded('passcode', 'Wrong passcode.');
    const action = API_ACTIONS[request.action];
    if (!action) throw new Error('Unknown request: ' + request.action);
    out = { ok: true, result: action.apply(null, request.args || []) };
  } catch (err) {
    out = { ok: false, error: err.message, code: err.code || '' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Opening the web-app link in a browser just says it's running. */
function doGet() {
  return ContentService.createTextOutput('Docket Recorder is running. Open the app from its GitHub Pages link.');
}

function coded(code, message) {
  const err = new Error(message);
  err.code = code;
  return err;
}

/** Everything the phone app needs: visits (oldest first), dockets, settings and column names. */
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet(VISITS_TAB);
  return {
    visits: readVisits(),
    dockets: readDockets(),
    settings: readSettings(),
    columns: COLUMNS.map(function (c) {
      return { key: c.key, header: c.header, editable: EDITABLE_KEYS.indexOf(c.key) !== -1, required: !!c.required };
    }),
    sheetUrl: ss.getUrl() + '#gid=' + sheet.getSheetId(),
  };
}


/* ═════════════════════════  THE VOICE CONVERSATION  ═════════════════════════
 *
 *  1. proposeEntry(): he has said something. Claude proposes a new visit, or a change to an
 *     existing visit, plus a short sentence for the app to read back. Nothing is saved.
 *  2. He answers. A correction goes back to proposeEntry() with the earlier proposal.
 *  3. commitEntry(): he said yes. The proposal is saved to the sheet.
 *
 *  `scope` says where he started talking from:
 *    { type: 'any' }                   main screen: a new visit, or a change to any visit
 *    { type: 'docket', docket: '4521' } a docket's page: a new visit to it, or a change to one of its visits
 *    { type: 'row', row: 12 }          a visit's Edit screen: a change to that visit
 */

/**
 *   notes:    everything he has said in this conversation, in order (first note, then replies)
 *   previous: the last proposal he is replying to, or null
 */
function proposeEntry(notes, scope, previous) {
  notes = cleanNotes(notes);
  if (!notes.length) throw new Error('Nothing to check: the note is empty.');
  scope = scope || { type: 'any' };
  const ctx = loadContext();

  const message = [
    todayLine(),
    ctx.home ? 'His home address: ' + ctx.home : 'His home address is not set.',
    scopeContext(scope, ctx),
    previous ? 'Your last proposal, which he is replying to:\n' + JSON.stringify({
      action: previous.action, target_row: previous.row || 0, fields: pick(previous.fields || {}, AI_KEYS), you_said: previous.confirmation,
    }, null, 2) : '',
    'What he has said so far, in order:\n' + notes.map(function (n, i) { return (i + 1) + '. """' + n + '"""'; }).join('\n'),
  ].filter(Boolean).join('\n\n');

  const answer = callClaude(buildProposalInstructions(scope), message, proposalSchema());
  return finishProposal(answer, scope, ctx);
}

/** Saves a confirmed proposal. With no proposal (Claude couldn't be reached), saves the raw note. */
function commitEntry(proposal, notes, scope) {
  notes = cleanNotes(notes);
  scope = scope || { type: 'any' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!proposal) {
      if (scope.type === 'row') throw new Error('Couldn\'t check that change, so nothing was saved. You can edit the fields by hand.');
      return appendVisit({ docket_number: scope.docket || '' }, notes.join(' / '),
        ['Saved without being checked by Claude: fill in the details by hand']);
    }
    if (proposal.action === 'new') {
      const problems = proposal.needs_review && proposal.review_reason ? [proposal.review_reason] : [];
      return appendVisit(proposal.fields, notes.join(' / '), problems);
    }
    if (proposal.action === 'update') return applyChange(proposal, notes);
    throw new Error('It isn\'t clear yet which visit this is for, so nothing was saved.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saves edits made by hand on the phone's Edit screen to the visit's row.
 *   row, fingerprint: which visit (the fingerprint checks the row hasn't moved in the sheet)
 *   changes:          { key: newValue } for the editable columns
 */
function updateVisit(row, fingerprint, changes) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const visit = findVisit(readVisits(), row, fingerprint);
    const updated = Object.assign({}, visit);
    EDITABLE_KEYS.forEach(function (key) {
      if (changes[key] !== undefined) updated[key] = key === 'reviewed' ? changes[key] === true : String(changes[key]).trim();
    });
    updated.hours = hoursBetween(updated.arrived, updated.departed);
    applyDistance(updated, readSettings().home_address, visit);
    writeCells(getSheet(VISITS_TAB), visit.row, updated, EDITABLE_KEYS.concat(['hours', 'km_source']));
    ensureDocket(updated);
    return { docket: updated.docket_number || '', row: visit.row };
  } finally {
    lock.releaseLock();
  }
}

/** Saves a docket's client and site address (from the phone's Edit docket screen). */
function saveDocket(docketNumber, changes) {
  docketNumber = String(docketNumber || '').trim();
  if (!docketNumber) throw new Error('No docket number.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet(DOCKETS_TAB);
    const found = readDockets().filter(function (d) { return d.docket_number === docketNumber; })[0];
    const row = found ? found.row : sheet.getLastRow() + 1;
    if (!found) sheet.getRange(row, 1).setNumberFormat('@').setValue(docketNumber);
    if (!found) sheet.getRange(row, 4).setValue(new Date());
    if (changes.client !== undefined) sheet.getRange(row, 2).setValue(String(changes.client).trim());
    if (changes.site !== undefined) sheet.getRange(row, 3).setValue(String(changes.site).trim());
    return { docket: docketNumber };
  } finally {
    lock.releaseLock();
  }
}

/** Saves settings from the phone. The home address is checked with Google Maps first. */
function saveSettings(changes) {
  const sheet = getSheet(SETTINGS_TAB);
  const result = {};
  if (changes.home_address !== undefined) {
    const typed = String(changes.home_address).trim();
    const place = typed ? findPlace(typed) : '';
    if (typed && !place) throw new Error('Google Maps couldn\'t find that address. Try adding the town or postcode.');
    writeSetting(sheet, 'home_address', place);
    result.home_address = place;
  }
  return result;
}


/* ═════════════════════════════  SAVING  ═════════════════════════════ */

/** Adds a new visit as a new row. `problems` are reasons the office should check it. */
function appendVisit(fields, transcript, problems) {
  const ctx = loadContext();
  fields = Object.assign({}, fields);
  const docket = ctx.dockets[fields.docket_number];
  if (docket) {
    if (!fields.client) fields.client = docket.client;
    if (!fields.site) fields.site = docket.site;
  }
  if (!fields.km_source && fields.travel_km === undefined) applyDistance(fields, ctx.home, null);
  if (!fields.start_address) fields.start_address = ctx.home;
  if (!fields.end_address) fields.end_address = ctx.home;

  const values = Object.assign({}, fields, {
    recorded_at: new Date(),
    hours: hoursBetween(fields.arrived, fields.departed),
    transcript: transcript,
    status: statusFor(fields, problems),
    reviewed: false,
  });
  const sheet = getSheet(VISITS_TAB);
  sheet.appendRow(COLUMNS.map(function (c) {
    const v = values[c.key];
    return (v === undefined || v === null) ? '' : v;
  }));
  const row = sheet.getLastRow();
  sheet.getRange(row, colIndex('reviewed')).insertCheckboxes();
  sheet.getRange(row, colIndex('docket_number')).setNumberFormat('@').setValue(values.docket_number || '');
  ensureDocket(values);
  return { action: 'new', docket: values.docket_number || '', row: row, status: values.status };
}

/** Applies a confirmed voice change to an existing visit's row. */
function applyChange(proposal, notes) {
  const visit = findVisit(readVisits(), proposal.row, proposal.fingerprint);
  const updated = Object.assign({}, visit);
  AI_KEYS.concat(['km_source']).forEach(function (k) {
    if (proposal.fields[k] !== undefined) updated[k] = String(proposal.fields[k]).trim();
  });
  updated.hours = hoursBetween(updated.arrived, updated.departed);

  // Keep a record of what was said, so the office can see why the row changed.
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM HH:mm');
  updated.transcript = (visit.transcript ? visit.transcript + '\n' : '') + '[Changed ' + stamp + '] ' + notes.join(' / ');

  // A change to a visit the office already ticked needs ticking again.
  const problems = [];
  if (proposal.needs_review && proposal.review_reason) problems.push(proposal.review_reason);
  if (visit.reviewed) problems.push('Changed after it was reviewed');
  updated.status = statusFor(updated, problems);
  if (visit.reviewed) updated.reviewed = false;

  writeCells(getSheet(VISITS_TAB), visit.row, updated, AI_KEYS.concat(['hours', 'km_source', 'transcript', 'status', 'reviewed']));
  ensureDocket(updated);
  return { action: 'update', docket: updated.docket_number || '', row: visit.row, status: updated.status };
}

/** Adds a docket to the Dockets tab the first time it's seen, or fills in a blank client/address. */
function ensureDocket(fields) {
  const number = String(fields.docket_number || '').trim();
  if (!number) return;
  const found = readDockets().filter(function (d) { return d.docket_number === number; })[0];
  const sheet = getSheet(DOCKETS_TAB);
  if (!found) {
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 4).setValues([[number, fields.client || '', fields.site || '', new Date()]]);
    return;
  }
  if (!found.client && fields.client) sheet.getRange(found.row, 2).setValue(fields.client);
  if (!found.site && fields.site) sheet.getRange(found.row, 3).setValue(fields.site);
}

/** "OK", or "Needs review: …" listing anything the office should check. */
function statusFor(fields, problems) {
  problems = (problems || []).slice();
  if (!fields.docket_number) problems.push('No docket number');
  if (!fields.arrived || !fields.departed) problems.push('Arrival or departure time missing');
  return problems.length ? 'Needs review: ' + unique(problems).join('; ') : 'OK';
}

/** Writes the given keys of `values` into one row of the Visits tab. */
function writeCells(sheet, row, values, keys) {
  keys.forEach(function (key) {
    const cell = sheet.getRange(row, colIndex(key));
    if (key === 'docket_number') cell.setNumberFormat('@');
    if (key === 'reviewed') cell.insertCheckboxes();
    cell.setValue(values[key] === undefined || values[key] === null ? '' : values[key]);
  });
}


/* ═════════════════════════════  DRIVING DISTANCE  ═════════════════════════════ */

/**
 * Fills in the start and end addresses (home unless he said otherwise) and works out the driving
 * distance start → site → end with Google Maps, unless he said the km himself.
 * `before` is the visit as it was (for a change), so an unchanged route isn't looked up again.
 * Changes `fields` in place and returns a sentence for the read-back ('' if nothing to say).
 */
function applyDistance(fields, home, before) {
  if (!fields.start_address) fields.start_address = home || '';
  if (!fields.end_address) fields.end_address = home || '';
  const km = String(fields.travel_km || '').trim();
  const kmBefore = before ? String(before.travel_km || '').trim() : '';

  // He said the distance himself.
  if (km && (!before || km !== kmBefore)) { fields.km_source = 'Entered'; return ''; }

  // A change that doesn't touch the route keeps the distance it had.
  const sameRoute = before && kmBefore && ROUTE_KEYS.every(function (k) { return String(fields[k] || '') === String(before[k] || ''); });
  if (sameRoute) { fields.travel_km = kmBefore; fields.km_source = before.km_source || ''; return ''; }

  fields.travel_km = ''; fields.km_source = '';
  if (!fields.site) return '';
  if (!fields.start_address || !fields.end_address) {
    return 'I couldn\'t work out the driving distance because the home address isn\'t set.';
  }
  const found = routeKm([fields.start_address, fields.site, fields.end_address]);
  if (found === null) return 'I couldn\'t find ' + shortPlace(fields.site) + ' on Google Maps, so I don\'t have the driving distance.';
  fields.travel_km = String(found);
  fields.km_source = 'Google Maps';
  const spoken = Math.round(found);
  if (kmBefore) return 'The driving distance changes from ' + Math.round(Number(kmBefore)) + ' to ' + spoken + ' kilometres.';
  if (fields.start_address === home && fields.end_address === home) return 'Driving there and back from home is ' + spoken + ' kilometres.';
  return 'Driving from ' + placeName(fields.start_address, home) + ' to the site and on to ' + placeName(fields.end_address, home) +
    ' is ' + spoken + ' kilometres.';
}

/** Driving distance in km (one decimal) through the given stops, or null if Maps can't route it. */
function routeKm(stops) {
  const cache = CacheService.getScriptCache();
  const key = 'km:' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, stops.join('|')));
  const hit = cache.get(key);
  if (hit) return hit === 'none' ? null : Number(hit);
  let km = null;
  try {
    const finder = Maps.newDirectionFinder()
      .setOrigin(stops[0])
      .setDestination(stops[stops.length - 1])
      .setMode(Maps.DirectionFinder.Mode.DRIVING);
    stops.slice(1, -1).forEach(function (s) { finder.addWaypoint(s); });
    if (MAPS_REGION) finder.setRegion(MAPS_REGION);
    const res = finder.getDirections();
    if (res && res.routes && res.routes.length) {
      const metres = res.routes[0].legs.reduce(function (sum, leg) { return sum + leg.distance.value; }, 0);
      km = Math.round(metres / 100) / 10;
    }
  } catch (err) {
    Logger.log('Maps error: ' + err);
  }
  cache.put(key, km === null ? 'none' : String(km), km === null ? 600 : 21600);
  return km;
}

/** Google Maps' full version of an address, or '' if it can't be found. */
function findPlace(address) {
  try {
    const geocoder = Maps.newGeocoder();
    if (MAPS_REGION) geocoder.setRegion(MAPS_REGION);
    const res = geocoder.geocode(address);
    return res && res.results && res.results.length ? res.results[0].formatted_address : '';
  } catch (err) {
    Logger.log('Maps error: ' + err);
    return '';
  }
}

/** "home", or the first part of an address ("48 Kowhai Lane"), for reading out. */
function placeName(address, home) {
  return address === home ? 'home' : shortPlace(address);
}
function shortPlace(address) {
  return String(address || '').split(',')[0].trim() || 'that address';
}


/* ═════════════════════════════  CLAUDE  ═════════════════════════════ */

/** The home address, dockets and visits, read once per request. */
function loadContext() {
  const visits = readVisits();
  const dockets = {};
  // Start from what the visits say, then let the Dockets tab override it.
  visits.forEach(function (v) {
    if (!v.docket_number) return;
    const d = dockets[v.docket_number] || (dockets[v.docket_number] = { client: '', site: '' });
    if (v.client) d.client = v.client;
    if (v.site && !d.site) d.site = v.site;
  });
  readDockets().forEach(function (d) {
    const k = dockets[d.docket_number] || (dockets[d.docket_number] = { client: '', site: '' });
    if (d.client) k.client = d.client;
    if (d.site) k.site = d.site;
  });
  return { visits: visits, dockets: dockets, home: readSettings().home_address || '' };
}

/** What Claude is told about the dockets and visits, depending on where he started talking. */
function scopeContext(scope, ctx) {
  const visits = ctx.visits;
  const docketList = Object.keys(ctx.dockets).slice(-150).map(function (d) {
    return d + ' | ' + (ctx.dockets[d].client || '-') + ' | ' + (ctx.dockets[d].site || '-');
  });
  const docketsText = 'Dockets (number | client | site address):\n' + (docketList.join('\n') || 'none yet');

  if (scope.type === 'row') {
    const v = visits.filter(function (x) { return x.row === Number(scope.row); })[0];
    if (!v) throw new Error('That visit couldn\'t be found. It may have been changed in the sheet.');
    return 'He is changing this visit (row ' + v.row + '):\n' + JSON.stringify(pick(v, AI_KEYS), null, 2) +
      '\n\nOther visits to docket ' + (v.docket_number || '-') + ':\n' +
      (visitLines(visits.filter(function (x) { return x.docket_number && x.docket_number === v.docket_number && x.row !== v.row; }), 10, ctx.home) || 'none') +
      '\n\n' + docketsText;
  }
  if (scope.type === 'docket') {
    const own = visits.filter(function (x) { return x.docket_number === scope.docket; });
    const k = ctx.dockets[scope.docket] || {};
    return 'He is on the page for docket ' + scope.docket + ' (client: ' + (k.client || 'unknown') + ', site address: ' +
      (k.site || 'unknown') + ').\nIts visits (newest first):\n' + (visitLines(own, 30, ctx.home) || 'none yet') + '\n\n' + docketsText;
  }
  return docketsText + '\n\nMost recently recorded visits (newest first):\n' + (visitLines(visits, 40, ctx.home) || 'none yet');
}

/** One line per visit, newest first, for Claude to pick from. */
function visitLines(visits, max, home) {
  const where = function (a) { return !a || a === home ? 'home' : a; };
  return visits.slice()
    .sort(function (a, b) { return (b.job_date || '').localeCompare(a.job_date || '') || b.row - a.row; })
    .slice(0, max)
    .map(function (v) {
      return 'row ' + v.row + ' | docket ' + (v.docket_number || 'none') + ' | ' + (v.job_date || 'no date') + ' | ' +
        (v.client || '-') + ' | ' + (v.site || '-') + ' | ' + (v.arrived || '?') + '–' + (v.departed || '?') +
        ' | from ' + where(v.start_address) + ', then to ' + where(v.end_address) +
        (v.travel_km ? ' | ' + v.travel_km + ' km' : '') + ' | work: ' + String(v.work_done || '').slice(0, 100) +
        (v.follow_up ? ' | follow-up: ' + String(v.follow_up).slice(0, 60) : '');
    }).join('\n');
}

/** The instructions Claude gets for proposing a record. */
function buildProposalInstructions(scope) {
  const required = AI_COLUMNS.filter(function (c) { return c.required && !c.auto; }).map(function (c) { return c.header; });
  const where = scope.type === 'row'
    ? 'He started from one visit\'s Edit screen, so this is always a change to that visit: action "update" with its row number.'
    : scope.type === 'docket'
      ? 'He started from docket ' + scope.docket + '\'s page, so a new visit belongs to docket ' + scope.docket +
        ', and a change must be to one of that docket\'s visits.'
      : 'He started from the main screen, so this could be a new visit or a change to any visit listed.';
  return [
    'You help a land surveyor keep his job records by voice while he drives. Each record is one visit',
    'to a job; a docket can have several visits. From what he says, work out whether he is describing',
    'a NEW visit or a CHANGE to a visit already recorded, and propose the record. The app reads your',
    'confirmation out loud and he answers. Nothing is saved until he says yes.',
    '',
    'What he says comes from phone speech-to-text, so expect run-on sentences, filler words and',
    'mis-heard words. Fix obvious speech-to-text mistakes, but never invent facts. Leave a field',
    'empty if he hasn\'t said it.',
    '',
    'Every field must hold the actual information, never a reference to it. If he says things like',
    '"same place as last time", "same location as the previous visit", "same client" or "same hours',
    'as yesterday", look up the real value in the dockets and visits you are given and use that.',
    '',
    'Driving: trips normally start and end at his home, and the app works out the distance with',
    'Google Maps. Only fill start_address or end_address when he says he came from, or went on to,',
    'somewhere other than home; use a docket\'s site address when he names another job. Only fill',
    'travel_km when he says the distance himself.',
    '',
    'Choosing the action:',
    '- "new": a visit that isn\'t recorded yet. target_row is 0.',
    '- "update": correcting or adding to a visit that is already recorded, e.g. "on Tuesday\'s Olsen',
    '  visit I actually left at noon" or "add to the last Henderson visit that I left two pegs".',
    '  target_row is that visit\'s row number from the list. fields is the complete record after the',
    '  change: copy every field he didn\'t ask to change exactly as it is.',
    '- "unclear": you can\'t tell which visit he means, or whether it\'s new. Ask him in the confirmation.',
    '- "cancel": he wants to stop and save nothing at all, e.g. "cancel that", "no, forget the whole',
    '  thing", "I don\'t want to save this". (Clearing one field, like "cancel the follow-up", is an',
    '  update, not a cancel.) The confirmation is then just "Okay, I won\'t save that."',
    where,
    '',
    'For a new visit with no docket number: if it clearly matches exactly one existing docket (same',
    'client or site), use that docket and set needs_review true with the reason "Docket matched by name".',
    'If he says a docket number, always use the number he said.',
    '',
    'If he is replying to your last proposal, apply his reply to it (a correction, an extra detail or',
    'an answer to your question) and propose the whole record again.',
    '',
    'Required details: ' + required.join(', ') + '.',
    '',
    'confirmation: what the app will read aloud to him. Plain spoken English, no symbols, lists or',
    'abbreviations, under 60 words. For a new visit, briefly say the docket, client, site, day, times',
    'and work. For a change, say which visit (docket and day) and only what is changing. Don\'t',
    'mention driving or kilometres unless he said the km himself: the app adds the distance. Then name',
    'any required details that are still missing, and end with exactly "Shall I save it?". If the',
    'action is "unclear", ask your question instead. Say times naturally ("half past eight") and days',
    'relative to today ("today", "yesterday", "Tuesday the twelfth").',
    '',
    'Fields:',
    AI_COLUMNS.map(function (c) { return '- ' + c.key + ': ' + c.ask; }).join('\n'),
    '',
    'needs_review: true if anything is unclear, contradictory or looks mis-heard; otherwise false.',
    'review_reason: if needs_review is true, a few words saying what to check. Otherwise empty.',
  ].join('\n');
}

/** The exact JSON shape Claude must reply with. */
function proposalSchema() {
  const fieldProps = {};
  AI_KEYS.forEach(function (k) { fieldProps[k] = { type: 'string' }; });
  const properties = {
    action: { type: 'string', enum: ['new', 'update', 'unclear', 'cancel'] },
    target_row: { type: 'integer' },
    fields: { type: 'object', properties: fieldProps, required: AI_KEYS, additionalProperties: false },
    needs_review: { type: 'boolean' },
    review_reason: { type: 'string' },
    confirmation: { type: 'string' },
  };
  return { type: 'object', properties: properties, required: Object.keys(properties), additionalProperties: false };
}

/**
 * Checks Claude's proposal against the sheet and adds what the phone needs to show it:
 * the visit being changed, what changes, hours, distance, and which required details are missing.
 */
function finishProposal(answer, scope, ctx) {
  const visits = ctx.visits;
  const p = {
    action: answer.action,
    row: 0,
    fields: pick(answer.fields || {}, AI_KEYS),
    needs_review: !!answer.needs_review,
    review_reason: answer.review_reason || '',
    confirmation: answer.confirmation || '',
  };
  let target = null;

  // He asked to stop: nothing more to work out, and the phone won't listen again.
  if (p.action === 'cancel') {
    p.confirmation = 'Okay, I won\'t save that.';
    p.missing = [];
    p.docket = '';
    return p;
  }

  if (scope.type === 'row') { p.action = 'update'; answer.target_row = Number(scope.row); }

  if (p.action === 'update') {
    target = visits.filter(function (v) { return v.row === Number(answer.target_row); })[0];
    const allowed = target && (scope.type !== 'docket' || target.docket_number === scope.docket);
    if (!allowed) {
      target = null;
      p.action = 'unclear';
      p.confirmation = 'Sorry, I couldn\'t tell which visit you meant. Which docket and day was it?';
    } else {
      p.row = target.row;
      p.fingerprint = { recorded_at: target.recorded_at, transcript: target.transcript };
      p.target = { docket: target.docket_number, job_date: target.job_date, client: target.client, site: target.site };
    }
  }

  if (p.action === 'new') {
    if (scope.type === 'docket') p.fields.docket_number = scope.docket;
    const docket = ctx.dockets[p.fields.docket_number];
    if (docket) {
      if (!p.fields.client) p.fields.client = docket.client;
      if (!p.fields.site) p.fields.site = docket.site;
    }
  }

  if (p.action !== 'unclear') {
    addToConfirmation(p, applyDistance(p.fields, ctx.home, target));
  }
  if (target) {
    p.changes = AI_COLUMNS.filter(function (c) { return String(target[c.key] || '') !== String(p.fields[c.key] || ''); })
      .map(function (c) { return { key: c.key, label: c.header, from: String(target[c.key] || ''), to: String(p.fields[c.key] || '') }; });
  }

  p.docket = p.fields.docket_number || '';
  p.hours = hoursBetween(p.fields.arrived, p.fields.departed);
  p.missing = AI_COLUMNS.filter(function (c) { return c.required && !String(p.fields[c.key] || '').trim(); })
    .map(function (c) { return c.header; });
  return p;
}

/** Puts a sentence into the read-back, just before "Shall I save it?". */
function addToConfirmation(p, sentence) {
  if (!sentence) return;
  const question = /\s*Shall I save it\?\s*$/i;
  p.confirmation = question.test(p.confirmation)
    ? p.confirmation.replace(question, ' ' + sentence + ' Shall I save it?')
    : (p.confirmation + ' ' + sentence).trim();
}

/** "Today is 2026-11-15 (Sunday) and the time now is 14:05." */
function todayLine() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  return 'Today is ' + Utilities.formatDate(now, tz, 'yyyy-MM-dd (EEEE)') +
         ' and the time now is ' + Utilities.formatDate(now, tz, 'HH:mm') + '.';
}

/** Sends one request to Claude and returns the JSON object it replies with. */
function callClaude(system, message, schema) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY is not set in Script properties');

  const request = {
    model: MODEL,
    max_tokens: 4000,
    system: system,
    messages: [{ role: 'user', content: message }],
    output_config: {
      effort: 'low', // quick extraction: low effort is fast and cheap
      format: { type: 'json_schema', schema: schema },
    },
    fallbacks: 'default', // if this model declines a request, the API retries on a suitable fallback
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2026-10-21',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
    },
    payload: JSON.stringify(request),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  const body = JSON.parse(response.getContentText());
  if (code !== 200) {
    throw new Error('API error ' + code + ': ' + (body.error && body.error.message || response.getContentText()));
  }
  if (body.stop_reason === 'refusal') throw new Error('Claude declined to process this note');
  if (body.stop_reason === 'max_tokens') throw new Error('Reply was cut off');

  // The reply can include "thinking" blocks first; the answer is the text block.
  const textBlock = (body.content || []).filter(function (b) { return b.type === 'text'; })[0];
  if (!textBlock) throw new Error('No answer in the reply');
  return JSON.parse(textBlock.text);
}


/* ═════════════════════════════  TESTING  ═════════════════════════════ */

/**
 * Run these from the editor (choose one in the dropdown, then Run) and open the Execution log.
 * Neither saves anything.
 */
function testPropose() {
  const sample = 'Docket 4521, got to the Henderson place at 12 Mill Road at quarter past nine, ' +
    'set out four boundary pegs along the north fence, left about eleven thirty.';
  Logger.log(JSON.stringify(proposeEntry([sample], { type: 'any' }, null), null, 2));
}

function testDistance() {
  const home = readSettings().home_address;
  if (!home) return Logger.log('Put the home address in the Settings tab first.');
  Logger.log('Home: ' + home);
  const site = (readDockets().filter(function (d) { return d.site; })[0] || {}).site;
  if (!site) return Logger.log('Add a docket with a site address first (or record a visit).');
  Logger.log('Home → ' + site + ' → home: ' + routeKm([home, site, home]) + ' km');
}


/* ═════════════════════════════  HELPERS  ═════════════════════════════ */

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('No "' + name + '" tab yet. Run setupSheet() in the Apps Script editor first.');
  return sheet;
}

/**
 * Every row of the Visits tab as a plain object keyed by the COLUMNS keys, plus `row`
 * (the sheet row number). Dates come back as YYYY-MM-DD and times as HH:mm.
 */
function readVisits() {
  const sheet = getSheet(VISITS_TAB);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const range = sheet.getRange(2, 1, last - 1, COLUMNS.length);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  return values.map(function (r, i) {
    const visit = { row: i + 2 };
    COLUMNS.forEach(function (c, j) {
      const v = r[j];
      if (c.key === 'reviewed') visit.reviewed = v === true;
      else if (v instanceof Date && c.key === 'job_date') visit[c.key] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      else if (v instanceof Date && c.key === 'recorded_at') visit[c.key] = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm');
      else if (v instanceof Date) visit[c.key] = toHHmm(display[i][j]); // a time of day
      else visit[c.key] = String(v).trim();
    });
    return visit;
  }).filter(function (v) { return v.recorded_at || v.docket_number || v.transcript; });
}

/** Every row of the Dockets tab: { docket_number, client, site, row }. */
function readDockets() {
  const sheet = getSheet(DOCKETS_TAB);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 3).getDisplayValues().map(function (r, i) {
    return { docket_number: r[0].trim(), client: r[1].trim(), site: r[2].trim(), row: i + 2 };
  }).filter(function (d) { return d.docket_number; });
}

/** The Settings tab as { home_address: '…' }. */
function readSettings() {
  const sheet = getSheet(SETTINGS_TAB);
  const last = sheet.getLastRow();
  const rows = last >= 2 ? sheet.getRange(2, 1, last - 1, 2).getDisplayValues() : [];
  const out = {};
  SETTINGS.forEach(function (s) {
    const row = rows.filter(function (r) { return r[0].trim() === s.label; })[0];
    out[s.key] = row ? row[1].trim() : '';
  });
  return out;
}

function writeSetting(sheet, key, value) {
  const label = SETTINGS.filter(function (s) { return s.key === key; })[0].label;
  const last = sheet.getLastRow();
  const labels = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return r[0].trim(); }) : [];
  const i = labels.indexOf(label);
  const row = i === -1 ? Math.max(last, 1) + 1 : i + 2;
  sheet.getRange(row, 1, 1, 2).setValues([[label, value]]);
}

/**
 * Finds a visit by its row number, double-checking it's the same visit (same Recorded at
 * and transcript). If rows were sorted or deleted in the sheet, it's looked for elsewhere.
 */
function findVisit(visits, row, fingerprint) {
  fingerprint = fingerprint || {};
  const same = function (v) { return v.recorded_at === fingerprint.recorded_at && v.transcript === fingerprint.transcript; };
  const atRow = visits.filter(function (v) { return v.row === Number(row); })[0];
  if (atRow && same(atRow)) return atRow;
  const moved = visits.filter(same)[0];
  if (moved) return moved;
  throw new Error('This visit has moved or been deleted in the sheet. Go back, tap Refresh and try again.');
}

/** Just the listed keys of an object, as strings. */
function pick(obj, keys) {
  const out = {};
  keys.forEach(function (k) { out[k] = obj[k] === undefined || obj[k] === null ? '' : String(obj[k]); });
  return out;
}

function cleanNotes(notes) {
  return (notes || []).map(function (n) { return String(n || '').trim(); }).filter(Boolean);
}

/** "9:15:00", "09:15" or "9:15:00 PM" → "09:15" / "21:15". */
function toHHmm(text) {
  const m = /(\d{1,2}):(\d{2})/.exec(String(text));
  if (!m) return String(text);
  let h = Number(m[1]);
  if (/pm/i.test(text) && h < 12) h += 12;
  if (/am/i.test(text) && h === 12) h = 0;
  return ('0' + h).slice(-2) + ':' + m[2];
}

/** 1-based column number for a key in COLUMNS (the Visits tab). */
function colIndex(key) {
  for (let i = 0; i < COLUMNS.length; i++) if (COLUMNS[i].key === key) return i + 1;
  throw new Error('Unknown column key: ' + key);
}

/** "09:15" and "11:30" → 2.25. Empty if either time is missing. Handles past-midnight jobs. */
function hoursBetween(start, end) {
  const toMinutes = function (t) {
    const m = /^(\d{1,2}):(\d{2})$/.exec(String(t || '').trim());
    return m ? Number(m[1]) * 60 + Number(m[2]) : null;
  };
  const a = toMinutes(start), b = toMinutes(end);
  if (a === null || b === null) return '';
  let diff = b - a;
  if (diff < 0) diff += 24 * 60;
  return Math.round(diff / 60 * 100) / 100;
}

function unique(list) {
  return list.filter(function (x, i) { return list.indexOf(x) === i; });
}
