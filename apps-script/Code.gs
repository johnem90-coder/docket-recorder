/**
 * Docket Recorder: Google Apps Script backend.
 *
 * The surveyor talks to the app. Claude works out whether he's describing a NEW visit to a
 * job or a CHANGE to a visit already recorded, and proposes the record. The app reads the
 * proposal back to him (or shows it), he confirms or corrects it, and only then is it saved.
 * Each row of the "Dockets" tab is one visit; all rows with the same Docket # make a docket.
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
 *  (Keep docket_number, client, site, job_date, arrived and departed: the app uses them.)
 * ───────────────────────────────────────────────────────────────────────────────
 */

// Which Claude model reads the notes. Change this one line to switch models.
const MODEL = 'claude-sonnet-5-5';

// Name of the tab the visits are written to.
const SHEET_NAME = 'Dockets';

// The columns of the sheet, left to right. Each row is one visit to one docket.
// Columns WITH an `ask` are filled in by Claude; the others by this script.
const COLUMNS = [
  { key: 'recorded_at',   header: 'Recorded at' },
  { key: 'docket_number', header: 'Docket #', required: true,
    ask: 'The docket or job number exactly as spoken (e.g. "4521" or "J-2024-17"). Empty if not mentioned. Never make one up.' },
  { key: 'job_date',      header: 'Job date', required: true,
    ask: 'Date the work was done, as YYYY-MM-DD. Work out words like "today" or "yesterday" from the date you are given. If no date is mentioned, use today.' },
  { key: 'client',        header: 'Client', required: true,
    ask: 'Client or customer name. Empty if not mentioned.' },
  { key: 'site',          header: 'Site / address', required: true,
    ask: 'Site name, address, road or property description. Empty if not mentioned.' },
  { key: 'arrived',       header: 'Arrived', required: true,
    ask: 'Time he arrived on site, as 24-hour HH:mm (e.g. "quarter past nine" → "09:15"). Empty if not mentioned.' },
  { key: 'departed',      header: 'Departed', required: true,
    ask: 'Time he left the site, as 24-hour HH:mm. Empty if not mentioned.' },
  { key: 'hours',         header: 'Hours on site' },
  { key: 'work_done',     header: 'Work done', required: true,
    ask: 'A short, clear summary of the work done on site, in his words where possible. Use "; " between separate tasks.' },
  { key: 'equipment',     header: 'Equipment / materials',
    ask: 'Equipment used or materials used/left on site (e.g. "GPS rover; 4 pegs"). Empty if not mentioned.' },
  { key: 'travel_km',     header: 'Travel km', required: true,
    ask: 'Kilometres driven for this job, digits only (e.g. "42"). Empty if not mentioned.' },
  { key: 'follow_up',     header: 'Follow-up needed',
    ask: 'Anything that still needs doing, a return visit, or something to tell the office/client. Empty if none.' },
  { key: 'notes',         header: 'Other notes',
    ask: 'Anything else he said that does not fit the other fields. Empty if nothing.' },
  { key: 'transcript',    header: 'Raw transcript' },
  { key: 'status',        header: 'Status' },
  { key: 'reviewed',      header: 'Reviewed ✓' },
];

// The fields Claude fills in (every column that has an `ask`).
const AI_COLUMNS = COLUMNS.filter(function (c) { return c.ask; });
const AI_KEYS = AI_COLUMNS.map(function (c) { return c.key; });

// Columns that can be changed by hand on the phone's Edit screen: everything Claude fills in,
// plus the office's Reviewed tick. (Hours are re-worked out from the times on save.)
const EDITABLE_KEYS = AI_KEYS.concat(['reviewed']);


/* ═════════════════════════════  SHEET SETUP  ═════════════════════════════ */

/**
 * Run this once from the Apps Script editor (and again after changing COLUMNS).
 * Creates the "Dockets" tab if needed and writes a tidy header row.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);

  const headers = COLUMNS.map(function (c) { return c.header; });
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8eef7')
    .setWrap(true);
  sheet.setFrozenRows(1);

  // Keep docket numbers as text so leading zeros like "0042" survive.
  sheet.getRange(2, colIndex('docket_number'), sheet.getMaxRows() - 1, 1).setNumberFormat('@');

  // Sensible widths: wide for the long text columns, narrow for the rest.
  COLUMNS.forEach(function (c, i) {
    const wide = ['work_done', 'transcript', 'follow_up', 'notes', 'site'].indexOf(c.key) !== -1;
    sheet.setColumnWidth(i + 1, wide ? 260 : 110);
  });

  Logger.log('Sheet "%s" is ready with %s columns.', SHEET_NAME, headers.length);
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

/**
 * Everything the phone page needs: every visit (one per sheet row), oldest first.
 * The page groups the visits into dockets itself.
 */
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet();
  return {
    visits: readVisits(sheet),
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
  const visits = readVisits(getSheet());

  const message = [
    todayLine(),
    scopeContext(scope, visits),
    previous ? 'Your last proposal, which he is replying to:\n' + JSON.stringify({
      action: previous.action, target_row: previous.row || 0, fields: previous.fields, you_said: previous.confirmation,
    }, null, 2) : '',
    'What he has said so far, in order:\n' + notes.map(function (n, i) { return (i + 1) + '. """' + n + '"""'; }).join('\n'),
  ].filter(Boolean).join('\n\n');

  const answer = callClaude(buildProposalInstructions(scope), message, proposalSchema());
  return finishProposal(answer, scope, visits);
}

/** Saves a confirmed proposal. With no proposal (Claude couldn't be reached), saves the raw note. */
function commitEntry(proposal, notes, scope) {
  notes = cleanNotes(notes);
  scope = scope || { type: 'any' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet();
    if (!proposal) {
      if (scope.type === 'row') throw new Error('Couldn\'t check that change, so nothing was saved. You can edit the fields by hand.');
      return appendVisit(sheet, { docket_number: scope.docket || '' }, notes.join(' / '),
        ['Saved without being checked by Claude: fill in the details by hand']);
    }
    if (proposal.action === 'new') {
      const problems = proposal.needs_review && proposal.review_reason ? [proposal.review_reason] : [];
      return appendVisit(sheet, proposal.fields, notes.join(' / '), problems);
    }
    if (proposal.action === 'update') return applyChange(sheet, proposal, notes);
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
    const sheet = getSheet();
    const visit = findVisit(readVisits(sheet), row, fingerprint);
    const updated = Object.assign({}, visit);
    EDITABLE_KEYS.forEach(function (key) {
      if (changes[key] !== undefined) updated[key] = key === 'reviewed' ? changes[key] === true : String(changes[key]).trim();
    });
    updated.hours = hoursBetween(updated.arrived, updated.departed);
    writeCells(sheet, visit.row, updated, EDITABLE_KEYS.concat(['hours']));
    return { docket: updated.docket_number || '', row: visit.row };
  } finally {
    lock.releaseLock();
  }
}


/* ═════════════════════════════  SAVING  ═════════════════════════════ */

/** Adds a new visit as a new row. `problems` are reasons the office should check it. */
function appendVisit(sheet, fields, transcript, problems) {
  const visits = readVisits(sheet);
  const known = knownDockets(visits);
  fields = Object.assign({}, fields);

  // Fill in a missing client or site from earlier visits to the same docket.
  const previous = known[fields.docket_number];
  if (previous) {
    if (!fields.client) fields.client = previous.client;
    if (!fields.site) fields.site = previous.site;
  }

  const values = Object.assign({}, fields, {
    recorded_at: new Date(),
    hours: hoursBetween(fields.arrived, fields.departed),
    transcript: transcript,
    status: statusFor(fields, problems),
    reviewed: false,
  });
  sheet.appendRow(COLUMNS.map(function (c) {
    const v = values[c.key];
    return (v === undefined || v === null) ? '' : v;
  }));
  const row = sheet.getLastRow();
  sheet.getRange(row, colIndex('reviewed')).insertCheckboxes();
  sheet.getRange(row, colIndex('docket_number')).setNumberFormat('@').setValue(values.docket_number || '');
  return { action: 'new', docket: values.docket_number || '', row: row, status: values.status };
}

/** Applies a confirmed voice change to an existing visit's row. */
function applyChange(sheet, proposal, notes) {
  const visit = findVisit(readVisits(sheet), proposal.row, proposal.fingerprint);
  const updated = Object.assign({}, visit);
  AI_KEYS.forEach(function (k) {
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

  writeCells(sheet, visit.row, updated, AI_KEYS.concat(['hours', 'transcript', 'status', 'reviewed']));
  return { action: 'update', docket: updated.docket_number || '', row: visit.row, status: updated.status };
}

/** "OK", or "Needs review: …" listing anything the office should check. */
function statusFor(fields, problems) {
  problems = (problems || []).slice();
  if (!fields.docket_number) problems.push('No docket number');
  if (!fields.arrived || !fields.departed) problems.push('Arrival or departure time missing');
  return problems.length ? 'Needs review: ' + unique(problems).join('; ') : 'OK';
}

/** Writes the given keys of `values` into one row of the sheet. */
function writeCells(sheet, row, values, keys) {
  keys.forEach(function (key) {
    const cell = sheet.getRange(row, colIndex(key));
    if (key === 'docket_number') cell.setNumberFormat('@');
    if (key === 'reviewed') cell.insertCheckboxes();
    cell.setValue(values[key] === undefined || values[key] === null ? '' : values[key]);
  });
}


/* ═════════════════════════════  CLAUDE  ═════════════════════════════ */

/** What Claude is told about the dockets and visits, depending on where he started talking. */
function scopeContext(scope, visits) {
  if (scope.type === 'row') {
    const v = visits.filter(function (x) { return x.row === Number(scope.row); })[0];
    if (!v) throw new Error('That visit couldn\'t be found. It may have been changed in the sheet.');
    return 'He is changing this visit (row ' + v.row + '):\n' + JSON.stringify(pick(v, AI_KEYS), null, 2) +
      (v.docket_number ? '\n\nOther visits to docket ' + v.docket_number + ':\n' +
        (visitLines(visits.filter(function (x) { return x.docket_number === v.docket_number && x.row !== v.row; }), 10) || 'none') : '');
  }
  if (scope.type === 'docket') {
    const own = visits.filter(function (x) { return x.docket_number === scope.docket; });
    const k = knownDockets(visits)[scope.docket] || {};
    return 'He is on the page for docket ' + scope.docket + ' (client: ' + (k.client || 'unknown') + ', site: ' +
      (k.site || 'unknown') + ').\nIts visits (newest first):\n' + (visitLines(own, 30) || 'none yet');
  }
  const known = knownDockets(visits);
  const dockets = Object.keys(known).slice(-150).map(function (d) {
    return d + ' | ' + (known[d].client || '-') + ' | ' + (known[d].site || '-');
  });
  return 'Existing dockets (number | client | latest site):\n' + (dockets.join('\n') || 'none yet') +
    '\n\nMost recently recorded visits (newest first):\n' + (visitLines(visits, 40) || 'none yet');
}

/** One line per visit, newest first, for Claude to pick from. */
function visitLines(visits, max) {
  return visits.slice()
    .sort(function (a, b) { return (b.job_date || '').localeCompare(a.job_date || '') || b.row - a.row; })
    .slice(0, max)
    .map(function (v) {
      return 'row ' + v.row + ' | docket ' + (v.docket_number || 'none') + ' | ' + (v.job_date || 'no date') + ' | ' +
        (v.client || '-') + ' | ' + (v.site || '-') + ' | ' + (v.arrived || '?') + '–' + (v.departed || '?') +
        ' | ' + (v.travel_km ? v.travel_km + ' km | ' : '') + 'work: ' + String(v.work_done || '').slice(0, 100) +
        (v.follow_up ? ' | follow-up: ' + String(v.follow_up).slice(0, 60) : '');
    }).join('\n');
}

/** The instructions Claude gets for proposing a record. */
function buildProposalInstructions(scope) {
  const required = AI_COLUMNS.filter(function (c) { return c.required; }).map(function (c) { return c.header; });
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
    'as yesterday", look up the real value in the visits you are given and use that.',
    '',
    'Choosing the action:',
    '- "new": a visit that isn\'t recorded yet. target_row is 0.',
    '- "update": correcting or adding to a visit that is already recorded, e.g. "on Tuesday\'s Olsen',
    '  visit I actually left at noon" or "add to the last Henderson visit that I left two pegs".',
    '  target_row is that visit\'s row number from the list. fields is the complete record after the',
    '  change: copy every field he didn\'t ask to change exactly as it is.',
    '- "unclear": you can\'t tell which visit he means, or whether it\'s new. Ask him in the confirmation.',
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
    'abbreviations, under 60 words. For a new visit, briefly say the docket, client, site, day, times,',
    'work and km. For a change, say which visit (docket and day) and only what is changing. Then name',
    'any required details that are still missing, and end with "Shall I save it?". If the action is',
    '"unclear", ask your question instead. Say times naturally ("half past eight", "eleven fifteen")',
    'and days relative to today ("today", "yesterday", "Tuesday the twelfth").',
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
    action: { type: 'string', enum: ['new', 'update', 'unclear'] },
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
 * the visit being changed, what changes, hours, and which required details are missing.
 */
function finishProposal(answer, scope, visits) {
  const p = {
    action: answer.action,
    row: 0,
    fields: pick(answer.fields || {}, AI_KEYS),
    needs_review: !!answer.needs_review,
    review_reason: answer.review_reason || '',
    confirmation: answer.confirmation || '',
  };

  if (scope.type === 'row') { p.action = 'update'; answer.target_row = Number(scope.row); }

  if (p.action === 'update') {
    const target = visits.filter(function (v) { return v.row === Number(answer.target_row); })[0];
    const allowed = target && (scope.type !== 'docket' || target.docket_number === scope.docket);
    if (!allowed) {
      p.action = 'unclear';
      p.confirmation = 'Sorry, I couldn\'t tell which visit you meant. Which docket and day was it?';
    } else {
      p.row = target.row;
      p.fingerprint = { recorded_at: target.recorded_at, transcript: target.transcript };
      p.target = { docket: target.docket_number, job_date: target.job_date, client: target.client, site: target.site };
      p.changes = AI_COLUMNS.filter(function (c) { return String(target[c.key] || '') !== String(p.fields[c.key] || ''); })
        .map(function (c) { return { label: c.header, from: String(target[c.key] || ''), to: String(p.fields[c.key] || '') }; });
    }
  }

  if (p.action === 'new') {
    if (scope.type === 'docket') p.fields.docket_number = scope.docket;
    const known = knownDockets(visits)[p.fields.docket_number];
    if (known) {
      if (!p.fields.client) p.fields.client = known.client;
      if (!p.fields.site) p.fields.site = known.site;
    }
  }

  p.docket = p.fields.docket_number || '';
  p.hours = hoursBetween(p.fields.arrived, p.fields.departed);
  p.missing = AI_COLUMNS.filter(function (c) { return c.required && !String(p.fields[c.key] || '').trim(); })
    .map(function (c) { return c.header; });
  return p;
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
 * Run this from the editor (choose testPropose, then Run), then open the Execution log.
 * It checks that Claude is working, using your real dockets, without saving anything.
 */
function testPropose() {
  const sample = 'Docket 4521, got to the Henderson place on Mill Road at quarter past nine, ' +
    'set out four boundary pegs along the north fence, left about eleven thirty, drove 42 k.';
  const proposal = proposeEntry([sample], { type: 'any' }, null);
  Logger.log(JSON.stringify(proposal, null, 2));
}


/* ═════════════════════════════  HELPERS  ═════════════════════════════ */

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('No "' + SHEET_NAME + '" tab yet. Run setupSheet() in the editor first.');
  return sheet;
}

/**
 * Every row of the sheet as a plain object keyed by the COLUMNS keys, plus `row`
 * (the sheet row number). Dates come back as YYYY-MM-DD and times as HH:mm.
 */
function readVisits(sheet) {
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

/** { "4521": { client, site }, ... } using the latest non-empty value for each docket. */
function knownDockets(visits) {
  const known = {};
  visits.forEach(function (v) {
    if (!v.docket_number) return;
    const k = known[v.docket_number] || (known[v.docket_number] = { client: '', site: '' });
    if (v.client) k.client = v.client;
    if (v.site) k.site = v.site;
  });
  return known;
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

/** 1-based column number for a key in COLUMNS. */
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
