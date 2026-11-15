# Docket Recorder: notes for Claude Code

## Who you're working with
The person working on this project runs the office side of a small surveying business and is
new to programming, Git and GitHub. Explain what you're doing in plain English, one step at a
time. Ask before installing anything, and never assume they know a technical term. After a
change, tell them exactly what they need to do outside this folder (in Google Apps Script, on
GitHub, on the phone) to make it live.

## What the app is
A surveyor records job visits by talking to his Android phone. Claude (through the API) turns
what he says into a proposed record, the app reads it back, and he confirms by voice. Records
go into a Google Sheet that the office reviews. Driving distances come from Google Maps.
`README.md` has the full setup and user guide.

## The two parts, and how each goes live
- `docs/index.html`: the whole phone app (one page: HTML, CSS and JavaScript, no build step).
  It's hosted by GitHub Pages from the `/docs` folder of the `main` branch. **Goes live when
  committed and pushed to GitHub**, within a minute or two. Phones may cache the old page for
  up to about 10 minutes.
- `apps-script/Code.gs`: the back end, a Google Apps Script attached to the Google Sheet. It
  is **not** deployed from here. To go live, they paste the whole file into the Apps Script
  editor, click Save, run `setupSheet` if the columns changed, and then go to **Deploy →
  Manage deployments → Edit (pencil) → Version: New version → Deploy**. Without the new
  version, the old code keeps running. Always spell these steps out after changing Code.gs.

The page talks to Apps Script with `fetch` POST requests to the web-app link (ending `/exec`),
sending `{ passcode, action, args }`. `doPost` in Code.gs checks the passcode and only runs
the functions listed in `API_ACTIONS`. A new server function the page calls must be added
there, and to the list in `server()` in docs/index.html.

## The Google Sheet
Three tabs: **Visits** (one row per visit), **Dockets** (one row per docket: client and site
address), **Settings** (home address). The Visits columns are defined by the `COLUMNS` list at
the top of Code.gs. Columns with an `ask` are filled in by Claude; `required: true` makes the
read-back mention them when missing.

**Add new columns at the end of `COLUMNS`** (after `reviewed`). The sheet already holds rows,
and inserting a column in the middle would put existing data under the wrong headings. Never
rename a column `key`: the code and existing rows depend on it.

## Secrets
The Claude API key (`ANTHROPIC_API_KEY`) and the app passcode (`APP_PASSCODE`) live only in
Apps Script → Project Settings → Script properties. Never put them in code, commit them, or
ask for them in chat. The repository is public.

## Working habits
- Make one change at a time, and commit with a clear message before and after.
- The phone app is used hands-free while driving, so voice flows must never need a tap
  mid-conversation, and anything spoken back should be short and plain.
- Keep the code commented in plain English. The person maintaining it reads the comments.
- Update `README.md` when behaviour or setup steps change.
