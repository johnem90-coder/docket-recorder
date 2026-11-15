# Docket Recorder

A phone app for logging survey jobs by voice. You dictate a note, Claude sorts it into columns (docket #, date, client, site, arrival and departure times, work done and more), and it's saved as a new row in a Google Sheet.

**How it's organised:** the Google Sheet has three tabs.
- **Visits:** one row per visit to a job, the main record. All the visits with the same Docket # make up one *docket*.
- **Dockets:** one row per docket, with its client and **site address**. Added automatically the first time a docket number is used. New visits to that docket default to this address.
- **Settings:** the **home address** that every trip starts and ends at, unless you say otherwise.

**Driving distances** are worked out with Google Maps for each visit: start address → site → end address. These default to home → the docket's address → home. If you say *"came straight from the Olsen job"* or *"went on to the Patel job"*, that visit's start or end changes to that job's address. If you say the km yourself (*"it was 50 k"*), your number is used and marked **Entered** in the **Km from** column, instead of **Google Maps**. The read-back tells you the distance, e.g. *"Driving there and back from home is 44 kilometres."*

What's on the phone:
- **Dockets screen:** a card for every docket showing the docket #, client and address. You can search, and sort by latest visit, docket # or client. Notes saved without a docket number are grouped at the top so they get sorted out.
- **Docket screen:** tap a card to see its visits, newest first, with total visits, hours and km, and each visit's route (🚗 Home → 9 Rimu Road → Home · 44.2 km). **✏️ Edit docket client / address** changes the docket's own details. Tap a visit to see everything recorded, including what was said and a link to its row in the sheet.
- **Talking to it:** **🎤 New note or change** (main screen) and **🎤 Add a visit or change** (docket screen) open the same voice conversation. Say what you did, or what to change on a visit you've already recorded ("on Tuesday's Olsen visit I actually left at noon"). Claude works out whether it's a new visit or a change, and which visit. It then shows (and, with **🔊 Hands-free** on, reads out) what it will save, mentions any required details that are missing, and waits. Say **"yes"** to save, say a correction, or say **"cancel"**. Nothing is saved until you confirm.
- **Saying "same as last time" is fine.** Phrases like "same location as the previous visit" are looked up in the earlier visits, so the real address goes in the sheet, not the phrase.
- **Editing by hand:** **✏️ Edit this visit** lets the office change any field and tick **Checked by the office**. **🎤 Change by voice instead** opens the conversation for just that visit.
- **Changes are tracked:** a voice change adds a dated line to the visit's **Raw transcript**. If the visit had already been ticked as reviewed, it's unticked and marked "Changed after it was reviewed".
- **Required details** (the app tells you when one is missing) are set with `required: true` in the `COLUMNS` list at the top of `Code.gs`. They are: docket #, job date, client, site, arrived, departed, work done and travel km. Travel km is worked out automatically, so it's only "missing" if Google Maps can't find the address.

- `apps-script/Code.gs`: the back end, pasted into the Google Sheet's Apps Script. Edit the `COLUMNS` list at the top to change what gets recorded.
- `docs/index.html`: the phone app, published as a web page with GitHub Pages. It talks to the Apps Script back end using a passcode.

## Setup (about 15 minutes)

### 1. Get a Claude API key
1. Go to <https://console.anthropic.com>, sign in, and add a few dollars of credit under **Billing**. Each note costs well under 1¢.
2. Go to **API keys → Create key**, name it `docket-recorder`, and copy the key. You'll paste it in step 3.

### 2. Create the Sheet and the script
1. In Google Drive, choose **New → Google Sheets** and name it **Docket Log**.
2. In the Sheet's menu, choose **Extensions → Apps Script**. A code editor opens in a new tab.
3. Rename the project (top left, "Untitled project") to **Docket Recorder**.
4. **Code.gs:** select everything in the editor, delete it, and paste in the contents of `apps-script/Code.gs`.
5. Click 💾 **Save**.

### 3. Add the API key, a passcode, and check the timezone
1. Click ⚙️ **Project Settings** in the left bar.
2. Under **General settings**, make sure **Time zone** is your local one.
3. Scroll to **Script properties → Add script property**:
   - Property: `ANTHROPIC_API_KEY`
   - Value: the key from step 1
4. Add a second script property:
   - Property: `APP_PASSCODE`
   - Value: a passcode for the phone app, e.g. three random words and a number (`kowhai-ladder-river-42`). Anyone with the app link *and* this passcode can read and add dockets, so don't use something guessable.
5. Click **Save script properties**.

### 4. Set up the sheet and test Claude
1. Go back to the **Editor** (`< >` icon).
2. In the function dropdown at the top, pick **setupSheet** and click **Run**.
3. Google asks for permission. Click **Review permissions**, choose your account, then **Advanced → Go to Docket Recorder (unsafe) → Allow**. ("Unsafe" just means Google hasn't reviewed the app. It's your own script running in your own account.)
4. Back in the Sheet, there are now **Visits**, **Dockets** and **Settings** tabs with header rows. (If you set the app up before these tabs existed, running **setupSheet** renames the old visits tab and adds the new ones. Nothing is deleted.)
5. In the **Settings** tab, type your home address next to **Home address**. You can also do this later in the app under ⚙ **Settings**, which checks it with Google Maps.
6. Pick **testPropose** in the dropdown and click **Run**. The **Execution log** should show Claude's proposal for a sample note, including a `confirmation` sentence. Nothing is saved.
7. Optional: if Google Maps keeps picking the wrong town for addresses, set `MAPS_REGION` near the top of Code.gs to your two-letter country code (e.g. `'nz'`).

### 5. Open the back end to the phone app
1. Click **Deploy → New deployment**. Click ⚙️ next to "Select type" and choose **Web app**.
   - Description: `v1`
   - Execute as: **Me**
   - Who has access: **Anyone** (the passcode is what keeps it private)
2. Click **Deploy** and copy the **Web app URL** (it ends in `/exec`). Opening it in a browser just says "Docket Recorder is running".

### 6. Publish the phone app with GitHub Pages
1. Put this project on GitHub (a free account is enough). The repository has to be **public** for free GitHub Pages. Nothing secret is in it: the API key and passcode live only in Script properties.
2. On GitHub, open the repository's **Settings → Pages**. Under **Build and deployment**, choose **Deploy from a branch**, branch **main**, folder **/docs**, and **Save**.
3. After a minute, the page shows the site's address, like `https://your-name.github.io/docket-recorder/`.

### 7. Put it on the phone
1. Open the GitHub Pages address on the phone in **Chrome**.
2. On the **Connect** screen, paste the Web app URL from step 5 and type the passcode from step 3, then tap **Connect**. The phone remembers both. (⚙ **Settings → Change connection** opens this again.)
3. In Chrome, tap **⋮ → Add to Home screen**. It now opens like an app.
4. The first time you tap 🎤, Chrome asks to use the microphone: tap **Allow**.

## Using it
Tap **🎤 New note or change** (or open a docket and tap **🎤 Add a visit or change**). With Hands-free on it beeps and starts listening straight away; otherwise tap **🎤 Tap to talk**. It keeps listening through pauses. Say **"that's it"** when you've finished, or tap **Done**. (If the Tap to talk button is missing, tap the text box, use the 🎤 on the Gboard keyboard, then tap **Next**.) For example:
> "Docket 4521, Henderson place on Mill Road, got there quarter past nine, set four boundary pegs on the north fence, left eleven thirty, 42 k, need to come back for the east side. That's it."

With **🔊 Hands-free** on (the default), the app starts listening as soon as you tap 🎤, reads back what it understood, beeps, and listens for your answer: **"yes"** saves, a correction ("no, it was 40 k") updates it and reads it back again, and **"cancel"** throws it away, and **"read it again"** repeats it. With Hands-free off, the same check appears on screen with **Save**, **Answer / change** and **Discard** buttons. After saving, the app opens that docket with the visit highlighted. Rows with missing or unclear details are marked **Needs review** in the **Status** column, and the words you said are always kept in **Raw transcript**.

**Safety:** record before pulling away or after parking, with the phone in a mount. Don't type or tap while driving.

## Making changes
After editing **Code.gs**, the phone keeps using the old version until you redeploy: **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The URL stays the same. After editing **docs/index.html**, push it to GitHub; the phone picks it up within a minute or two (close and reopen the app).

Fixing a docket number, client or anything else directly in the sheet is fine. The app picks up the change the next time it's opened, or when you tap **↻ Refresh**.
