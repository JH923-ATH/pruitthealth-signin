# PruittHealth Construction Services · Jobsite Sign In
## Setup and deployment guide

This walks you through going from the four files in this folder to a live sign in system that workers open by scanning a QR code at the gate. Do the parts in order. Nothing here costs money.

You need a Google account (personal Gmail is fine) and a GitHub account.

---

## Part A. Create the Google Sheet

1. Open [sheets.new](https://sheets.new) in your browser. This creates a blank Sheet in your Google Drive.
2. Rename it something like **PruittHealth Jobsite Sign Ins**. The name only matters to you.
3. Leave it open. You will paste the script into it next.

You do not need to add any columns yourself. The script creates them.

---

## Part B. Paste the Apps Script

1. In the Sheet, go to **Extensions → Apps Script**. A new tab opens with a code editor.
2. Delete the placeholder `function myFunction() { ... }` in `Code.gs`.
3. Open `Code.gs` from this folder in a plain text editor, select all, copy, and paste the entire contents into the Apps Script editor.
4. Click the disk icon (or press Ctrl+S / Cmd+S) to save. Name the project **Jobsite Sign In** when prompted.

### Run setup once

1. At the top of the Apps Script editor, in the function dropdown next to the Run button, choose **setup**.
2. Click **Run**.
3. Google will ask for permissions the first time. Click **Review permissions**, pick your Google account, click **Advanced → Go to Jobsite Sign In (unsafe)** (this warning is normal for scripts you write yourself), then **Allow**.
4. Look at the execution log. You should see `Setup complete. Sheet ready.`
5. Go back to the Sheet. You should see a new tab called **SignIns** with a bold header row.

---

## Part C. Deploy the script as a web app

This gives you a public URL that the phone app posts to.

1. In the Apps Script editor, click **Deploy → New deployment**.
2. Click the gear icon next to **Select type** and choose **Web app**.
3. Fill in:
   * **Description**: `Jobsite sign in v1`
   * **Execute as**: **Me (your@email)**
   * **Who has access**: **Anyone**
4. Click **Deploy**.
5. Google may ask you to authorize again. Approve it.
6. Copy the **Web app URL**. It ends in `/exec`. This is the URL the phone app talks to.

Keep this URL. You will paste it into `index.html` in the next part.

> **Important for future updates.** When you change `Code.gs`, do NOT click **New deployment** again. Instead go to **Deploy → Manage deployments**, click the pencil icon on the existing deployment, and choose **Version → New version**. Deploy. This keeps the `/exec` URL the same so your QR codes never break. Creating a new deployment issues a new URL and orphans every QR placard at every jobsite.

---

## Part D. Wire up index.html and publish to GitHub Pages

1. Open `index.html` in a text editor.
2. Near the top of the `<script>` block find:
   ```
   const APP_URL = "PASTE_APPS_SCRIPT_URL_HERE";
   ```
   Replace the placeholder with the `/exec` URL you copied in Part C. Keep the quotes.
3. Save the file.

Now put it online for free:

1. Sign in to GitHub. Create a new **public** repository, for example `pruitthealth-signin`. Public is required for free Pages hosting; the security posture in `PROJECT_BRIEF.md` accounts for this.
2. Upload `index.html` to the repo. If you have a PruittHealth logo file (transparent PNG), also upload it as `logo.png` in the same folder.
3. In the repo, go to **Settings → Pages**.
4. Under **Source**, pick **Deploy from a branch**. Branch: `main`. Folder: `/ (root)`. Save.
5. Wait about a minute. GitHub shows the live URL at the top of the Pages screen, something like `https://yourname.github.io/pruitthealth-signin/`.

Open that URL on your phone. You should see the PruittHealth header and a jobsite dropdown. If you see a red banner saying setup is incomplete, you did not paste the `/exec` URL into `index.html`.

---

## Part E. Test loop

Do this from your phone on cell data, not the office wifi, so you know it works from a jobsite.

1. Open the Pages URL.
2. Pick any site.
3. Tap **Sign in**. Enter a test name, pick a trade, enter a company, check the safety box, tap **Sign in now**.
4. Open the Google Sheet. You should see a new row in the **SignIns** tab with an Eastern time timestamp and Safety Ack `v1 2026-06-12`.
5. Go back to the phone. From the site screen tap **Sign out**. You should see your test name in the roster. Tap it.
6. Check the Sheet again. A second row with Action `Out` should appear.

If any of that fails, see **Troubleshooting** at the bottom.

---

## Part F. Optional daily email log (7 PM Eastern)

By default the daily email is off. To turn it on:

1. Open `Code.gs` in the Apps Script editor.
2. Change `const DAILY_EMAIL = '';` to your email address, for example `const DAILY_EMAIL = 'you@pruitthealth.com';`.
3. Save.
4. Deploy the change as a new **version** of the existing deployment (see the warning in Part C).
5. In the Apps Script editor click the clock icon (Triggers) on the left, then **Add Trigger**.
6. Choose function `dailyEmail`, event source **Time driven**, type **Day timer**, time **7pm to 8pm**.
7. Save. Google may ask for one more permission (send mail as you). Approve.

You will get a nightly email listing every sign in and sign out grouped by site. Anyone who signed in but never signed out is flagged.

---

## Part G. QR code placards

Once the Pages URL is live, generate a QR code that points to it. Any free QR generator works. Print one placard per site (they all point to the same URL because workers pick their site in the app). Weatherproof laminate is worth the two dollars.

### Showing the QR code from your phone (staff only)

Superintendents and the PM can pull the sign in QR code up on their own phone and hold it out for the crew to scan, handy when a placard is missing or someone is signing in away from the gate.

1. Open the **staff link**: the normal Pages URL with `?staff` on the end, e.g. `https://jh923-ath.github.io/pruitthealth-signin/?staff`.
2. Bookmark it, or **Add to Home Screen** from that page; the icon keeps staff mode.
3. On the first screen, tap **Show sign in QR code**. The screen stays on while it's showing; tap **Done** to go back.

Workers use the plain link and the gate QR, so they never see the button. Only share the staff link with supers. It isn't a password: the QR it shows is the same one printed on the gate placards, so nothing private is behind it. It just keeps the button off workers' screens. The QR on screen points to the plain link, so anyone who scans it gets the normal worker view.

### Skipping the QR code for regulars

The QR code is just a convenient way to get the Pages URL onto a phone without typing it; anyone can also bookmark that URL directly or, better, use their phone's **Add to Home Screen** (iOS Safari or Android Chrome, share menu → Add to Home Screen). That creates a tap-to-open icon with the PHCS name and icon below, no browser address bar, feels like a real app. Good for PruittHealth employees and regular subs who sign in daily; the QR code stays the easier path for one-time visitors and delivery drivers.

This works because `manifest.json`, `apple-touch-icon.png`, `icon-192.png`, and `icon-512.png` are already in the repo and linked from `index.html`. The icons are the PruittHealth tree mark on a white tile with **PHCS** underneath, so they stand out on a home screen.

---

## Part H. Delivery capture, OCR, and notifications

When a driver picks Trade = Delivery on sign in, the app shows extra fields (material, PO or ticket, for-whom, notes) and requires a photo of the material for condition verification. A bill of lading photo is highly recommended but OPTIONAL. When the driver takes a BOL photo, the server OCR's it, extracts the PO number (and material description if the driver left it blank), and appends the raw OCR text to Notes. When the BOL photo is skipped (memory constrained phone, unscannable BOL), the driver types material and PO manually. Each photo has a Remove button under its preview so the driver can drop or retake either one before signing in. A notification email fires with photos inlined; two-column layout when both are present, single-column when the BOL was skipped (with a flag in the email that the BOL was not provided).

### Enable the Advanced Drive Service (one time)

This lets `scanBol_` in `Code.gs` OCR the bill of lading photo. It is a Google service, no external API key needed.

1. In the Apps Script editor, click the **Services** panel on the left (looks like a stack icon).
2. Click the **+ Add a service** button.
3. Scroll to **Drive API**, select it. Keep the default **Identifier** `Drive` and default **Version** `v2`.
4. Click **Add**. The service now appears in the left panel.
5. Save the project.

If you skip this step, delivery sign in still works but OCR is silently skipped (the row saves without a detected PO).

### One time configuration

1. Open `Code.gs` in the Apps Script editor.
2. Find the CONFIG block near the top and set:
   * `DELIVERY_ALERT_PRIMARY`: your email address, e.g. `'you@pruitthealth.com'`. Leaving this blank turns the delivery notification feature off entirely; deliveries still sign in and photos still save, no email goes out.
   * `SITE_SUPERS`: for each site, put the site super's email in quotes. Leave any site blank if no super is assigned yet. If a site has more than one super, comma separate them inside the same string, e.g. `"Carrollton ALF / SNF": "sam@x.com, alex@x.com",`. Site names must exactly match the `SITES` array in `index.html`.
3. Save and publish as a **New version** of the existing deployment (see the warning in Part C).

### Run setup again

1. In the Apps Script editor, in the function dropdown pick **setup**, click **Run**.
2. Google will ask for two more permissions the first time: **See, edit, create, and delete Google Drive files** and **Send email as you**. Approve both.
3. The execution log will show something like `Setup complete. Sheet ready. Photo folder: https://drive.google.com/...`. Click that link. You will see the new empty **Jobsite Deliveries** folder. Confirm the sharing icon on the folder shows "Anyone with the link can view".
4. If the sharing icon does not show anyone-with-link, you are on a Google Workspace account where an admin has restricted external sharing. Right click the folder, Share, and set it to anyone with the link manually. If that is blocked too, the notification email still works (photo is attached), but the Drive link inside the email will not open for outside recipients.

### Test the delivery flow

1. On your phone open the live Pages URL.
2. Pick any site, tap **Sign in**.
3. Enter a test name and company. Pick Trade = **Delivery**. The Delivery details card appears.
4. Tap **Scan bill of lading** and snap a picture of any BOL or shipping doc.
5. Tap **Photo of the material** and snap a picture of the load (a stack of anything works for the test).
6. Check the safety box, tap **Sign in now**. The button shows "Saving and scanning" for a few extra seconds while OCR runs.
7. Success screen appears. If OCR read a PO from the BOL you snapped, it flashes "Detected PO xxxx" under the check mark.
8. Within 30 seconds you should get an email at `DELIVERY_ALERT_PRIMARY` with both photos inline side by side. Subject includes the detected PO when found. Site super (if set) is CC'd.
9. Open the Google Sheet. The new row should have PO (if OCR found it), the material (if you typed it, or if OCR guessed it), Notes including the raw BOL scan text, and two Drive URLs in the Material Photo URL and BOL Photo URL columns.

### Notes on the delivery photos

* Photos compress to about 250 KB on the phone before upload. Fast on cell data.
* Material photo is REQUIRED. BOL photo is optional but strongly recommended.
* Each photo preview has a Remove button so a driver can drop or retake either one before signing in.
* If a phone can't handle both photos (very old Android with limited memory), the driver takes only the material photo and types the BOL details manually into Material and PO.
* The camera capture asks the browser to open the rear camera. On older iOS or in-app browsers this can fall back to a file picker; if that happens, the driver picks a photo from their camera roll.
* Material text is optional. If left blank AND a BOL photo is present, OCR tries to fill it. If OCR can't tell, or the BOL was skipped, the column stays blank and the reviewer reads the BOL from the photo or the typed fields.
* OCR failures do not block the sign in. Row still saves; PO/material stay blank; reviewer works from what's there.

### Troubleshooting delivery notifications and OCR

**No email arrived.** Check Executions in the Apps Script editor (left sidebar, clock icon). If you see a failure on `MailApp.sendEmail`, you have not granted the send-email permission yet. Re-run `setup` and approve.

**Photo URL in the Sheet returns "Sorry, you need permission".** Your account restricts anyone-with-link sharing. Open the Jobsite Deliveries folder in Drive, right click, Share, set anyone with the link to Viewer. If your organization blocks that entirely, the recipient can still see photos in the email attachment.

**OCR never fills PO even on clean BOLs.** Confirm the Advanced Drive Service is enabled: script editor → Services (left panel) → Drive should be listed. If it is not, follow the enable step above and redeploy as a New version.

**Delivery card does not appear.** You are running an older cached copy of `index.html`. Hard refresh the phone browser, or append `?v=X` to the URL to bust the cache.

**Sign in stalls on "Saving and scanning" for more than 30 seconds.** Cell signal or a very large BOL photo. Try again on stronger signal.

---

## On-site meeting types

The **On-site meeting sign in** asks attendees to pick a meeting type (OAC meeting, bid job walk, preconstruction / kickoff, subcontractor coordination, site walk, inspection, punch list walk, owner / facility meeting, visitor, or Other with a description). It lands in the **Meeting Type** column of the **Meetings** tab. To change the list, edit `MEETING_TYPES` near the top of `index.html`; nothing in Apps Script needs to change.

---

## Part I. Morning safety meeting (toolbox talk) attendance

No extra deployment needed; it uses the same `/exec` URL and the same `Code.gs` already deployed.

On the landing choice screen, workers can pick **Morning safety meeting** instead of a jobsite sign in. The first person to open it for a project that day (normally the superintendent, running the meeting before anyone signs off) picks the topic from a dropdown (edit the `SAFETY_TOPICS` list in `index.html` to change the presets; picking **Other** reveals a free text field) and types who led it. Everyone who opens it after that sees the topic and leader already filled in and locked, and only enters their own name, trade, and company. There's no login behind this: it works because the superintendent is normally first in practice, not because the app enforces who can set the topic.

Each attendee also checks the same site safety acknowledgement. One row is written per attendee to a new **SafetyMeetings** tab (created automatically on first submission, same as the **Meetings** tab).

Run it daily or weekly, whichever fits the site; the tool does not assume a cadence. It is independent of the per sign in safety checkbox in the regular jobsite flow, so a worker can show up in both tabs on the same day.

### Exporting a report

After a morning safety meeting, you can generate a printable PDF attendance report instead of reading the raw SafetyMeetings tab:

1. In the Apps Script editor, in the function dropdown pick **exportSafetyMeetingReports**, click **Run**.
2. It builds one PDF per project that had a safety meeting today, styled like the corporate PruittHealth letterhead (logo centered at the top, "CONSTRUCTION SERVICES" beneath, mint accent bar, brand green headings), listing the topic, who led it, and a table of attendees with their company, trade, and sign in time. PDFs land in a **Safety Meeting Reports** folder in your Drive (auto created the first time). The logo is embedded inside Code.gs itself, so there is nothing extra to upload. Running this from the editor uses whatever Code.gs is saved there, so after pasting a new Code.gs you can run it right away; redeploying is only needed for the phone app side.
3. To have the PDFs emailed too, set up the report email once (see **Report email** below). Each run sends one email with all of its PDFs attached.

### Report email

Report emails go to the address in a **Script Property** named `REPORT_EMAIL`, not to an address typed into Code.gs. That keeps the address out of the public GitHub repo, and it stays put when you paste in a newer Code.gs.

1. In the Apps Script editor, click **Project Settings** (gear icon on the left).
2. Scroll to **Script Properties** and click **Add script property**.
3. Property: `REPORT_EMAIL`. Value: the address to send to, e.g. your work email. For several people, comma separate them.
4. Click **Save script properties**.

No redeploy is needed; the next report run picks it up. If `REPORT_EMAIL` is empty, reports are only saved to Drive. Advanced: `SAFETY_REPORT_EMAIL` and `SIGNIN_REPORT_EMAIL` in the Code.gs CONFIG block override `REPORT_EMAIL` for just that report; leave them blank normally.

This only covers today's meetings, matching the normal "run it right after the briefing" workflow. For older dates, open the SafetyMeetings tab directly and use Sheets' own **File → Download** (PDF, CSV, Excel) after filtering or sorting to the rows you want.

Re-running `exportSafetyMeetingReports` for a project that already has a report today replaces the old PDF rather than creating duplicates.

### Weekly reports (automatic)

Two weekly reports run together on one schedule, each covering the **7 full days ending yesterday** with one PDF per project:

* **Safety meetings** (`exportWeeklySafetyMeetingReports`): a section for each meeting in date order. Saved to **Safety Meeting Reports**, named like `Brookhaven Safety Meetings 2026-10-21 to 2026-10-27.pdf`.
* **Jobsite sign ins** (`exportWeeklySignInReports`): a daily summary (people on site, hours, anyone who never signed out), manpower by company (worker days and hours), each day's sign ins with in/out times and hours, and the week's deliveries (driver, material, PO). Saved to **Jobsite Sign In Reports**, named like `Brookhaven Jobsite Sign Ins 2026-10-21 to 2026-10-27.pdf`.

Both reports arrive together in **one email** (subject like `Weekly jobsite reports, 2026-10-23 to 2026-10-29`) with every PDF attached and Drive links listed, sent to `REPORT_EMAIL` (see **Report email** above). A week with no activity still sends a short "no activity" email so you know it ran. If one report fails, the email says which one.

To have both run on their own every week:

1. In the Apps Script editor, pick **installWeeklyReportTrigger** in the function dropdown and click **Run** (once). Approve the extra permission if Google asks.
2. That's it. By default they run every **Friday between 5 and 6 AM Eastern**, so they're ready first thing. Each report covers last Friday through Thursday; that Friday morning itself is left out and lands in the next week's report, so nothing is missed or counted twice. If you run a weekly report by hand on some other day, it still covers the 7 days before that day.

To change the day or time, edit `WEEKLY_REPORT_WEEKDAY` (MONDAY to SUNDAY) and `WEEKLY_REPORT_HOUR` (0 to 23) in the Code.gs CONFIG block, save, and run **installWeeklyReportTrigger** again; it replaces the old schedule rather than adding a second one. To stop the weekly reports, open **Triggers** (clock icon on the left) and delete the `exportWeeklyReports` trigger. You can also run **exportWeeklySafetyMeetingReports** or **exportWeeklySignInReports** by hand any time for an on demand report.

Notes:

* Google runs scheduled triggers within the chosen hour, not at an exact minute. Projects with no activity that week get no PDF.
* Hours come from pairing each person's sign in with their next sign out the same day, so someone who steps out for lunch and signs back in shows as two lines. Anyone who never signed out shows **Not signed out** and adds no hours, so the hours total is only as good as people's sign outs.
* Delivery drivers are listed under Deliveries rather than counted as workers on site.
* If one report fails, the other still runs, and Google emails you about the failure.

---

## Troubleshooting

**Phone shows the setup incomplete banner.** You did not paste the `/exec` URL into `index.html`, or you pasted it without the quotes. Fix and re upload to GitHub.

**Phone says "Could not save".** Almost always cell signal at the gate. Move a few feet and retry. If it fails on strong signal, check that the deployment is set to **Who has access: Anyone**.

**Sign out roster is empty even though you just signed in.** You are looking at a different site than you signed in to. The roster is filtered by site.

**Row lands in the Sheet but the timestamp column is blank.** The `SignIns` tab was renamed or replaced. Rename it back to `SignIns` or run `setup` again.

**You changed Code.gs and now the phone gets "Could not save".** You created a new deployment instead of publishing a new version. Go to **Deploy → Manage deployments**, edit the old deployment, publish a new version, and put that URL back into `index.html` if it changed. From now on always use New version.

---

## Files in this folder

* `index.html` is the phone app
* `Code.gs` is the Apps Script backend
* `SETUP_GUIDE.md` is this file
* `PROJECT_BRIEF.md` covers background, decisions, and rules for anyone touching this later
