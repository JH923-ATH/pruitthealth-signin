# PruittHealth Construction Services · Jobsite Sign In System
## Project brief for Claude Cowork handoff

### What this is

A digital replacement for paper jobsite sign in sheets across PruittHealth Construction Services jobsites. Subcontractors and employees scan a QR code at the gate, which opens a branded web app on their phone. They pick the jobsite from a dropdown, enter name, trade, and company, acknowledge the site safety requirements, and sign in. When leaving, they tap their name from a list of people currently on site to sign out. Every action lands as a timestamped row in a Google Sheet owned by John.

### Architecture (three pieces)

1. **index.html**: the entire app, a single file with no build step, hosted free on GitHub Pages. All configuration (backend URL, jobsite list, trade list, safety rules and version) lives in one clearly marked CONFIG block near the top of the script section.
2. **Code.gs**: a Google Apps Script bound to a Google Sheet in John's personal Google account, deployed as a web app. Receives sign ins via POST, serves the sign out roster via GET, and optionally emails a daily log. Timestamps are stamped server side in Eastern time so phone clocks cannot fake arrival times.
3. **The Google Sheet**: the permanent record. Columns: Timestamp, Date, Action, Name, Company, Trade, Site, Safety Ack, Material, PO/Ticket, For Trade, Notes, Photo URL. The last five columns are populated only on delivery sign ins.
4. **Google Drive folder "Jobsite Deliveries"**: photos of every delivery, auto created by setup(), set to anyone with the link can view. Each row's Photo URL column links here.

### Data flow

App (GitHub Pages) → POST JSON as text/plain → Apps Script doPost → appendRow to Sheet.
App sign out screen → GET ?site=X → Apps Script doGet → returns ONLY people currently on site today at that site.

### Security posture (do not weaken)

1. The GitHub repo is public (required for free Pages), so the Apps Script URL is visible in the source. Accepted risk: a stranger could submit a junk row, equivalent to scribbling on a paper sheet.
2. The doGet endpoint is deliberately locked down: it ignores date parameters, serves today only, returns only currently on site people (name, company, trade, time in), and never returns historical data, the Safety Ack column, or ANY delivery column (Material, PO/Ticket, For Trade, Notes, Photo URL). Full history and delivery details exist only in the Sheet.
3. Timestamps are server side only.

### Key product decisions already made

1. QR codes for all sites point at one URL; subs pick their site in the app dropdown.
2. Trade list includes "PruittHealth Employee" as the first entry. Selecting it auto fills the Company field with "PruittHealth" and locks it, keeping company data consistent. Form order is name, then trade, then company, so the auto fill lands after the trade choice.
3. Safety acknowledgment: five requirements shown at sign in with a required checkbox. Each sign in row stores SAFETY_VERSION (currently "v1 2026-06-12"). When wording changes, bump the version so records show which text each person agreed to. Wording still needs PruittHealth EHS review before launch.
4. Morning safety meeting (toolbox talk) attendance: a separate landing choice from the regular jobsite sign in. Documents that a subcontractor attended and acknowledged that day's (or week's) toolbox talk, independent of the per sign in safety checkbox above, since they record different things: site access vs. a specific briefing. Topic and leader are set once per project per day: whoever reaches the form first (normally the superintendent, running the meeting before anyone else signs off) picks the topic from a preset SAFETY_TOPICS dropdown in index.html (edit that list anytime; picking "Other" reveals a free text field) and types who's leading it. Everyone after sees those two fields locked to whatever was already set and only fills in their own name, trade, and company; the app checks via a GET ?meetingTopic=<project> lookup against today's SafetyMeetings rows (getTodaysSafetyMeetingTopic_ in Code.gs) before deciding which form to show. There is no login, so "first person through" is an honor system assumption like everywhere else in this app, not something enforced server side. One submission per attendee lands in its own SafetyMeetings tab via handleSafetyMeeting_ in Code.gs. Works for either a daily or weekly cadence; nothing in the tool assumes one or the other. A companion function, exportSafetyMeetingReports (run manually from the Apps Script editor, never exposed through the public /exec URL), generates a printable signature-sheet-style PDF per project for that day's meeting(s) and saves it to a "Safety Meeting Reports" Drive folder, optionally emailing it if SAFETY_REPORT_EMAIL is set. Older dates stay accessible via Sheets' own File > Download on the SafetyMeetings tab.
5. Sign out is honor system via tap your name; the daily email flags anyone who never signed out.
6. Error banners are single instance and field specific (recently debugged; do not reintroduce stacking banners or generic combined validation messages).
7. Delivery capture: when Trade = Delivery is picked, extra fields appear plus up to two photos. Material photo (condition verification) is REQUIRED. BOL photo is OPTIONAL but strongly recommended; if the driver skips it, they type material and PO manually. Each photo preview has a Remove button. Photos compress client side to 900px longest side JPEG q0.72 (~250 KB) before upload, stored as Blobs (not base64 strings) during capture to survive tight memory budgets on older phones. Server saves photos to Drive with filenames suffixed _MAT and _BOL, then (only when BOL photo present) OCRs it via Advanced Drive Service (Drive.Files.insert with ocr:true) to extract PO and best-effort material description. Driver typed values always win over OCR guesses. Raw OCR text appended to Notes. Notification email inlines the photos; two column layout when both present, single column when BOL skipped (with a flag noting BOL was not provided). Subject line includes detected PO when found. Notification and OCR failures never block the sign in row from saving. Client POST uses fetch mode:'no-cors' to sidestep the Apps Script /exec cross origin redirect bug that surfaces as HTTP 404 on some mobile browsers; response body is not readable so the success screen no longer shows "Detected PO" (email + Sheet carry that info).

### Design system

1. Brand: PruittHealth Construction Services. Header shows a logo slot (img src="logo.png", hides gracefully if missing) beside a "PruittHealth" wordmark with "CONSTRUCTION SERVICES" beneath.
2. Colors: deep pine and leaf green family in CSS variables at the top of the stylesheet (--ph-dark #17402A, --ph-green #4E9D54, plus tints). These are approximations; official hex values from PruittHealth marketing should replace them when available. Changing the two main variables retheme the whole app.
3. Typeface: Nunito Sans from Google Fonts as a free stand in. A marked TYPEFACE SWAP POINT at the top of the stylesheet makes it a two line change once the actual pruitthealth.com font is identified.
4. Style: soft rounded cards, gentle shadows, big tap targets and strong contrast because this is used outdoors on phones, often with gloves.

### Current status

Working and debugged in local testing. NOT yet deployed by John. Deployment steps live in SETUP_GUIDE.md: paste Code.gs into a Google Sheet's Apps Script, run setup once, deploy as web app (execute as Me, access Anyone), paste the /exec URL into index.html's CONFIG, upload to a public GitHub repo with Pages enabled.

### Open items

1. Obtain official PruittHealth logo file (transparent PNG of the tree mark) and add to repo as logo.png.
2. Obtain official brand hex codes and typeface name from marketing; swap into the marked spots.
3. EHS review of the safety acknowledgment wording; bump SAFETY_VERSION on any change.
4. Deploy per SETUP_GUIDE.md and run the test loop: sign in, verify row in Sheet, sign out, verify roster empties.
5. Optional: enable the daily 7 PM log email trigger (Part E of the guide).
6. Printable QR gate placards, one per site, matching the app theme, once the live GitHub Pages URL exists.

### Rules for anyone modifying this project

1. Never break the CONFIG block contract in index.html: APP_URL, SITES, TRADES, SAFETY_VERSION, SAFETY_RULES, SAFETY_TOPICS, PHOTO_MAX_DIM, PHOTO_JPEG_QUALITY must remain simple constants that a non developer can edit on GitHub. Same rule for Code.gs CONFIG: DELIVERY_ALERT_PRIMARY, SITE_SUPERS, DELIVERY_PHOTO_FOLDER_NAME, DAILY_EMAIL, SAFETY_REPORT_FOLDER_NAME, SAFETY_REPORT_EMAIL.
2. Never loosen the doGet lockdown in Code.gs. The ?site= roster lookup must never return Safety Ack or any delivery column, ever. The ?meetingTopic= lookup must stay equally narrow: only a project's topic and leader for today, never an attendee's name, company, trade, or Safety Ack. Any future doGet addition should follow the same rule: today only, and only the minimum fields the client actually needs, since the /exec URL is public.
3. POST must stay Content Type text/plain (avoids a CORS preflight that Apps Script cannot answer).
4. If Code.gs changes, redeploy as a NEW VERSION of the EXISTING deployment so the URL never changes (Deploy → Manage deployments → pencil → New version). A new deployment would orphan the QR codes.
5. Keep tap targets large and contrast high; this runs in direct sunlight.
6. Do not use dashes as punctuation in any user facing text or documents for John.

### Files in this package

1. index.html: the app (current, all fixes applied)
2. Code.gs: the backend script
3. SETUP_GUIDE.md: deployment walkthrough
4. PROJECT_BRIEF.md: this document

An earlier architecture (Microsoft Forms plus Power Automate into an Excel workbook, jobsite_signin.xlsx) was built, verified, and then retired in favor of the Google backend. It is not part of the active system.