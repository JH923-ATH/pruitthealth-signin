/**
 * PruittHealth Construction Services · Jobsite Sign In
 * Backend: Google Apps Script bound to a Google Sheet.
 *
 * Contract (do not weaken):
 *   Sheet columns: Timestamp | Date | Action | Name | Company | Trade | Site | Safety Ack |
 *                  Material | PO/Ticket | For Trade | Notes |
 *                  Material Photo URL | BOL Photo URL
 *   Delivery columns (Material through BOL Photo URL) are populated only when
 *     Trade = "Delivery" on a sign in. Every other row leaves them blank.
 *   Deliveries capture up to two photos: one of the material for condition
 *     verification (Material Photo URL, REQUIRED) and one of the bill of lading
 *     for OCR + audit (BOL Photo URL, OPTIONAL). If the driver skips the BOL
 *     photo (memory constrained phone, or BOL not scannable), they can type
 *     material and PO manually. When BOL photo is present, server side OCR
 *     auto fills PO and, when the driver left the Material field blank, tries
 *     to infer material from the BOL text. Raw OCR text is appended to Notes
 *     so nothing is lost.
 *   Timestamps are stamped SERVER SIDE in Eastern time (phone clocks cannot fake arrival).
 *   POST payload is JSON sent as text/plain (avoids a CORS preflight Apps Script cannot answer).
 *   doGet is deliberately locked down: today only, currently on site only,
 *     returns only Name / Company / Trade / Timestamp, never Safety Ack, never any
 *     delivery column (Material, PO, For Trade, Notes, Material Photo URL,
 *     BOL Photo URL), ignores any date parameter.
 *
 *   Separately, index.html's on-site meeting flow (the "On-site meeting sign
 *     in" choice on the landing screen) posts { action: 'MeetingSignIn', ... }
 *     to this same web app, handled by handleMeetingSignIn_ into its own
 *     "Meetings" tab. Independent of the SignIns contract above.
 *
 *   Separately again, index.html's "Morning Safety Meeting" choice posts
 *     { action: 'SafetyMeeting', ... } to this same web app, handled by
 *     handleSafetyMeeting_ into its own "SafetyMeetings" tab. Documents that
 *     a subcontractor attended and acknowledged that day's (or week's)
 *     toolbox talk; independent of the SignIns contract and the per sign in
 *     safety checkbox, so a worker can appear in both.
 *
 * When redeploying after a code change, use:
 *   Deploy → Manage deployments → pencil icon on the existing deployment → New version
 * A brand new deployment would issue a new /exec URL and orphan the QR codes.
 */

/* ===================== CONFIG ======================
   TIMEZONE                    : Sheet is written in this timezone.
   SHEET_NAME                  : Tab in the bound Sheet. setup() creates it if missing.
   HEADERS                     : Column order. Must stay in sync with index.html payload.
   DAILY_EMAIL                 : Recipient for the optional 7 PM daily log. Blank to disable.
   DELIVERY_ALERT_PRIMARY      : Always CC'd on delivery arrivals. Blank turns the feature off.
   SITE_SUPERS                 : Per site super email(s), comma separated for multiple.
                                 Site names must match SITES in index.html exactly.
   DELIVERY_PHOTO_FOLDER_NAME  : Drive folder where delivery photos are stored.
                                 Auto created by setup() and set to anyone with link viewable.
   ================================================== */
const TIMEZONE   = 'America/New_York';
const SHEET_NAME = 'SignIns';
const HEADERS    = [
  'Timestamp', 'Date', 'Action', 'Name', 'Company', 'Trade', 'Site', 'Safety Ack',
  'Material', 'PO/Ticket', 'For Trade', 'Notes', 'Material Photo URL', 'BOL Photo URL'
];

// On-site meeting sign in (meeting-signin.html). Separate tab, separate
// contract from the jobsite SignIns sheet above; created automatically on
// first submission.
const MEETING_SHEET_NAME = 'Meetings';
const MEETING_HEADERS    = [
  'Timestamp', 'Meeting Type', 'Project', 'Meeting Date',
  'Name', 'Company', 'Role', 'Email', 'Phone'
];

// Morning (or weekly) safety meeting / toolbox talk attendance. Separate tab,
// separate contract from both SignIns and Meetings above; one row per
// attendee, created automatically on first submission.
const SAFETY_MEETING_SHEET_NAME = 'SafetyMeetings';
const SAFETY_MEETING_HEADERS    = [
  'Timestamp', 'Project', 'Meeting Date', 'Topic', 'Leader',
  'Name', 'Company', 'Trade', 'Safety Ack'
];
// Report PDF exports. exportSafetyMeetingReports (today's safety meetings) is
// run by hand from the Apps Script editor. exportWeeklyReports (last 7 days of
// safety meetings AND jobsite sign ins) runs on its own once
// installWeeklyReportTrigger has been run once. Emails are optional; blank
// means the PDFs are only saved to Drive, no email sent.
const SAFETY_REPORT_FOLDER_NAME = 'Safety Meeting Reports';
const SAFETY_REPORT_EMAIL = '';   // e.g. 'you@pruitthealth.com'. Blank = no email.
const SIGNIN_REPORT_FOLDER_NAME = 'Jobsite Sign In Reports';
const SIGNIN_REPORT_EMAIL = '';   // e.g. 'you@pruitthealth.com'. Blank = no email.
// When the weekly reports run. Day is MONDAY..SUNDAY; hour is 0-23 Eastern.
// After changing either, run installWeeklyReportTrigger again.
const WEEKLY_REPORT_WEEKDAY = 'FRIDAY';
const WEEKLY_REPORT_HOUR    = 16;   // 4 PM
const DAILY_EMAIL = '';   // e.g. 'you@pruitthealth.com'

const DELIVERY_ALERT_PRIMARY = '';   // e.g. 'you@pruitthealth.com'. Blank = feature off.
const SITE_SUPERS = {
  "Carrollton ALF / SNF": "",
  "Augusta Canopy":       "",
  "Christian City ALF":   "",
  "Christian City SNF":   "",
  "Creekside Dialysis":   "",
  "Moncks Corner":        "",
  "Brookhaven":           "",
};
const DELIVERY_PHOTO_FOLDER_NAME = 'Jobsite Deliveries';
/* ================== END CONFIG ===================== */


/**
 * Run this ONCE from the Apps Script editor after pasting this file.
 * Creates the SignIns tab if missing, writes the header row, forces plain text
 * on the Timestamp and Date columns, and creates or finds the Drive folder for
 * delivery photos with anyone with link viewable sharing.
 * Safe to run again; it will not clobber existing rows.
 */
function setup() {
  const sheet = getSheet_();
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const headerMatches = HEADERS.every((h, i) => firstRow[i] === h);
  if (!headerMatches) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  sheet.setColumnWidth(1, 170); // Timestamp
  sheet.setColumnWidth(2, 110); // Date
  sheet.setColumnWidth(3, 70);  // Action
  // Force Timestamp and Date columns to plain text so Google Sheets does not
  // silently auto parse them into Date objects. If it did, doGet would compare
  // a Date to today's yyyy-MM-dd string, always fail, and the sign out roster
  // would appear empty even for freshly signed in workers.
  sheet.getRange('A:B').setNumberFormat('@');

  // Create or find the Drive folder for delivery photos, and set it to
  // "anyone with the link can view" so alert recipients can open photo URLs
  // without needing Drive access. Files added later inherit this sharing.
  const folder = getPhotoFolder_();
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log('Could not set folder sharing (may be blocked by Workspace admin): ' + err);
  }
  Logger.log('Setup complete. Sheet ready. Photo folder: ' + folder.getUrl());
}


/**
 * POST handler. Body is JSON as text/plain.
 * Expected payload:
 *   { action: 'In' | 'Out', name, company, trade, site,
 *     ack?: SAFETY_VERSION,
 *     material?, po?, forTrade?, notes?,
 *     materialPhotoDataUrl?: 'data:image/jpeg;base64,...',
 *     bolPhotoDataUrl?:      'data:image/jpeg;base64,...' }
 * Delivery fields are only used when trade === 'Delivery' and action === 'In'.
 * BOTH photos are required on a delivery sign in. The BOL photo is OCR'd server
 * side; extracted PO and material (if driver left them blank) go into their
 * respective columns, and raw OCR text is appended to Notes. OCR failure is
 * logged but never blocks the sign in.
 * Returns: { ok: true, detectedPo?, detectedMaterial? } on success.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut_({ ok: false, error: 'empty body' });
    }
    const p = JSON.parse(e.postData.contents);
    const action = String(p.action || '').trim();
    if (action === 'MeetingSignIn') {
      return handleMeetingSignIn_(p);
    }
    if (action === 'SafetyMeeting') {
      return handleSafetyMeeting_(p);
    }
    if (action !== 'In' && action !== 'Out') {
      return jsonOut_({ ok: false, error: 'bad action' });
    }
    const name    = trim_(p.name);
    const company = trim_(p.company);
    const trade   = trim_(p.trade) || 'Not given';
    const site    = trim_(p.site);
    if (!name || !company || !site) {
      return jsonOut_({ ok: false, error: 'missing field' });
    }
    // Safety Ack is only stored on sign in. Sign out never records it.
    const safetyAck = action === 'In' ? trim_(p.ack) : '';

    const isDelivery = (action === 'In' && trade === 'Delivery');

    // Delivery detail fields default to blank on non delivery rows so column
    // positions stay stable in the Sheet.
    let material = '', po = '', forTrade = '', notes = '';
    let materialPhotoUrl = '', bolPhotoUrl = '';
    let materialPhotoBlob = null, bolPhotoBlob = null;
    let detectedPo = '', detectedMaterial = '', ocrText = '';

    if (isDelivery) {
      material = trim_(p.material);
      po       = trim_(p.po);
      forTrade = trim_(p.forTrade);
      notes    = trim_(p.notes);
      // Only the material photo is required. BOL photo is optional; if the
      // driver skips it (memory constrained phone, unscannable BOL), we skip
      // OCR and rely on the driver typed values.
      if (!p.materialPhotoDataUrl) return jsonOut_({ ok: false, error: 'material photo required' });

      try {
        const savedMat = savePhoto_(p.materialPhotoDataUrl, site, name, 'MAT');
        materialPhotoUrl = savedMat.url;
        materialPhotoBlob = savedMat.blob;

        if (p.bolPhotoDataUrl) {
          const savedBol = savePhoto_(p.bolPhotoDataUrl, site, name, 'BOL');
          bolPhotoUrl = savedBol.url;
          bolPhotoBlob = savedBol.blob;
        }
      } catch (err) {
        return jsonOut_({ ok: false, error: 'photo save failed: ' + (err && err.message || err) });
      }

      // OCR the BOL only if a BOL photo was provided. Failures are logged and
      // swallowed; the sign in still saves.
      if (bolPhotoBlob) {
        try {
          const scan = scanBol_(bolPhotoBlob);
          ocrText          = scan.rawText          || '';
          detectedPo       = scan.detectedPo       || '';
          detectedMaterial = scan.detectedMaterial || '';
          // Driver typed values always win over OCR guesses.
          if (!po       && detectedPo)       po = detectedPo;
          if (!material && detectedMaterial) material = detectedMaterial;
          // Append raw OCR text so nothing is lost even when regex misses.
          if (ocrText) {
            notes = notes
              ? (notes + '\n\n--- BOL scan ---\n' + ocrText)
              : ('BOL scan:\n' + ocrText);
          }
        } catch (err) {
          Logger.log('scanBol_ failed (sign in still saved): ' + err);
        }
      }
    }

    const now = new Date();
    const tsStr   = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const dateStr = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');

    getSheet_().appendRow([
      tsStr, dateStr, action, name, company, trade, site, safetyAck,
      material, po, forTrade, notes, materialPhotoUrl, bolPhotoUrl
    ]);

    // Fire notification after the row is safely written. Any email failure is
    // logged but does not fail the sign in, because the row already exists.
    if (isDelivery && DELIVERY_ALERT_PRIMARY) {
      try {
        notifyDelivery_({
          name: name, company: company, site: site, tsStr: tsStr,
          material: material, po: po, forTrade: forTrade, notes: notes,
          materialPhotoUrl: materialPhotoUrl, materialPhotoBlob: materialPhotoBlob,
          bolPhotoUrl: bolPhotoUrl, bolPhotoBlob: bolPhotoBlob,
          detectedPo: detectedPo
        });
      } catch (err) {
        Logger.log('notifyDelivery_ failed: ' + err);
      }
    }

    const out = { ok: true };
    if (detectedPo) out.detectedPo = detectedPo;
    if (detectedMaterial) out.detectedMaterial = detectedMaterial;
    return jsonOut_(out);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}


/**
 * GET handler. Two independent lookups, picked by which query parameter is
 * present:
 *
 *   ?site=<siteName>
 *     Sign out roster. Returns a JSON array of { Name, Company, Trade,
 *     Timestamp } for people currently on site at that site TODAY (signed in
 *     and not yet signed out).
 *     Locked down on purpose:
 *       - Only today. Any date parameter in the query is ignored.
 *       - Only people currently on site. Historical rows are never returned.
 *       - Never returns Safety Ack.
 *       - Never returns any delivery column (Material, PO, For Trade, Notes,
 *         Photo URL).
 *       - If site is missing or unknown, returns an empty array (not an error).
 *
 *   ?meetingTopic=<project>
 *     Today's safety meeting lookup, so index.html's Morning safety meeting
 *     form can tell whether a topic/leader has already been set for that
 *     project today (by whoever got there first, normally the
 *     superintendent) and lock those fields for everyone after. Returns
 *     { topic, leader } from the first matching SafetyMeetings row for that
 *     project today, or {} if none yet. Equally locked down: only today,
 *     only topic + leader, never an attendee's name, company, trade, or
 *     Safety Ack. See handleSafetyMeeting_ for how rows are written.
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.meetingTopic) {
      return jsonOut_(getTodaysSafetyMeetingTopic_(String(params.meetingTopic)));
    }

    const site = params.site ? String(params.site) : '';
    if (!site) return jsonOut_([]);

    const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return jsonOut_([]);
    const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();

    // Walk chronologically, keep a map of "who is currently in" keyed by name+company.
    // A later Out for the same key clears them from the map.
    const onSite = {};
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const rDate = toDateStr_(r[1]);      // yyyy-MM-dd (handles Date objects too)
      const rAction = String(r[2]);
      const rSite = String(r[6]);
      if (rDate !== today) continue;
      if (rSite !== site) continue;
      const key = String(r[3]).toLowerCase() + '|' + String(r[4]).toLowerCase();
      if (rAction === 'In') {
        onSite[key] = {
          Name: r[3],
          Company: r[4],
          Trade: r[5],
          Timestamp: toTimestampStr_(r[0])  // 'yyyy-MM-dd HH:mm:ss' in Eastern time
        };
      } else if (rAction === 'Out') {
        delete onSite[key];
      }
    }
    const out = Object.keys(onSite).map(k => onSite[k]);
    return jsonOut_(out);
  } catch (err) {
    return jsonOut_([]);
  }
}


/**
 * OPTIONAL. Daily 7 PM email of today's activity.
 * To enable: set DAILY_EMAIL above, then in the Apps Script editor open
 * Triggers, add a new time driven trigger for dailyEmail, Day timer, 7 PM.
 * Flags anyone who signed in and never signed out.
 */
function dailyEmail() {
  if (!DAILY_EMAIL) return;
  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) {
    MailApp.sendEmail(DAILY_EMAIL,
      'Jobsite sign in log ' + today,
      'No activity recorded today.');
    return;
  }
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const today_rows = rows.filter(r => toDateStr_(r[1]) === today);

  // Group by site, and within site by name+company, to find never signed out.
  const bySite = {};
  today_rows.forEach(r => {
    const site = String(r[6]);
    bySite[site] = bySite[site] || [];
    bySite[site].push(r);
  });

  let body = 'PruittHealth Construction Services\nJobsite sign in log for ' + today + '\n\n';
  const siteNames = Object.keys(bySite).sort();
  if (!siteNames.length) {
    body += 'No activity recorded today.\n';
  }
  siteNames.forEach(site => {
    body += '=== ' + site + ' ===\n';
    const site_rows = bySite[site];
    const status = {};
    site_rows.forEach(r => {
      const key = String(r[3]).toLowerCase() + '|' + String(r[4]).toLowerCase();
      status[key] = status[key] || { name: r[3], company: r[4], trade: r[5], inTs: '', outTs: '' };
      const tsStr = toTimestampStr_(r[0]);
      if (String(r[2]) === 'In')  status[key].inTs  = tsStr;
      if (String(r[2]) === 'Out') status[key].outTs = tsStr;
    });
    Object.keys(status).forEach(k => {
      const s = status[k];
      const flag = (s.inTs && !s.outTs) ? '  [NEVER SIGNED OUT]' : '';
      body += s.name + ' (' + s.company + ', ' + s.trade + ')  in ' +
              (s.inTs || 'n/a') + '  out ' + (s.outTs || 'n/a') + flag + '\n';
    });
    body += '\n';
  });
  MailApp.sendEmail(DAILY_EMAIL, 'Jobsite sign in log ' + today, body);
}


/* ---------- meeting sign in ---------- */

/**
 * Handles { action: 'MeetingSignIn', meetingType, project, date, name,
 * company, role, email, phone } from meeting-signin.html. Appends a row to
 * the Meetings tab, creating the tab and header row on first use. project is
 * what lets rows be told apart later; it's picked from the same project list
 * as the jobsite sign in (kept in sync by hand in meeting-signin.html, since
 * this page has no build step to share it with index.html). Timestamp is
 * stamped server side; Meeting Date is whatever the attendee picked in the
 * form. Unrelated to, and does not touch, the jobsite SignIns sheet or its
 * In/Out contract.
 * Returns { ok: true } or { ok: false, error }.
 */
function handleMeetingSignIn_(p) {
  const name    = trim_(p.name);
  const email   = trim_(p.email);
  const project = trim_(p.project);
  if (!name || !email || !project) {
    return jsonOut_({ ok: false, error: 'missing field' });
  }
  const meetingType = trim_(p.meetingType) || 'On-Site Meeting';
  const date         = trim_(p.date);
  const company      = trim_(p.company);
  const role         = trim_(p.role);
  const phone        = trim_(p.phone);

  const tsStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  getMeetingSheet_().appendRow([
    tsStr, meetingType, project, date, name, company, role, email, phone
  ]);

  return jsonOut_({ ok: true });
}

function getMeetingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(MEETING_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(MEETING_SHEET_NAME);
    sh.getRange(1, 1, 1, MEETING_HEADERS.length).setValues([MEETING_HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, MEETING_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}


/* ---------- safety meeting (toolbox talk) attendance ---------- */

/**
 * Handles { action: 'SafetyMeeting', project, date, topic, leader, name,
 * company, trade, ack } from index.html's "Morning Safety Meeting" choice.
 * Documents that a subcontractor attended and acknowledged that day's (or
 * week's) toolbox talk. Appends one row per attendee to the SafetyMeetings
 * tab, creating the tab and header row on first use. Independent of the
 * SignIns contract and the per sign in safety checkbox; a worker can appear
 * in both, since they document different things (site access vs. a specific
 * briefing).
 * Returns { ok: true } or { ok: false, error }.
 */
function handleSafetyMeeting_(p) {
  const project = trim_(p.project);
  const date    = trim_(p.date);
  const topic   = trim_(p.topic);
  const name    = trim_(p.name);
  const company = trim_(p.company);
  if (!project || !date || !topic || !name || !company) {
    return jsonOut_({ ok: false, error: 'missing field' });
  }
  const leader = trim_(p.leader);
  const trade  = trim_(p.trade) || 'Not given';
  const ack    = trim_(p.ack);

  const tsStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  getSafetyMeetingSheet_().appendRow([
    tsStr, project, date, topic, leader, name, company, trade, ack
  ]);

  return jsonOut_({ ok: true });
}

function getSafetyMeetingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SAFETY_MEETING_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SAFETY_MEETING_SHEET_NAME);
    sh.getRange(1, 1, 1, SAFETY_MEETING_HEADERS.length).setValues([SAFETY_MEETING_HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, SAFETY_MEETING_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

/**
 * Finds the first SafetyMeetings row for a project today and returns its
 * topic + leader, so index.html can lock those fields for everyone after
 * whoever set them up (normally the superintendent, who is in practice the
 * first person through). Returns {} if no meeting has been set up for that
 * project yet today. Called from doGet's ?meetingTopic= branch; intentionally
 * returns nothing beyond topic/leader (see doGet's doc comment).
 */
function getTodaysSafetyMeetingTopic_(project) {
  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  const sheet = getSafetyMeetingSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const rows = sheet.getRange(2, 1, last - 1, SAFETY_MEETING_HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (toTimestampStr_(r[0]).slice(0, 10) !== today) continue;
    if (String(r[1]) !== project) continue;
    return { topic: String(r[3]), leader: String(r[4]) };
  }
  return {};
}


/**
 * Run this from the Apps Script editor (function dropdown: pick
 * exportSafetyMeetingReports, click Run) any time after a morning safety
 * meeting to produce a printable, signature-sheet-style PDF documenting who
 * attended and acknowledged it. Not reachable through doGet / the public
 * /exec URL, so it does not weaken the doGet lockdown; it only runs when a
 * signed in Google user (you) triggers it from the script editor.
 *
 * Covers TODAY only. For a 7 day roll up see exportWeeklySafetyMeetingReports.
 * Older ranges: open the SafetyMeetings tab and use Sheets' File > Download.
 */
function exportSafetyMeetingReports() {
  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  writeSafetyMeetingReports_(today, today, 'Safety Meeting ' + today, "today's");
}

/**
 * What the weekly trigger runs: both weekly reports. Each runs even if the
 * other fails; any failure is rethrown at the end so Google's trigger failure
 * email still reaches you.
 */
function exportWeeklyReports() {
  const errors = [];
  [exportWeeklySafetyMeetingReports, exportWeeklySignInReports].forEach(fn => {
    try { fn(); } catch (err) {
      errors.push(fn.name + ': ' + (err && err.message || err));
      Logger.log(fn.name + ' failed: ' + err);
    }
  });
  if (errors.length) throw new Error('Weekly report errors: ' + errors.join(' | '));
}

/**
 * Weekly safety meeting roll up: one PDF per project covering the last 7 days
 * (today plus the 6 days before), with a section per meeting in date order.
 * With the default Friday 4 PM schedule the window is Saturday through Friday.
 * Can also be run by hand from the editor.
 */
function exportWeeklySafetyMeetingReports() {
  const w = lastSevenDays_();
  writeSafetyMeetingReports_(w.from, w.to, 'Safety Meetings ' + w.from + ' to ' + w.to, "this week's");
}

/**
 * Weekly jobsite sign in roll up: one PDF per site covering the last 7 days,
 * with a daily headcount/hours summary, manpower by company, each person's
 * in/out times per day (never signed out flagged), and the week's deliveries.
 * Can also be run by hand from the editor.
 */
function exportWeeklySignInReports() {
  const w = lastSevenDays_();
  writeSignInReports_(w.from, w.to, 'Jobsite Sign Ins ' + w.from + ' to ' + w.to);
}

/**
 * Run ONCE from the Apps Script editor to schedule exportWeeklyReports (both
 * weekly reports) every WEEKLY_REPORT_WEEKDAY around WEEKLY_REPORT_HOUR.
 * Safe to run again, e.g. after changing the day or hour: it removes the
 * previous weekly trigger first so you never end up with duplicates. To stop
 * the weekly reports, delete the trigger under Triggers (clock icon).
 */
function installWeeklyReportTrigger() {
  const day = ScriptApp.WeekDay[String(WEEKLY_REPORT_WEEKDAY).toUpperCase()];
  if (!day) throw new Error('WEEKLY_REPORT_WEEKDAY must be MONDAY..SUNDAY, got: ' + WEEKLY_REPORT_WEEKDAY);

  // Also clears a trigger from the earlier safety-meetings-only installer.
  const handlers = ['exportWeeklyReports', 'exportWeeklySafetyMeetingReports'];
  ScriptApp.getProjectTriggers()
    .filter(t => handlers.indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('exportWeeklyReports')
    .timeBased()
    .onWeekDay(day)
    .atHour(WEEKLY_REPORT_HOUR)
    .inTimezone(TIMEZONE)
    .create();
  Logger.log('Weekly reports scheduled: ' + WEEKLY_REPORT_WEEKDAY +
             ' around ' + WEEKLY_REPORT_HOUR + ':00 ' + TIMEZONE);
}

/** { from, to } as yyyy-MM-dd for today and the 6 days before, in TIMEZONE. */
function lastSevenDays_() {
  const now = new Date();
  return {
    from: Utilities.formatDate(new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000), TIMEZONE, 'yyyy-MM-dd'),
    to:   Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd')
  };
}

/**
 * Saves a report PDF into the named Drive folder (created if missing),
 * replacing any earlier file of the same name so re-runs don't pile up
 * duplicates, and emails it when an address is given.
 */
function saveReportPdf_(folderName, filename, pdfBlob, email, subject, text) {
  pdfBlob.setName(filename);
  const folder = getReportFolder_(folderName);
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);
  const file = folder.createFile(pdfBlob);
  Logger.log('Report saved: ' + file.getUrl());
  if (email) {
    MailApp.sendEmail(email, subject, text,
      { attachments: [pdfBlob], name: 'PruittHealth Jobsite Sign In' });
  }
}

/** Project name made safe for a Drive file name: "Carrollton ALF / SNF" -> "Carrollton ALF SNF". */
function safeName_(name) {
  return String(name).replace(/[^A-Za-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function getReportFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

/**
 * Shared by the daily and weekly safety meeting exports. Collects
 * SafetyMeetings rows whose submit date falls in [fromDate, toDate]
 * (yyyy-MM-dd, inclusive) and saves one PDF per project.
 */
function writeSafetyMeetingReports_(fromDate, toDate, fileLabel, emailWord) {
  const periodLabel = fromDate === toDate ? fromDate : fromDate + ' to ' + toDate;
  const sheet = getSafetyMeetingSheet_();
  const last = sheet.getLastRow();
  if (last < 2) { Logger.log('No safety meeting rows yet.'); return; }
  const rows = sheet.getRange(2, 1, last - 1, SAFETY_MEETING_HEADERS.length).getValues();

  // Timestamp may come back as a Date object or a string depending on how
  // Sheets auto parsed it, so normalize first. yyyy-MM-dd strings compare
  // correctly as plain strings.
  const byProject = {};
  rows.forEach(r => {
    const tsStr = toTimestampStr_(r[0]);
    const day = tsStr.slice(0, 10);
    if (day < fromDate || day > toDate) return;
    const project = String(r[1]);
    byProject[project] = byProject[project] || [];
    byProject[project].push({
      ts: tsStr, topic: String(r[3]), leader: String(r[4]),
      name: String(r[5]), company: String(r[6]), trade: String(r[7]), ack: String(r[8])
    });
  });

  const projects = Object.keys(byProject);
  if (!projects.length) { Logger.log('No safety meetings recorded for ' + periodLabel + '.'); return; }

  projects.forEach(project => {
    saveReportPdf_(SAFETY_REPORT_FOLDER_NAME,
      safeName_(project) + ' ' + fileLabel + '.pdf',
      buildSafetyMeetingPdf_(project, periodLabel, byProject[project]),
      SAFETY_REPORT_EMAIL,
      'Safety meeting report: ' + project + ', ' + periodLabel,
      'Attached: ' + emailWord + ' safety meeting attendance report for ' + project + '.');
  });
}

/**
 * Collects SignIns rows dated within [fromDate, toDate] (yyyy-MM-dd,
 * inclusive) and saves one PDF per site. Same Date-object-vs-string
 * normalization as dailyEmail.
 */
function writeSignInReports_(fromDate, toDate, fileLabel) {
  const periodLabel = fromDate === toDate ? fromDate : fromDate + ' to ' + toDate;
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) { Logger.log('No sign in rows yet.'); return; }
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();

  const bySite = {};
  rows.forEach(r => {
    const day = toDateStr_(r[1]);
    if (day < fromDate || day > toDate) return;
    const site = String(r[6]);
    bySite[site] = bySite[site] || [];
    bySite[site].push({
      ts: toTimestampStr_(r[0]), day: day, action: String(r[2]),
      name: String(r[3]), company: String(r[4]), trade: String(r[5]),
      material: String(r[8]), po: String(r[9]), forTrade: String(r[10])
    });
  });

  const sites = Object.keys(bySite);
  if (!sites.length) { Logger.log('No jobsite sign ins recorded for ' + periodLabel + '.'); return; }

  sites.forEach(site => {
    saveReportPdf_(SIGNIN_REPORT_FOLDER_NAME,
      safeName_(site) + ' ' + fileLabel + '.pdf',
      buildSignInPdf_(site, periodLabel, summarizeSignIns_(bySite[site])),
      SIGNIN_REPORT_EMAIL,
      'Jobsite sign in report: ' + site + ', ' + periodLabel,
      'Attached: the jobsite sign in report for ' + site + ', ' + periodLabel + '.');
  });
}

/**
 * Pairs each person's In and Out rows into visits, per day, in time order.
 * A person can have several visits a day (out for lunch, back in). An In
 * with no later Out that day is a visit with out = null (never signed out);
 * an Out with no earlier In that day is kept with in = null. Delivery sign
 * ins are listed separately rather than as visits, since drivers normally
 * never sign out. Hours assume in and out fall on the same day, which the
 * per-day pairing guarantees.
 */
function summarizeSignIns_(entries) {
  const sorted = entries.slice().sort((a, b) => a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0);
  const visits = [], deliveries = [], open = {};
  const minutes = ts => {
    const p = ts.slice(11, 19).split(':');
    return Number(p[0]) * 60 + Number(p[1]) + Number(p[2] || 0) / 60;
  };
  sorted.forEach(e => {
    if (e.trade === 'Delivery') {
      if (e.action === 'In') deliveries.push(e);
      return;
    }
    const key = e.day + '|' + e.name.toLowerCase() + '|' + e.company.toLowerCase();
    if (e.action === 'In') {
      const v = { day: e.day, name: e.name, company: e.company, trade: e.trade, in: e.ts, out: null, hours: null };
      visits.push(v);
      open[key] = v;
    } else if (e.action === 'Out') {
      const v = open[key];
      if (v) {
        v.out = e.ts;
        v.hours = Math.max(0, minutes(e.ts) - minutes(v.in)) / 60;
        delete open[key];
      } else {
        visits.push({ day: e.day, name: e.name, company: e.company, trade: e.trade, in: null, out: e.ts, hours: null });
      }
    }
  });
  return { visits: visits, deliveries: deliveries };
}

// Brand colors sampled from the corporate PruittHealth logo; keep in sync
// with --ph-dark / --ph-mint in index.html.
const PH_DARK_HEX = '#0B4F30';
const PH_MINT_HEX = '#8ED6A8';

// Corporate PruittHealth logo (tree mark + wordmark, no facility name),
// cropped from the corporate Letterhead Template.docx. Embedded as base64 so
// the report needs no separate Drive upload. 520x109 PNG.
const PRUITT_LOGO_PNG_BASE64 = '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';
const PRUITT_LOGO_W = 520;
const PRUITT_LOGO_H = 109;

/**
 * Starts a report Doc with the PruittHealth letterhead: centered logo,
 * "CONSTRUCTION SERVICES", a thin mint bar, then the title and the
 * project / period lines. The corporate template's facility address footer
 * is left out since these reports aren't tied to one facility. Returns
 * { doc, body }; pass both to finishLetterheadPdf_ when done.
 */
function newLetterheadDoc_(title, project, periodLabel) {
  const doc = DocumentApp.create('Report_temp_' + Date.now());
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  const logoBlob = Utilities.newBlob(
    Utilities.base64Decode(PRUITT_LOGO_PNG_BASE64), 'image/png', 'logo.png');
  const firstPara = body.getParagraphs()[0];
  const logo = firstPara.appendInlineImage(logoBlob);
  const logoWidth = 220;
  logo.setWidth(logoWidth).setHeight(Math.round(logoWidth * PRUITT_LOGO_H / PRUITT_LOGO_W));
  firstPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const division = body.appendParagraph('CONSTRUCTION SERVICES');
  division.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  division.editAsText().setBold(true).setFontSize(9).setForegroundColor(PH_DARK_HEX);

  const bar = body.appendTable([['']]);
  bar.setBorderWidth(0);
  const barCell = bar.getRow(0).getCell(0);
  barCell.setBackgroundColor(PH_MINT_HEX).setPaddingTop(0).setPaddingBottom(0);
  barCell.editAsText().setFontSize(2);

  const titlePara = body.appendParagraph(title);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titlePara.editAsText().setBold(true).setFontSize(18).setForegroundColor(PH_DARK_HEX);

  const multiDay = periodLabel.indexOf(' to ') !== -1;
  body.appendParagraph('Project: ' + project).editAsText().setFontSize(11);
  body.appendParagraph((multiDay ? 'Week: ' : 'Date: ') + periodLabel).editAsText().setFontSize(11);
  return { doc: doc, body: body };
}

function appendSectionHeading_(body, text) {
  body.appendParagraph('');
  body.appendParagraph(text).editAsText().setBold(true).setFontSize(12).setForegroundColor(PH_DARK_HEX);
}

/** Appends a table whose first row is the brand-green header row. */
function appendBrandTable_(body, rows) {
  const table = body.appendTable(rows);
  table.setBorderColor('#D9E5D8');
  const header = table.getRow(0);
  for (let c = 0; c < header.getNumCells(); c++) {
    const cell = header.getCell(c);
    cell.setBackgroundColor(PH_DARK_HEX);
    cell.editAsText().setBold(true).setForegroundColor('#FFFFFF');
  }
  return table;
}

/**
 * Adds the footer, exports the Doc as a PDF blob, then trashes the temp Doc
 * (same create-then-trash pattern as scanBol_'s OCR step) so Drive stays clean.
 */
function finishLetterheadPdf_(doc, body) {
  const footer = body.appendParagraph('PruittHealth Construction Services  ·  PruittHealth.com');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(18);
  footer.editAsText().setFontSize(8).setForegroundColor(PH_DARK_HEX);

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf');
  file.setTrashed(true);
  return pdfBlob;
}

/**
 * Printable, signature-sheet-style PDF for one project's safety meeting(s)
 * over periodLabel (a single date, or "from to to" for the weekly roll up):
 * one section per meeting (same date + topic + leader), in date order, each
 * with an attendee table.
 */
function buildSafetyMeetingPdf_(project, periodLabel, entries) {
  const d = newLetterheadDoc_('Safety Meeting Attendance Report', project, periodLabel);
  const body = d.body;
  const multiDay = periodLabel.indexOf(' to ') !== -1;

  // One section per meeting: same day + topic + leader. A day can have more
  // than one (e.g. a second shift toolbox talk).
  const byMeeting = {};
  entries.forEach(e => {
    const day = e.ts.slice(0, 10);
    const key = day + '|' + e.topic + '|' + e.leader;
    byMeeting[key] = byMeeting[key] || { day: day, topic: e.topic, leader: e.leader, attendees: [] };
    byMeeting[key].attendees.push(e);
  });

  Object.keys(byMeeting).sort().forEach(key => {
    const m = byMeeting[key];
    appendSectionHeading_(body, (multiDay ? m.day + '  ·  ' : '') + 'Topic: ' + (m.topic || 'Not given'));
    body.appendParagraph('Led by: ' + (m.leader || 'Not given')).editAsText().setFontSize(11);

    const tableData = [['Name', 'Company', 'Trade', 'Time', 'Acknowledged']];
    m.attendees.forEach(a => {
      tableData.push([a.name, a.company, a.trade, a.ts.slice(11, 16), a.ack ? 'Yes (' + a.ack + ')' : 'Yes']);
    });
    appendBrandTable_(body, tableData);
  });

  return finishLetterheadPdf_(d.doc, body);
}

/**
 * Jobsite sign in report for one site over periodLabel, from
 * summarizeSignIns_ output: daily summary (people, hours, never signed out),
 * manpower by company, each day's visits, then deliveries.
 */
function buildSignInPdf_(site, periodLabel, summary) {
  const d = newLetterheadDoc_('Jobsite Sign In Report', site, periodLabel);
  const body = d.body;
  const visits = summary.visits;
  const hrs = h => h == null ? '' : h.toFixed(1);
  const person = v => v.name.toLowerCase() + '|' + v.company.toLowerCase();
  const days = Array.from(new Set(visits.map(v => v.day))).sort();

  if (visits.length) {
    appendSectionHeading_(body, 'Daily summary');
    const summaryRows = [['Date', 'People on site', 'Hours', 'Not signed out']];
    let totalHours = 0, totalOpen = 0;
    days.forEach(day => {
      const dv = visits.filter(v => v.day === day);
      const h = dv.reduce((s, v) => s + (v.hours || 0), 0);
      const open = dv.filter(v => v.in && !v.out).length;
      totalHours += h; totalOpen += open;
      summaryRows.push([dayLabel_(day), String(new Set(dv.map(person)).size), hrs(h), String(open)]);
    });
    summaryRows.push(['Total', String(new Set(visits.map(person)).size) + ' different people',
                      hrs(totalHours), String(totalOpen)]);
    const st = appendBrandTable_(body, summaryRows);
    st.getRow(summaryRows.length - 1).editAsText().setBold(true);

    appendSectionHeading_(body, 'Manpower by company');
    const byCo = {};
    visits.forEach(v => {
      const co = v.company || 'Not given';
      byCo[co] = byCo[co] || { workerDays: new Set(), hours: 0 };
      byCo[co].workerDays.add(v.day + '|' + v.name.toLowerCase());
      byCo[co].hours += v.hours || 0;
    });
    const coRows = [['Company', 'Worker days', 'Hours']];
    Object.keys(byCo)
      .sort((a, b) => byCo[b].workerDays.size - byCo[a].workerDays.size || a.localeCompare(b))
      .forEach(co => coRows.push([co, String(byCo[co].workerDays.size), hrs(byCo[co].hours)]));
    appendBrandTable_(body, coRows);

    days.forEach(day => {
      appendSectionHeading_(body, dayLabel_(day));
      const dayVisits = visits.filter(v => v.day === day);
      const rows = [['Name', 'Company', 'Trade', 'In', 'Out', 'Hours']];
      dayVisits.forEach(v => rows.push([
        v.name, v.company, v.trade,
        v.in ? v.in.slice(11, 16) : 'Not recorded',
        v.out ? v.out.slice(11, 16) : 'Not signed out',
        hrs(v.hours)
      ]));
      const t = appendBrandTable_(body, rows);
      dayVisits.forEach((v, i) => {
        if (!v.out) t.getRow(i + 1).getCell(4).editAsText().setForegroundColor('#A0442F').setBold(true);
      });
    });
  } else {
    appendSectionHeading_(body, 'No worker sign ins this period');
  }

  if (summary.deliveries.length) {
    appendSectionHeading_(body, 'Deliveries');
    const rows = [['Date', 'Time', 'Driver', 'Company', 'Material', 'PO / Ticket', 'For']];
    summary.deliveries.forEach(e => rows.push([
      e.day, e.ts.slice(11, 16), e.name, e.company, e.material, e.po, e.forTrade
    ]));
    appendBrandTable_(body, rows);
  }

  return finishLetterheadPdf_(d.doc, body);
}

/** "Mon 2026-10-26". Noon UTC keeps the date stable in any US timezone. */
function dayLabel_(day) {
  return Utilities.formatDate(new Date(day + 'T12:00:00Z'), TIMEZONE, 'EEE') + ' ' + day;
}


/* ---------- helpers ---------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function trim_(v) {
  return v == null ? '' : String(v).trim();
}

/**
 * Normalize a value read from the Date column to 'yyyy-MM-dd' in Eastern time.
 * Google Sheets may return either the original string we wrote or a Date object
 * (it auto parses strings that look like dates). Handles both.
 */
function toDateStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
  }
  return String(v);
}

/**
 * Normalize a value read from the Timestamp column to 'yyyy-MM-dd HH:mm:ss'
 * in Eastern time. Same rationale as toDateStr_.
 */
function toTimestampStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(v);
}


/* ---------- delivery helpers ---------- */

/**
 * Find or create the shared photo folder in the account's Drive.
 * Uses the first folder that matches DELIVERY_PHOTO_FOLDER_NAME. If several
 * exist (previous manual creation), the first one wins. Sharing is set in
 * setup() so files added later inherit anyone with link view access.
 */
function getPhotoFolder_() {
  const it = DriveApp.getFoldersByName(DELIVERY_PHOTO_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DELIVERY_PHOTO_FOLDER_NAME);
}

/**
 * Decode a data URL and save it as a JPEG in the photo folder.
 * Filename shape: <SITE>_<yyyy-MM-dd_HHmmss>_<lastname>_<kind>.jpg with unsafe
 * chars stripped. `kind` is 'MAT' for material condition or 'BOL' for the bill
 * of lading. Returns { url, blob } so the caller can both record the URL and
 * inline the image in the notification email without re fetching from Drive.
 */
function savePhoto_(dataUrl, site, driverName, kind) {
  const match = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/.exec(String(dataUrl));
  if (!match) throw new Error('bad photo data url');
  const mime = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const ext = mime === 'image/png' ? 'png' : 'jpg';

  const now = new Date();
  const stamp = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd_HHmmss');
  const safeSite = String(site || 'site').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  const lastName = (String(driverName || '').trim().split(/\s+/).pop() || 'driver')
                     .replace(/[^A-Za-z0-9]+/g, '');
  const kindTag = String(kind || 'IMG').replace(/[^A-Za-z0-9]+/g, '').toUpperCase() || 'IMG';
  const filename = safeSite + '_' + stamp + '_' + lastName + '_' + kindTag + '.' + ext;

  const blob = Utilities.newBlob(bytes, mime, filename);
  const file = getPhotoFolder_().createFile(blob);
  return { url: file.getUrl(), blob: blob };
}

/**
 * OCR a photo of a bill of lading using the Advanced Drive Service.
 * Requires: Services (left sidebar in the script editor) → + → Drive API →
 * Add. Identifier stays as `Drive` (v2). See SETUP_GUIDE Part H.
 *
 * Approach: upload the image with mimeType application/vnd.google-apps.document
 * and ocr:true. Google converts the image to a Doc using its OCR engine. We
 * read the Doc's text, then trash the temp Doc so the folder stays clean.
 *
 * Returns { rawText, detectedPo, detectedMaterial }. Any field may be empty if
 * OCR read nothing or the regexes didn't hit.
 *
 * Best effort only. On low quality photos this returns garbage but never
 * throws to the caller as a hard fail; the caller wraps the call in try/catch
 * so a bad OCR still lets the sign in row save.
 */
function scanBol_(blob) {
  if (typeof Drive === 'undefined' || !Drive.Files || typeof Drive.Files.insert !== 'function') {
    throw new Error('Advanced Drive Service not enabled. See SETUP_GUIDE Part H.');
  }
  const created = Drive.Files.insert(
    { title: 'BOL_OCR_temp_' + Date.now(),
      mimeType: 'application/vnd.google-apps.document' },
    blob,
    { ocr: true, ocrLanguage: 'en' }
  );
  let rawText = '';
  try {
    rawText = DocumentApp.openById(created.id).getBody().getText() || '';
  } finally {
    try { DriveApp.getFileById(created.id).setTrashed(true); } catch (e) { /* best effort */ }
  }

  const detectedPo       = extractPo_(rawText);
  const detectedMaterial = extractMaterial_(rawText);
  return { rawText: rawText, detectedPo: detectedPo, detectedMaterial: detectedMaterial };
}

/**
 * Find the first plausible PO / purchase order number in OCR text.
 * Skips single digit "matches" (would be false positives on random labels).
 * Returns '' if no confident hit.
 */
function extractPo_(text) {
  if (!text) return '';
  // "PO", "P.O.", "P O", or "Purchase Order", optional "#", "No", ":", then id.
  const re = /(?:P\.?\s*O\.?|Purchase\s*Order)\s*(?:number|no\.?|#)?\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-]{3,20})/i;
  const m = text.match(re);
  return m ? m[1].toUpperCase() : '';
}

/**
 * Best effort material description from OCR text. Looks for common column
 * headers on a BOL (Description of Articles, Description, Material, Product,
 * Item) and grabs the next line of substance. If nothing hits, returns ''.
 */
function extractMaterial_(text) {
  if (!text) return '';
  const lines = text.split(/\r?\n/).map(l => l.trim());
  const headerRe = /^(?:description(?:\s+of\s+articles)?|material|product|item(?:s)?|commodity)\b/i;
  for (let i = 0; i < lines.length; i++) {
    if (!headerRe.test(lines[i])) continue;
    // grab the next non blank line
    for (let j = i + 1; j < Math.min(i + 5, lines.length); j++) {
      const s = lines[j];
      if (s && s.length >= 4 && !/^-+$/.test(s)) {
        return s.slice(0, 120);
      }
    }
  }
  return '';
}

/**
 * Send the delivery notification email.
 * To: DELIVERY_ALERT_PRIMARY
 * Cc: SITE_SUPERS[site] (comma separated list is fine, empty means no CC)
 * The photo is both inlined via CID and attached as a separate file so it
 * survives clients that strip inline images.
 */
function notifyDelivery_(d) {
  const to = DELIVERY_ALERT_PRIMARY;
  if (!to) return;
  const cc = String(SITE_SUPERS[d.site] || '').trim();
  const hasBol = !!d.bolPhotoBlob;

  // Trim the auto appended BOL scan from Notes for the email body; recipients
  // do not need the raw OCR dump inline. It stays in the Sheet for audit.
  const cleanNotes = String(d.notes || '').split(/\n\n?--- BOL scan ---|BOL scan:/)[0].trim();

  const poLabel = d.detectedPo && d.po === d.detectedPo
    ? d.po + ' (detected from BOL)'
    : (d.po || 'not given');

  const subject = 'Delivery arrived at ' + d.site
    + (d.detectedPo ? ' (PO ' + d.detectedPo + ')' : '');

  const plainLines = [
    'A delivery just signed in.',
    '',
    'Site: '        + d.site,
    'Time: '        + d.tsStr + ' Eastern',
    'Driver: '      + d.name,
    'Company: '     + d.company,
    'Material: '    + (d.material || 'not given'),
    'PO / Ticket: ' + poLabel,
    'For: '         + (d.forTrade || 'not given'),
    'Notes: '       + (cleanNotes || 'none'),
    '',
    'Material photo: ' + d.materialPhotoUrl
  ];
  if (hasBol) plainLines.push('BOL photo:      ' + d.bolPhotoUrl);
  else        plainLines.push('BOL photo:      not provided by driver');
  const plainBody = plainLines.join('\n');

  const photoBlockHtml = hasBol
    ? // Two column layout when BOL is present.
      '<table style="width:100%;margin-top:16px;border-collapse:separate;border-spacing:8px 0">' +
        '<tr>' +
          '<td style="width:50%;vertical-align:top">' +
            '<p style="margin:0 0 6px;font-weight:bold;color:#0B4F30">Material condition</p>' +
            '<img src="cid:materialPhoto" alt="Material" style="max-width:100%;border-radius:8px;border:1px solid #D9E5D8">' +
            '<p style="margin:6px 0 0;font-size:12px;color:#5B6B61"><a href="' + escapeHtml_(d.materialPhotoUrl) + '">Open in Drive</a></p>' +
          '</td>' +
          '<td style="width:50%;vertical-align:top">' +
            '<p style="margin:0 0 6px;font-weight:bold;color:#0B4F30">Bill of lading</p>' +
            '<img src="cid:bolPhoto" alt="BOL" style="max-width:100%;border-radius:8px;border:1px solid #D9E5D8">' +
            '<p style="margin:6px 0 0;font-size:12px;color:#5B6B61"><a href="' + escapeHtml_(d.bolPhotoUrl) + '">Open in Drive</a></p>' +
          '</td>' +
        '</tr>' +
      '</table>'
    : // Single column layout when BOL was skipped.
      '<div style="margin-top:16px">' +
        '<p style="margin:0 0 6px;font-weight:bold;color:#0B4F30">Material condition</p>' +
        '<img src="cid:materialPhoto" alt="Material" style="max-width:100%;border-radius:8px;border:1px solid #D9E5D8">' +
        '<p style="margin:6px 0 0;font-size:12px;color:#5B6B61"><a href="' + escapeHtml_(d.materialPhotoUrl) + '">Open in Drive</a></p>' +
      '</div>' +
      '<p style="margin:12px 0 0;font-size:13px;color:#A0442F">Driver did not provide a BOL photo. Details above are what they typed.</p>';

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:640px;color:#22302A">' +
      '<h2 style="color:#0B4F30;margin:0 0 10px">Delivery arrived at ' + escapeHtml_(d.site) + '</h2>' +
      '<table style="border-collapse:collapse;font-size:14px">' +
        row_('Time',        d.tsStr + ' Eastern') +
        row_('Driver',      d.name) +
        row_('Company',     d.company) +
        row_('Material',    d.material || 'not given') +
        row_('PO / Ticket', poLabel) +
        row_('For',         d.forTrade || 'not given') +
        row_('Notes',       cleanNotes || 'none') +
      '</table>' +
      photoBlockHtml +
    '</div>';

  const inlineImages = { materialPhoto: d.materialPhotoBlob };
  const attachments  = [ d.materialPhotoBlob ];
  if (hasBol) {
    inlineImages.bolPhoto = d.bolPhotoBlob;
    attachments.push(d.bolPhotoBlob);
  }

  const options = {
    htmlBody: htmlBody,
    inlineImages: inlineImages,
    attachments: attachments,
    name: 'PruittHealth Jobsite Sign In'
  };
  if (cc) options.cc = cc;

  MailApp.sendEmail(to, subject, plainBody, options);
}

function row_(label, value) {
  return '<tr>' +
    '<td style="padding:3px 12px 3px 0;color:#5B6B61">' + escapeHtml_(label) + '</td>' +
    '<td style="padding:3px 0;font-weight:bold">' + escapeHtml_(value) + '</td>' +
  '</tr>';
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
