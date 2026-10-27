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
 * GET handler for the sign out roster.
 * Query: ?site=<siteName>
 * Returns a JSON array of { Name, Company, Trade, Timestamp } for people
 * currently on site at that site TODAY (signed in and not yet signed out).
 *
 * Locked down on purpose:
 *   - Only today. Any date parameter in the query is ignored.
 *   - Only people currently on site. Historical rows are never returned.
 *   - Never returns Safety Ack.
 *   - Never returns any delivery column (Material, PO, For Trade, Notes, Photo URL).
 *   - If site is missing or unknown, returns an empty array (not an error).
 */
function doGet(e) {
  try {
    const site = (e && e.parameter && e.parameter.site) ? String(e.parameter.site) : '';
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
            '<p style="margin:0 0 6px;font-weight:bold;color:#17402A">Material condition</p>' +
            '<img src="cid:materialPhoto" alt="Material" style="max-width:100%;border-radius:8px;border:1px solid #D9E5D8">' +
            '<p style="margin:6px 0 0;font-size:12px;color:#5B6B61"><a href="' + escapeHtml_(d.materialPhotoUrl) + '">Open in Drive</a></p>' +
          '</td>' +
          '<td style="width:50%;vertical-align:top">' +
            '<p style="margin:0 0 6px;font-weight:bold;color:#17402A">Bill of lading</p>' +
            '<img src="cid:bolPhoto" alt="BOL" style="max-width:100%;border-radius:8px;border:1px solid #D9E5D8">' +
            '<p style="margin:6px 0 0;font-size:12px;color:#5B6B61"><a href="' + escapeHtml_(d.bolPhotoUrl) + '">Open in Drive</a></p>' +
          '</td>' +
        '</tr>' +
      '</table>'
    : // Single column layout when BOL was skipped.
      '<div style="margin-top:16px">' +
        '<p style="margin:0 0 6px;font-weight:bold;color:#17402A">Material condition</p>' +
        '<img src="cid:materialPhoto" alt="Material" style="max-width:100%;border-radius:8px;border:1px solid #D9E5D8">' +
        '<p style="margin:6px 0 0;font-size:12px;color:#5B6B61"><a href="' + escapeHtml_(d.materialPhotoUrl) + '">Open in Drive</a></p>' +
      '</div>' +
      '<p style="margin:12px 0 0;font-size:13px;color:#A0442F">Driver did not provide a BOL photo. Details above are what they typed.</p>';

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:640px;color:#22302A">' +
      '<h2 style="color:#17402A;margin:0 0 10px">Delivery arrived at ' + escapeHtml_(d.site) + '</h2>' +
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
