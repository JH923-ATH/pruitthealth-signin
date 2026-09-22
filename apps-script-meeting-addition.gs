/**
 * REFERENCE ONLY. This is not a standalone deployable file — it documents
 * exactly what was added to Code.gs to support meeting-signin.html. If the
 * live Apps Script project's Code.gs ever needs to be reconstructed or
 * patched by hand (rather than pasting the repo's Code.gs wholesale), copy
 * the four pieces below into their matching spots. If you're just deploying
 * from this repo, ignore this file and use Code.gs directly — it already
 * has all of this baked in.
 *
 * ---------------------------------------------------------------------
 * 1. Add near the top CONFIG block, alongside SHEET_NAME / HEADERS:
 * ---------------------------------------------------------------------
 *
 *   const MEETING_SHEET_NAME = 'Meetings';
 *   const MEETING_HEADERS    = [
 *     'Timestamp', 'Meeting Type', 'Project', 'Meeting Date',
 *     'Name', 'Company', 'Role', 'Email', 'Phone'
 *   ];
 *
 * ---------------------------------------------------------------------
 * 2. In doPost(e), route MeetingSignIn before the existing In/Out check:
 * ---------------------------------------------------------------------
 *
 *   const p = JSON.parse(e.postData.contents);
 *   const action = String(p.action || '').trim();
 *   if (action === 'MeetingSignIn') {
 *     return handleMeetingSignIn_(p);
 *   }
 *   if (action !== 'In' && action !== 'Out') {
 *     return jsonOut_({ ok: false, error: 'bad action' });
 *   }
 *
 * ---------------------------------------------------------------------
 * 3 & 4. Add these two functions anywhere in the file (Code.gs keeps them
 *        just above the existing getSheet_ helper):
 * ---------------------------------------------------------------------
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

/**
 * Both functions rely on TIMEZONE, trim_, and jsonOut_, which already exist
 * in Code.gs — nothing else to bring over. No changes to doGet, setup(),
 * or the SignIns sheet contract.
 */
