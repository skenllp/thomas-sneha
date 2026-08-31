/**
 * RSVP → Google Sheet backend for the thomas-sneha wedding invitation.
 *
 * SETUP (5 minutes):
 * 1. Go to https://sheets.google.com and create a new blank spreadsheet.
 *    Name it whatever you like, e.g. "Wedding RSVPs".
 * 2. In that sheet, go to Extensions → Apps Script.
 * 3. Delete anything in the editor and paste this whole file in.
 * 4. Click Deploy → New deployment.
 *      - Click the gear icon next to "Select type" → choose "Web app".
 *      - Description: anything, e.g. "RSVP intake".
 *      - Execute as: Me.
 *      - Who has access: Anyone.
 *    Click Deploy, then authorize the script when Google prompts you
 *    (click "Advanced" → "Go to ... (unsafe)" if you see a warning screen —
 *    this is expected for your own script).
 * 5. Copy the "Web app URL" it gives you.
 * 6. In index.html, find the line:
 *        const RSVP_SHEET_URL = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
 *    and paste your URL between the quotes.
 * 7. Re-deploy the site. Submit a test RSVP — a "RSVPs" tab will appear in
 *    your sheet automatically with each response as a new row.
 *
 * If you ever edit this script after the first deploy, use
 * Deploy → Manage deployments → Edit (pencil icon) → New version → Deploy,
 * so the same Web app URL keeps working.
 */

function doPost(e) {
  var sheet = getOrCreateSheet_();

  var name = (e.parameter.name || '').toString().trim();
  var attendance = (e.parameter.attendance || '').toString().trim();
  var message = (e.parameter.message || '').toString().trim();

  var attendanceLabel = attendance === 'joyfully-accept'
    ? 'Joyfully Accept'
    : (attendance === 'regretfully-decline' ? 'Regretfully Decline' : attendance);

  sheet.appendRow([
    new Date(),
    name,
    attendanceLabel,
    message
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RSVPs');
  if (!sheet) {
    sheet = ss.insertSheet('RSVPs');
    sheet.appendRow(['Timestamp', 'Name', 'Attendance', 'Message']);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
