/**
 * Receives wishes from wishes.html and appends them to this Google Sheet.
 *
 * Setup:
 *  1. Create a Google Sheet, then open Extensions > Apps Script.
 *  2. Replace the contents of Code.gs with this file and save.
 *  3. Deploy > New deployment > type "Web app"
 *       Execute as: Me
 *       Who has access: Anyone
 *  4. Copy the Web app URL (ends in /exec) into WISHES_ENDPOINT in wishes.html.
 *
 * After editing this script, publish it again via Deploy > Manage deployments >
 * Edit > Version: New version, so the same URL serves the new code.
 */

const SHEET_NAME = 'Wishes';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);

    // Spam trap field is hidden from people; pretend success for bots
    if (data.website) return json({ ok: true });

    const name = clean(data.name, 80);
    const message = clean(data.message, 1000);
    if (!name || !message) return json({ ok: false, error: 'missing' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['เวลา', 'ชื่อ', 'คำอวยพร']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), name, message]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Trim, cap length, and stop text like "=IMPORTXML(...)" from running as a formula
function clean(value, max) {
  const v = String(value || '').trim().slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
