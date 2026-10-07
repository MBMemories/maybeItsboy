/**
 * Receives wishes (wishes.html) and RSVPs (rsvp.html) and writes them to Google Sheets.
 *
 * Setup:
 *  1. Create a Google Sheet, then open Extensions > Apps Script.
 *  2. Replace the contents of Code.gs with this file and save.
 *  3. Deploy > New deployment > type "Web app"
 *       Execute as: Me
 *       Who has access: Anyone
 *  4. Copy the Web app URL (ends in /exec) into the ENDPOINT constants in
 *     wishes.html and rsvp.html.
 *  5. Select the authorize function and Run it once to grant access to the
 *     RSVP spreadsheet.
 *
 * After editing this script, publish it again via Deploy > Manage deployments >
 * Edit > Version: New version, so the same URL serves the new code.
 */

const WISHES_SHEET = 'Wishes';

// RSVPs go to the "รายชื่อ" tab of this spreadsheet, in the same column order
// as the original Google Form responses
const RSVP_SPREADSHEET_ID = '1fyyxTHv8v8QoYbd3fekhBi-PtEMZidcm5tjcAkp4Gq4';
const RSVP_SHEET = 'รายชื่อ';
const RSVP_HEADERS = ['ประทับเวลา', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'ความสัมพันธ์',
                      'การยืนยันการมาร่วมงาน', 'จำนวนผู้ติดตาม (รวมตัวแขก)', 'ข้อความถึงบ่าวสาว'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);

    // Spam trap field is hidden from people; pretend success for bots
    if (data.website) return json({ ok: true });

    return data.type === 'rsvp' ? saveRsvp(data) : saveWish(data);
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function saveWish(data) {
  const name = clean(data.name, 80);
  const message = clean(data.message, 1000);
  if (!name || !message) return json({ ok: false, error: 'missing' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(WISHES_SHEET) || ss.insertSheet(WISHES_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['เวลา', 'ชื่อ', 'คำอวยพร']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), name, message]);
  return json({ ok: true });
}

function saveRsvp(data) {
  const fullName = clean(data.fullName, 120);
  const nickname = clean(data.nickname, 60);
  const relation = clean(data.relation, 120);
  const attending = data.attending === 'yes';
  const guests = attending ? Number(data.guests) : 0;
  const message = clean(data.message, 1000);

  if (!fullName || !relation || !['yes', 'no'].includes(data.attending)) {
    return json({ ok: false, error: 'missing' });
  }
  if (attending && !(guests >= 1 && guests <= 4)) {
    return json({ ok: false, error: 'guests' });
  }

  const ss = SpreadsheetApp.openById(RSVP_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RSVP_SHEET) || ss.insertSheet(RSVP_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(RSVP_HEADERS);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date(),
    fullName,
    nickname,
    relation,
    attending ? 'ยินดีไปร่วมงาน' : 'ไม่สะดวกมาร่วมงาน',
    attending ? guests + ' ท่าน' : '',
    message
  ]);
  return json({ ok: true });
}

// Run once from the editor to grant access to the RSVP spreadsheet
function authorize() {
  const sheet = SpreadsheetApp.openById(RSVP_SPREADSHEET_ID).getSheetByName(RSVP_SHEET);
  Logger.log(sheet ? 'พบชีต ' + RSVP_SHEET : 'ไม่พบชีต ' + RSVP_SHEET + ' (จะสร้างให้อัตโนมัติเมื่อมีคนตอบรับ)');
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
