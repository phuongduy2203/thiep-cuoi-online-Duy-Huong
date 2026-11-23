// Google Apps Script: lưu lời chúc vào Google Sheets.
// Cách dùng: tạo Google Sheet mới -> Tiện ích mở rộng -> Apps Script -> dán file này -> Triển khai
// -> Ứng dụng web (Thực thi bằng: Tôi; Quyền truy cập: Bất kỳ ai) -> sao chép URL /exec vào CONFIG.wishApi trong index.html.
const SHEET_NAME = 'LoiChuc';
const RSVP_SHEET_NAME = 'XacNhan';
const MAX_NAME = 80;
const MAX_MESSAGE = 1000;
const SHOW_LAST = 50;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Thời gian', 'Tên', 'Lời chúc']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getRsvpSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RSVP_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(RSVP_SHEET_NAME);
    sh.appendRow(['Thời gian', 'Tên', 'Tham dự', 'Số người']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Chặn Google Sheets hiểu nội dung bắt đầu bằng = + - @ là công thức
function safe_(s) {
  s = String(s || '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = safe_(d.name).slice(0, MAX_NAME);
    let row, sheet;
    if (d.type === 'rsvp') {
      const guests = Math.max(0, Math.min(50, parseInt(d.guests, 10) || 0));
      const attendance = d.attendance === 'Không tham dự' ? 'Không tham dự' : 'Có tham dự';
      if (!name) return json_({ ok: false, error: 'missing' });
      sheet = getRsvpSheet_();
      row = [new Date(), name, attendance, guests];
    } else {
      const message = safe_(d.message).slice(0, MAX_MESSAGE);
      if (!name || !message) return json_({ ok: false, error: 'missing' });
      sheet = getSheet_();
      row = [new Date(), name, message];
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet.appendRow(row);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ ok: true, wishes: [] });
  const start = Math.max(2, last - SHOW_LAST + 1);
  const rows = sh.getRange(start, 1, last - start + 1, 3).getValues();
  const wishes = rows
    .map(r => ({ time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), name: String(r[1]), message: String(r[2]) }))
    .reverse();
  return json_({ ok: true, wishes: wishes });
}
