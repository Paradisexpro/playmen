/**
 * TRADEX Portfolio Tracker - Google Apps Script Backend
 * ใช้ Google Sheets เป็นฐานข้อมูลให้เว็บ Stock & Finance Tracker
 *
 * วิธีติดตั้ง:
 * 1. เปิด Google Sheet ใหม่ 1 ไฟล์
 * 2. เมนู Extensions > Apps Script
 * 3. วางโค้ดนี้ทั้งหมด แล้วกดบันทึก
 * 4. กด Deploy > New deployment > เลือก type "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. คัดลอก Web app URL ไปวางแทนค่า API_URL ใน index.html
 */

const SHEET_NAME = 'Transactions';
const HEADERS = ['id', 'type', 'amount', 'description', 'datetime'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toLocalInput(value) {
  if (value instanceof Date) {
    const pad = (n) => String(n).padStart(2, '0');
    return value.getFullYear() + '-' + pad(value.getMonth() + 1) + '-' + pad(value.getDate()) +
      'T' + pad(value.getHours()) + ':' + pad(value.getMinutes());
  }
  return String(value);
}

function listTransactions() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values
    .filter((row) => row[0] !== '' && row[0] !== null)
    .map((row) => ({
      id: String(row[0]),
      type: String(row[1]),
      amount: Number(row[2]),
      description: String(row[3]),
      datetime: toLocalInput(row[4])
    }));
}

function addTransaction(t) {
  const sheet = getSheet();
  sheet.appendRow([
    String(t.id),
    String(t.type),
    Number(t.amount),
    String(t.description),
    String(t.datetime)
  ]);
}

function deleteTransaction(id) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function clearAll() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') {
      return jsonOutput({ status: 'ok', data: listTransactions() });
    }
    return jsonOutput({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action;

    if (action === 'add' && body.transaction) {
      addTransaction(body.transaction);
      return jsonOutput({ status: 'ok', data: listTransactions() });
    }
    if (action === 'delete' && body.id !== undefined) {
      deleteTransaction(body.id);
      return jsonOutput({ status: 'ok', data: listTransactions() });
    }
    if (action === 'clear') {
      clearAll();
      return jsonOutput({ status: 'ok', data: [] });
    }

    return jsonOutput({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  }
}
