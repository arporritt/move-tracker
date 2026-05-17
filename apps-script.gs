// ============================================================
// Move Tracker — Google Apps Script backend
// Paste this entire file into script.google.com (bound to your Sheet)
// Then Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone
// Paste the resulting Web App URL into index.html at SCRIPT_URL
// ============================================================

const PASSWORD = 'lakegeneva';
const SHEET_NAME = 'state';
const CELL = 'A1';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter.p : '';
  if (p !== PASSWORD) {
    return ContentService.createTextOutput('null').setMimeType(ContentService.MimeType.TEXT);
  }
  const sh = getSheet_();
  const val = sh.getRange(CELL).getValue();
  const out = val ? String(val) : 'null';
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  if (body.p !== PASSWORD) {
    return ContentService.createTextOutput('{"ok":false,"error":"unauthorized"}').setMimeType(ContentService.MimeType.JSON);
  }
  const sh = getSheet_();
  sh.getRange(CELL).setValue(JSON.stringify(body.data || {}));
  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}
