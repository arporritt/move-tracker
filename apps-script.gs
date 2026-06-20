// ============================================================
// Move Tracker — Google Apps Script backend
// HOW TO UPDATE (one time): script.google.com → open project → replace ALL code with this →
//   Deploy → Manage deployments → ✎ edit existing deployment → Version: New version → Deploy
//   (Same Web App URL stays; you'll re-authorize Drive access on first run.)
// ============================================================

const PASSWORD = 'lakegeneva';
const SHEET_NAME = 'state';
const CELL = 'A1';
const UPLOAD_FOLDER_NAME = 'Move Tracker Uploads';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function getUploadFolder_() {
  const it = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(UPLOAD_FOLDER_NAME);
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
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:'unauthorized'})).setMimeType(ContentService.MimeType.JSON);
  }
  // Image upload to Drive
  if (body.action === 'upload') {
    try {
      const folder = getUploadFolder_();
      const dataPart = body.imageData || '';
      const commaIdx = dataPart.indexOf(',');
      const b64 = commaIdx >= 0 ? dataPart.substring(commaIdx + 1) : dataPart;
      const bytes = Utilities.base64Decode(b64);
      const mime = body.mimeType || 'image/jpeg';
      const ext = mime === 'image/png' ? 'png' : (mime === 'image/webp' ? 'webp' : 'jpg');
      const baseName = (body.filename || ('upload_' + Date.now())).replace(/\.[a-z0-9]+$/i, '');
      const blob = Utilities.newBlob(bytes, mime, baseName + '.' + ext);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const id = file.getId();
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        id: id,
        url: 'https://lh3.googleusercontent.com/d/' + id
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  // Default: save state JSON to the cell
  const sh = getSheet_();
  sh.getRange(CELL).setValue(JSON.stringify(body.data || {}));
  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}
