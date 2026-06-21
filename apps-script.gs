// ============================================================
// Move Tracker — Google Apps Script backend
// State is now split across 3 cells (A1 core, A2 design, A3 layouts) for ~150KB capacity.
// ============================================================

const PASSWORD = 'lakegeneva';
const SHEET_NAME = 'state';
const CELLS = { core: 'A1', design: 'A2', layouts: 'A3' };
const UPLOAD_FOLDER_NAME = 'Move Tracker Uploads';

function authorizeDrive() {
  const folder = DriveApp.createFolder('__move_tracker_auth_test__');
  folder.setTrashed(true);
  return 'Drive write access granted.';
}

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

function readState_() {
  const sh = getSheet_();
  const result = {};
  const core = String(sh.getRange(CELLS.core).getValue() || '');
  if (core) {
    try { Object.assign(result, JSON.parse(core)); } catch (e) {}
  }
  const design = String(sh.getRange(CELLS.design).getValue() || '');
  if (design) {
    try {
      const d = JSON.parse(design);
      if (d.inspiration) result.inspiration = d.inspiration;
    } catch (e) {}
  }
  const layouts = String(sh.getRange(CELLS.layouts).getValue() || '');
  if (layouts) {
    try {
      const l = JSON.parse(layouts);
      if (l.layouts) result.layouts = l.layouts;
    } catch (e) {}
  }
  return result;
}

function writeState_(data) {
  const sh = getSheet_();
  const inspiration = data.inspiration;
  const layouts = data.layouts;
  const core = {};
  for (const k of Object.keys(data)) {
    if (k !== 'inspiration' && k !== 'layouts') core[k] = data[k];
  }
  sh.getRange(CELLS.core).setValue(JSON.stringify(core));
  sh.getRange(CELLS.design).setValue(JSON.stringify({ inspiration: inspiration || {} }));
  sh.getRange(CELLS.layouts).setValue(JSON.stringify({ layouts: layouts || {} }));
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter.p : '';
  if (p !== PASSWORD) {
    return ContentService.createTextOutput('null').setMimeType(ContentService.MimeType.TEXT);
  }
  const merged = readState_();
  return ContentService.createTextOutput(JSON.stringify(merged)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  if (body.p !== PASSWORD) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:'unauthorized'})).setMimeType(ContentService.MimeType.JSON);
  }
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
  writeState_(body.data || {});
  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}
