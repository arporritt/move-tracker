// ============================================================
// Move Tracker — Google Apps Script backend
// State is stored as a JSON file in Drive (unlimited size).
// On first run, auto-migrates any existing data from the old Sheet cells.
// ============================================================

const PASSWORD = 'lakegeneva';
const SHEET_NAME = 'state';
const STATE_FILENAME = 'movetracker_state.json';
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

function getStateFile_() {
  const folder = getUploadFolder_();
  const it = folder.getFilesByName(STATE_FILENAME);
  if (it.hasNext()) return it.next();
  return folder.createFile(STATE_FILENAME, '{}', 'application/json');
}

function readStateFromCells_() {
  // Legacy: read from up to 3 sheet cells (one-time migration source)
  const sh = getSheet_();
  const result = {};
  const cells = { core: 'A1', design: 'A2', layouts: 'A3' };
  const core = String(sh.getRange(cells.core).getValue() || '');
  if (core) { try { Object.assign(result, JSON.parse(core)); } catch (e) {} }
  const design = String(sh.getRange(cells.design).getValue() || '');
  if (design) { try { const d = JSON.parse(design); if (d.inspiration) result.inspiration = d.inspiration; } catch (e) {} }
  const layouts = String(sh.getRange(cells.layouts).getValue() || '');
  if (layouts) { try { const l = JSON.parse(layouts); if (l.layouts) result.layouts = l.layouts; } catch (e) {} }
  return result;
}

function readState_() {
  const file = getStateFile_();
  const content = file.getBlob().getDataAsString();
  if (content && content.length > 2) { // not '{}' or empty
    try {
      const parsed = JSON.parse(content);
      if (parsed && (parsed.tasks || parsed.budget || parsed.inspiration)) return parsed;
    } catch (e) {}
  }
  // First-time migration from old Sheet cells
  const cellData = readStateFromCells_();
  if (cellData && Object.keys(cellData).length > 0) {
    file.setContent(JSON.stringify(cellData));
    return cellData;
  }
  return {};
}

function writeState_(data) {
  const file = getStateFile_();
  file.setContent(JSON.stringify(data));
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter.p : '';
  if (p !== PASSWORD) {
    return ContentService.createTextOutput('null').setMimeType(ContentService.MimeType.TEXT);
  }
  // Pull the LRV from a Sherwin-Williams color page (server-side fetch avoids browser CORS).
  if (e && e.parameter && e.parameter.action === 'getLrv') {
    var out = { ok: false, lrv: '' };
    try {
      const url = e.parameter.url || '';
      if (/^https?:\/\//.test(url)) {
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15)' } });
        const html = resp.getContentText();
        var m = html.match(/"lrv"\s*:\s*"?(\d{1,3}(?:\.\d+)?)"?/i)
             || html.match(/lrv["'\s:>=]{1,14}?(\d{1,3}(?:\.\d+)?)/i)
             || html.match(/LRV[^0-9]{0,18}(\d{1,3}(?:\.\d+)?)/i);
        if (m) { out.lrv = m[1]; out.ok = true; }
      }
    } catch (err) { out.error = String(err); }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
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
        ok: true, id: id, url: 'https://lh3.googleusercontent.com/d/' + id
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  writeState_(body.data || {});
  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}
