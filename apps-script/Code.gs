/**
 * HROD backend — stores employee records in a Google Sheet.
 *
 * Setup (see README.md for the full step-by-step):
 * 1. Create a Google Sheet (any name).
 * 2. Extensions > Apps Script. Delete the placeholder code and paste this
 *    whole file in.
 * 3. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the Web app URL (ends in /exec) and paste it into the HROD page's
 *    "⚙️ ตั้งค่าฐานข้อมูล" (Settings) dialog.
 *
 * Data model: one sheet named "employees" with columns:
 *   A: id        (text, unique employee id)
 *   B: json      (the whole employee record, as a JSON string)
 *   C: updatedAt (ISO timestamp, set by this script)
 * The sheet is created automatically on first write if it doesn't exist.
 *
 * Security note: with "Who has access: Anyone", anyone who has this exact
 * URL can read and write every row in the sheet (there is no login).
 * That's the trade-off for a simple, no-backend, no-auth setup. Keep the
 * URL private to people who should be able to use the HR tool, and don't
 * put anything in it you wouldn't want someone with the link to see.
 */

var SHEET_NAME = 'employees';

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') {
      return jsonResponse({ ok: true, employees: listEmployees() });
    }
    return jsonResponse({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    if (action === 'set') {
      upsertEmployee(body.id, body.data);
      return jsonResponse({ ok: true });
    }
    if (action === 'delete') {
      deleteEmployee(body.id);
      return jsonResponse({ ok: true });
    }
    return jsonResponse({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'json', 'updatedAt']);
  }
  return sheet;
}

function listEmployees() {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var id = values[i][0];
    var json = values[i][1];
    var updatedAt = values[i][2];
    if (!id) continue;
    try {
      out.push({ id: String(id), data: JSON.parse(json), updatedAt: updatedAt });
    } catch (parseErr) {
      // skip a row that doesn't contain valid JSON rather than failing the whole list
    }
  }
  return out;
}

function findRow_(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

function upsertEmployee(id, data) {
  if (!id) throw new Error('missing id');
  var sheet = getSheet_();
  var row = findRow_(sheet, id);
  var json = JSON.stringify(data);
  var now = new Date().toISOString();
  if (row === -1) {
    sheet.appendRow([id, json, now]);
  } else {
    sheet.getRange(row, 2).setValue(json);
    sheet.getRange(row, 3).setValue(now);
  }
}

function deleteEmployee(id) {
  var sheet = getSheet_();
  var row = findRow_(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
