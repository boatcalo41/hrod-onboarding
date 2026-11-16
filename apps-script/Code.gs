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
 *   B: json      (the whole employee record, as a JSON string — source of truth)
 *   C: updatedAt (ISO timestamp, set by this script)
 *   D: name, E: code — mirrored out of the JSON so HR can see who's who at a glance
 *   F/G: orientation-quiz (round 1) score / date
 *   H/I: pre-probation-quiz (round 2) score / date
 *   J: how many hand-written answers are still waiting on an HR review
 * D–J are read-only summary columns written by this script every time an
 * employee record is saved — edit the person's record in the app, not these
 * columns directly, since they get overwritten on the next save.
 * The sheet is created automatically on first write if it doesn't exist.
 *
 * A second sheet named "config" holds small shared, app-wide settings that
 * HR edits from the HR Admin Portal — the onboarding/offboarding checklist
 * templates, the KPI table, the Learning Center content, and the list of
 * supervisor/buddy names. Same shape as "employees" (key, json, updatedAt),
 * created automatically on first write.
 *
 * Security note: with "Who has access: Anyone", anyone who has this exact
 * URL can read and write every row in the sheet (there is no login).
 * That's the trade-off for a simple, no-backend, no-auth setup. Keep the
 * URL private to people who should be able to use the HR tool, and don't
 * put anything in it you wouldn't want someone with the link to see.
 *
 * The orientation quiz now lives inside the app itself (employees tick
 * their answers in the "ข้อสอบปฐมนิเทศ" portal tab) and each employee's
 * answers/score are saved as part of their own record via the normal
 * "set" action below (emp.exams.round1 / emp.exams.round2) — no separate
 * Google Form response sheet is read by this script anymore.
 */

var SHEET_NAME = 'employees';
var CONFIG_SHEET_NAME = 'config';

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') {
      return jsonResponse({
        ok: true,
        employees: listEmployees(),
        config: listConfig()
      });
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
      upsertRow_(getSheet_(), body.id, body.data);
      return jsonResponse({ ok: true });
    }
    if (action === 'delete') {
      deleteRow_(getSheet_(), body.id);
      return jsonResponse({ ok: true });
    }
    if (action === 'setConfig') {
      upsertRow_(getConfigSheet_(), body.key, body.data);
      return jsonResponse({ ok: true });
    }
    return jsonResponse({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

var EMPLOYEE_SUMMARY_HEADERS = [
  'ชื่อ-นามสกุล', 'รหัสพนักงาน',
  'ข้อสอบปฐมนิเทศ (คะแนน)', 'วันที่สอบปฐมนิเทศ',
  'ข้อสอบก่อนผ่านทดลองงาน (คะแนน)', 'วันที่สอบรอบ 2',
  'รอ HR ตรวจคำตอบเขียน'
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'json', 'updatedAt'].concat(EMPLOYEE_SUMMARY_HEADERS));
  } else {
    ensureSummaryHeaders_(sheet);
  }
  return sheet;
}

function ensureSummaryHeaders_(sheet) {
  // Backfills the D..J summary header cells on a sheet that predates them,
  // without touching anything already there in A..C.
  var have = sheet.getRange(1, 4, 1, EMPLOYEE_SUMMARY_HEADERS.length).getValues()[0];
  var allBlank = have.every(function (v) { return v === '' || v === null; });
  if (allBlank) {
    sheet.getRange(1, 4, 1, EMPLOYEE_SUMMARY_HEADERS.length).setValues([EMPLOYEE_SUMMARY_HEADERS]);
  }
}

function formatExamCell_(round) {
  if (!round) return { score: '', date: '' };
  var score = (round.auto != null && round.autoMax != null) ? (round.auto + ' / ' + round.autoMax) : '';
  var date = '';
  if (round.submittedAt) {
    try { date = Utilities.formatDate(new Date(round.submittedAt), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); }
    catch (e) { date = String(round.submittedAt); }
  }
  return { score: score, date: date };
}

function employeeSummaryRow_(data) {
  var exams = data.exams || {};
  var r1 = formatExamCell_(exams.round1);
  var r2 = formatExamCell_(exams.round2);
  var pendingParts = [];
  if (exams.round1 && exams.round1.manualPending) pendingParts.push('รอบ1: ' + exams.round1.manualPending + ' ข้อ');
  if (exams.round2 && exams.round2.manualPending) pendingParts.push('รอบ2: ' + exams.round2.manualPending + ' ข้อ');
  return [
    data.name || '',
    data.code || '',
    r1.score, r1.date,
    r2.score, r2.date,
    pendingParts.join(' · ')
  ];
}

function getConfigSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.appendRow(['key', 'json', 'updatedAt']);
  }
  return sheet;
}

function listConfig() {
  var sheet = getConfigSheet_();
  var values = sheet.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < values.length; i++) {
    var key = values[i][0];
    var json = values[i][1];
    if (!key) continue;
    try { out[String(key)] = JSON.parse(json); } catch (parseErr) { /* skip bad row */ }
  }
  return out;
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

function upsertRow_(sheet, id, data) {
  if (!id) throw new Error('missing id');
  var row = findRow_(sheet, id);
  var json = JSON.stringify(data);
  var now = new Date().toISOString();
  var isEmployeeSheet = (sheet.getName() === SHEET_NAME);
  var summary = isEmployeeSheet ? employeeSummaryRow_(data) : null;
  if (row === -1) {
    var newRow = [id, json, now];
    if (summary) newRow = newRow.concat(summary);
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(row, 2).setValue(json);
    sheet.getRange(row, 3).setValue(now);
    if (summary) sheet.getRange(row, 4, 1, summary.length).setValues([summary]);
  }
}

function deleteRow_(sheet, id) {
  var row = findRow_(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
