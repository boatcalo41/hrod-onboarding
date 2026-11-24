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
var JD_SHEET_NAME = 'job_descriptions';
var JD_DRIVE_ROOT_FOLDER_NAME = 'HROD - Job Description';
var PHOTO_DRIVE_ROOT_FOLDER_NAME = 'HROD - Profile Photos';
// Same logo the HROD web app shows (assets/logo.png), served from GitHub
// Pages — fetched fresh each time a JD PDF is generated so the printed
// form always matches the app's current logo.
var ALLIED_TEK_LOGO_URL = 'https://boatcalo41.github.io/hrod-onboarding/assets/logo.png';

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') {
      return jsonResponse({
        ok: true,
        employees: listEmployees(),
        config: listConfig(),
        jobDescriptions: listJobDescriptions()
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
    if (action === 'setJD') {
      upsertJdRow_(body.id, body.data);
      return jsonResponse({ ok: true });
    }
    if (action === 'deleteJD') {
      deleteRow_(getJdSheet_(), body.id);
      return jsonResponse({ ok: true });
    }
    if (action === 'generateJdPdf') {
      var result = generateJdPdf_(body.id, body.data);
      return jsonResponse({ ok: true, pdfUrl: result.pdfUrl, pdfFileId: result.pdfFileId, pdfUpdatedAt: result.pdfUpdatedAt });
    }
    if (action === 'uploadPhoto') {
      var photo = uploadEmployeePhoto_(body.id, body.filename, body.mimeType, body.base64);
      return jsonResponse({ ok: true, photoUrl: photo.photoUrl, photoFileId: photo.photoFileId, photoUpdatedAt: photo.photoUpdatedAt });
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

/* ============== JOB DESCRIPTION (Smart JD, F-PM-25) ==============
 * Sheet "job_descriptions": same id/json/updatedAt shape as "employees",
 * plus read-only summary columns (D..H) so HR can scan the sheet directly:
 *   D: ตำแหน่งงาน  E: ฝ่าย  F: แผนก  G: สถานะ  H: ลิงก์ PDF ล่าสุด
 * Each row's JSON is the full Smart JD record (see README.md / the app's
 * jd-edit screen for the field list): title/division/department/section,
 * the 7 form sections, qualifications, signers, revision/effectiveDate,
 * and — once generated — pdfUrl/pdfFileId/pdfUpdatedAt.
 */
var JD_SUMMARY_HEADERS = ['ตำแหน่งงาน', 'ฝ่าย', 'แผนก', 'สถานะ', 'ลิงก์ PDF ล่าสุด'];

function getJdSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(JD_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(JD_SHEET_NAME);
    sheet.appendRow(['id', 'json', 'updatedAt'].concat(JD_SUMMARY_HEADERS));
  }
  return sheet;
}

function jdSummaryRow_(data) {
  var statusLabel = { draft: 'ฉบับร่าง', pending: 'รออนุมัติ', approved: 'อนุมัติแล้ว' }[data.status] || (data.status || '');
  return [
    data.title || '',
    data.division || '',
    data.department || '',
    statusLabel,
    data.pdfUrl || ''
  ];
}

function upsertJdRow_(id, data) {
  if (!id) throw new Error('missing id');
  var sheet = getJdSheet_();
  var row = findRow_(sheet, id);
  var json = JSON.stringify(data);
  var now = new Date().toISOString();
  var summary = jdSummaryRow_(data);
  if (row === -1) {
    sheet.appendRow([id, json, now].concat(summary));
  } else {
    sheet.getRange(row, 2).setValue(json);
    sheet.getRange(row, 3).setValue(now);
    sheet.getRange(row, 4, 1, summary.length).setValues([summary]);
  }
}

function listJobDescriptions() {
  var sheet = getJdSheet_();
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
      // skip a row that doesn't contain valid JSON
    }
  }
  return out;
}

/* ---- PDF generation + Google Drive storage, grouped by ฝ่าย/แผนก ---- */

function getOrCreateSubfolder_(parent, name) {
  name = (name || 'ไม่ระบุ').toString().trim() || 'ไม่ระบุ';
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function getJdRootFolder_() {
  var it = DriveApp.getFoldersByName(JD_DRIVE_ROOT_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(JD_DRIVE_ROOT_FOLDER_NAME);
}

// Drops the ALLIED-TEK logo into a table cell, sized to fit the header
// strip. Falls back to plain bold text if the fetch fails (e.g. the site
// is briefly unreachable) so PDF generation never breaks on this.
function insertLogo_(cell) {
  try {
    var blob = UrlFetchApp.fetch(ALLIED_TEK_LOGO_URL, { muteHttpExceptions: true }).getBlob();
    var img = cell.getChild(0).asParagraph().appendInlineImage(blob);
    var w = img.getWidth(), h = img.getHeight();
    var maxW = 130;
    if (w > maxW) { img.setWidth(maxW); img.setHeight(Math.round(h * (maxW / w))); }
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  } catch (err) {
    var par = cell.getChild(0).asParagraph();
    par.setText('ALLIED-TEK').setBold(true).setFontSize(13);
    par.editAsText().setForegroundColor('#C0272D');
    par.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
}

function sanitizeFileNamePart_(s) {
  return String(s || '').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();
}

/* ---- Employee profile photo upload ----
 * Employees upload their own photo from the Employee Portal (cropped to a
 * square on the client before sending). Stored in Drive under a single
 * "HROD - Profile Photos" folder, named by employee id so re-uploads are
 * easy to find; the previous file is removed when a new one replaces it so
 * the folder doesn't accumulate stale copies.
 * Returns a lh3.googleusercontent.com thumbnail URL rather than a
 * drive.google.com/uc link — the former renders reliably inside an <img>
 * tag without the occasional interstitial "can't scan for viruses" page
 * drive.google.com shows for some file types.
 */
function uploadEmployeePhoto_(id, filename, mimeType, base64) {
  if (!id) throw new Error('missing id');
  if (!base64) throw new Error('missing image data');
  var folder = getOrCreateSubfolder_(DriveApp.getRootFolder(), PHOTO_DRIVE_ROOT_FOLDER_NAME);

  // Remove any previous photo file(s) for this employee so the folder
  // doesn't accumulate old copies every time someone re-uploads.
  var prefix = 'profile_' + sanitizeFileNamePart_(id);
  var existing = folder.getFiles();
  while (existing.hasNext()) {
    var f = existing.next();
    if (f.getName().indexOf(prefix) === 0) f.setTrashed(true);
  }

  var bytes = Utilities.base64Decode(base64);
  var ext = (mimeType && mimeType.indexOf('png') !== -1) ? '.png' : '.jpg';
  var blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', prefix + ext);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var photoFileId = file.getId();
  var photoUrl = 'https://lh3.googleusercontent.com/d/' + photoFileId + '=s400';
  var now = new Date().toISOString();

  // Persist onto the employee's own record (merge, don't clobber other
  // fields saved around the same time).
  var sheet = getSheet_();
  var row = findRow_(sheet, id);
  var data = {};
  if (row !== -1) {
    try { data = JSON.parse(sheet.getRange(row, 2).getValue()) || {}; } catch (e) { data = {}; }
  }
  data.photoUrl = photoUrl;
  data.photoFileId = photoFileId;
  data.photoUpdatedAt = now;
  upsertRow_(sheet, id, data);

  return { photoUrl: photoUrl, photoFileId: photoFileId, photoUpdatedAt: now };
}

// Builds a Google Doc that mirrors the F-PM-25 "Smart Job Description" table
// layout (header block, 7 numbered sections, signature block), exports it
// as a PDF, uploads the PDF into Drive under ฝ่าย/แผนก, deletes the
// scratch Doc, stamps the JD record with the new file's link, and returns
// the summary used by the client + written back into the sheet row.
function generateJdPdf_(id, data) {
  if (!id) throw new Error('missing id');
  data = data || {};

  var doc = DocumentApp.create('tmp-jd-' + id);
  var body = doc.getBody();
  body.setMarginTop(28).setMarginBottom(28).setMarginLeft(28).setMarginRight(28);
  var FONT = 'Sarabun';

  function styleCell(cell, opts) {
    opts = opts || {};
    var par = cell.getChild(0).asParagraph();
    par.setFontFamily(FONT).setFontSize(opts.size || 10);
    if (opts.bold) par.setBold(true);
    if (opts.align) par.setAlignment(opts.align);
    if (opts.color) par.editAsText().setForegroundColor(opts.color);
    if (opts.bg) cell.setBackgroundColor(opts.bg);
    cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(5).setPaddingRight(5);
    return par;
  }

  // Header: Document / title / Revision / Effective date / Page
  var head = body.appendTable([
    ['', 'Document\nF-PM-25', 'รายละเอียดงาน (Smart Job Description)', 'Revision:\n' + (data.revision || '04')],
    ['', 'Effective Date:\n' + (data.effectiveDate || '02 Apr 2025'), '', 'Page 1 of 1']
  ]);
  for (var r = 0; r < head.getNumRows(); r++) {
    for (var c = 0; c < head.getRow(r).getNumCells(); c++) {
      styleCell(head.getRow(r).getCell(c), { size: 9, bold: (c !== 2), align: c === 2 ? DocumentApp.HorizontalAlignment.CENTER : null });
    }
  }
  var logoCell = head.getRow(0).getCell(0);
  logoCell.setBackgroundColor('#ffffff');
  insertLogo_(logoCell);

  styleCell(body.appendTable([[
    'แผนก : ' + (data.department || ''), 'ส่วน : ' + (data.section || ''), 'ฝ่าย : ' + (data.division || '')
  ]]).getRow(0).getCell(0), { size: 9 });

  var posTable = body.appendTable([['ส่วนประกอบของ Smart Job Description', 'ตำแหน่งงาน']]);
  styleCell(posTable.getRow(0).getCell(0), { size: 9, bold: true, bg: '#f3f3f3' });
  styleCell(posTable.getRow(0).getCell(1), { size: 9, bold: true, bg: '#f3f3f3', align: DocumentApp.HorizontalAlignment.CENTER });
  var posRow = body.appendTable([['', data.title || '']]);
  styleCell(posRow.getRow(0).getCell(1), { size: 10, bold: true, align: DocumentApp.HorizontalAlignment.CENTER });

  function sectionHeader(label) {
    var t = body.appendTable([[label]]);
    styleCell(t.getRow(0).getCell(0), { size: 9.5, bold: true, bg: '#1a1a1a', color: '#ffffff' });
  }
  function sectionRows(list) {
    (list && list.length ? list : ['']).forEach(function (line) {
      var t = body.appendTable([[line || '']]);
      styleCell(t.getRow(0).getCell(0), { size: 10 });
    });
  }
  function kvRows(pairs) {
    pairs.forEach(function (pair) {
      var t = body.appendTable([[pair[0], pair[1] || '']]);
      styleCell(t.getRow(0).getCell(0), { size: 10, bold: true });
      styleCell(t.getRow(0).getCell(1), { size: 10 });
    });
  }

  sectionHeader('ส่วนที่ 1 : วัตถุประสงค์ที่บริษัทมีหน่วยงานและตำแหน่งงาน (Objectives)');
  sectionRows(data.objectives);

  sectionHeader('ส่วนที่ 2 : หน้าที่ความรับผิดชอบหลัก (Key responsibilities)');
  sectionRows(data.responsibilities);

  sectionHeader('ส่วนที่ 3 : ตัวชี้วัดผลงานตามหน้าที่ความรับผิดชอบ (Performance Indicators)');
  sectionRows(data.kpis);

  sectionHeader('ส่วนที่ 4 : คุณสมบัติผู้ดำรงตำแหน่ง (ต้องมีก่อนสมัครงาน)');
  var q = data.qualifications || {};
  kvRows([
    ['อายุตัว', q.age], ['วุฒิการศึกษา', q.education], ['สาขา', q.major],
    ['ประสบการณ์ทำงาน', q.experience], ['ใบอนุญาตต่างๆ', q.licenses], ['อื่นๆ', q.others]
  ]);

  sectionHeader('ส่วนที่ 5 : บทบาทผู้ดำรงตำแหน่ง');
  kvRows([['5.1 อำนาจดำเนินการ', data.authority], ['5.2 สายบังคับบัญชา', data.chain]]);

  sectionHeader('ส่วนที่ 6 : ความรู้ ทักษะและคุณลักษณะที่จำเป็นต้องมีคือ (Job Competencies)');
  var skillHead = body.appendTable([['6.1 ความรู้/ทักษะ (Knowledge & Skills/Functional competencies)']]);
  styleCell(skillHead.getRow(0).getCell(0), { size: 9.5, bold: true, bg: '#f7f7f7' });
  sectionRows(data.knowledgeSkills);
  var persHead = body.appendTable([['6.2 คุณลักษณะ (Personality/Attitude/Mental abilities)']]);
  styleCell(persHead.getRow(0).getCell(0), { size: 9.5, bold: true, bg: '#f7f7f7' });
  sectionRows(data.personality);

  sectionHeader('ส่วนที่ 7 : การนำ Job Description ไปใช้งาน');
  sectionRows(data.usage);

  // Signature block
  var s = data.signers || {};
  function sigCell(label, name, position) {
    return label + '\n\nลงชื่อ.......................\n(' + (name || '                    ') + ')\nตำแหน่ง ' + (position || '.......................') + '\nวันที่ .......................';
  }
  var sigTable = body.appendTable([[
    sigCell('พนักงาน ลงชื่อรับทราบ', s.employeeName, ''),
    sigCell('ผู้จัดการแผนก (ผู้จัดทำ)', s.deptManagerName, s.deptManagerPosition),
    sigCell('ผู้จัดการส่วน', s.sectionManagerName, s.sectionManagerPosition),
    sigCell('ผู้จัดการฝ่าย (ผู้อนุมัติ)', s.divisionManagerName, s.divisionManagerPosition),
    sigCell('แผนกบุคคล รับทราบ', s.hrName, s.hrPosition)
  ]]);
  for (var sc = 0; sc < sigTable.getRow(0).getNumCells(); sc++) {
    styleCell(sigTable.getRow(0).getCell(sc), { size: 8, align: DocumentApp.HorizontalAlignment.CENTER });
  }

  doc.saveAndClose();

  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  var root = getJdRootFolder_();
  var divisionFolder = getOrCreateSubfolder_(root, data.division);
  var deptFolder = getOrCreateSubfolder_(divisionFolder, data.department);
  var fileName = 'F-PM-25_' + sanitizeFileNamePart_(data.title || id) + '_Rev' + sanitizeFileNamePart_(data.revision || '04') + '.pdf';

  // Replace a previous PDF of the same name so re-generating doesn't pile up copies.
  var existing = deptFolder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  var pdfFile = deptFolder.createFile(pdfBlob).setName(fileName);
  DriveApp.getFileById(doc.getId()).setTrashed(true); // drop the scratch Doc, keep only the PDF

  var now = new Date().toISOString();
  var merged = Object.assign({}, data, {
    pdfUrl: pdfFile.getUrl(),
    pdfFileId: pdfFile.getId(),
    pdfUpdatedAt: now
  });
  upsertJdRow_(id, merged);

  return { pdfUrl: pdfFile.getUrl(), pdfFileId: pdfFile.getId(), pdfUpdatedAt: now };
}
