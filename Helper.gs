/**
 * Helper.gs — Shared utilities, sheet access, script properties, and initial setup
 */

// ─── Script Properties Configuration ──────────────────────────────────────────
/**
 * Utility to retrieve configuration values from Apps Script Project Settings (Script Properties).
 * Falls back to default value if property is not set.
 */
function getConfigProperty(key, defaultValue) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    if (value !== null && value !== undefined && String(value).trim() !== '') {
      return String(value).trim();
    }
  } catch (e) {
    Logger.log('Error reading script property ' + key + ': ' + e.message);
  }
  return defaultValue;
}

/**
 * Utility to set a configuration value in Script Properties.
 */
function setConfigProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * Run initDefaultScriptProperties() once from Apps Script editor to auto-populate default Script Properties.
 */
function initDefaultScriptProperties() {
  const defaults = {
    'SPREADSHEET_ID':          '',
    'ALLOWED_DOMAIN':          'company.com',
    'ADMIN_EMAILS':            'admin@company.com',
    'ADMIN_USER':              'admin',
    'ADMIN_PASS':              'admin123',
    'APP_TITLE':               'TrainHub — Training Management System',
    'SHEET_EMPLOYEES':         'Employees',
    'SHEET_TRAININGS':         'Trainings',
    'SHEET_ATTENDANCE':        'Attendance',
    'SHEET_TRAINING_EVAL':     'TrainingEval',
    'SHEET_POST_EVAL':         'PostEval',
    'ROOT_FOLDER_ID':          '',
    'TEMPLATE_FOLDER_ID':      '',
    'ATTENDANCE_TEMPLATE_ID':  '',
    'EVALUATION_TEMPLATE_ID':  '',
    'CERTIFICATE_TEMPLATE_ID': '',
    'REPORT_TEMPLATE_ID':      '',
    // Google Sheets master copy of AP-HRD-F01-00. Every new programme gets a populated copy.
    'TRAINING_REQUISITION_TEMPLATE_ID': ''
  };
  PropertiesService.getScriptProperties().setProperties(defaults, false);
  Logger.log('Default Script Properties successfully initialized in Project Settings.');
}

// ─── Dynamic Sheet Names & Spreadsheet ID ─────────────────────────────────────
function getSpreadsheetId() {
  return getConfigProperty('SPREADSHEET_ID', '');
}

const SHEET_NAMES = {
  get employees()    { return getConfigProperty('SHEET_EMPLOYEES', 'Employees'); },
  get trainings()    { return getConfigProperty('SHEET_TRAININGS', 'Trainings'); },
  get attendance()   { return getConfigProperty('SHEET_ATTENDANCE', 'Attendance'); },
  get trainingEval() { return getConfigProperty('SHEET_TRAINING_EVAL', 'TrainingEval'); },
  get postEval()     { return getConfigProperty('SHEET_POST_EVAL', 'PostEval'); }
};

// ─── Spreadsheet Access ────────────────────────────────────────────────────────
function getSpreadsheet() {
  const spreadsheetId = getSpreadsheetId();
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheetHeaders(sheet, name);
  }
  return sheet;
}

// ─── Sheet Headers ─────────────────────────────────────────────────────────────
function initSheetHeaders(sheet, name) {
  const headers = {
    Employees:    ['ID', 'Name', 'Department', 'Position', 'Email', 'Phone', 'Status'],
    Trainings:    ['ID', 'Code', 'Name', 'Category', 'Trainer', 'Venue', 'StartDate',
                   'EndDate', 'Duration', 'TotalHours', 'Department', 'Objectives',
                   'Status', 'Stage', 'Participants',
                   'FolderID', 'AttendanceFolderID', 'EvaluationFolderID', 'CertificateFolderID',
                   'MaterialsFolderID', 'PhotosFolderID', 'ReportsFolderID', 'TrainerNotesFolderID',
                   'CreatedDate', 'UpdatedDate', 'CourseFee', 'RequisitionFormFileID'],
    Attendance:   ['ID', 'TrainingID', 'TrainingCode', 'EmployeeID', 'EmployeeName',
                   'Department', 'Day', 'Date', 'CheckIn', 'CheckOut', 'Hours',
                   'Status', 'Remarks', 'EditedBy', 'EditedAt'],
    TrainingEval: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName',
                   'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7',
                   'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt'],
    PostEval:     ['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID',
                   'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply',
                   'FurtherTraining', 'Comments', 'SubmittedAt']
  };

  const headerKey = Object.keys(headers).find(k => k === name || getConfigProperty('SHEET_' + k.toUpperCase(), k) === name);
  const matchedHeaders = headers[headerKey] || headers[name];

  if (matchedHeaders) {
    sheet.appendRow(matchedHeaders);
    sheet.getRange(1, 1, 1, matchedHeaders.length)
      .setFontWeight('bold')
      .setBackground('#2563EB')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
}

/** Adds fields introduced after the original training sheet was deployed. */
function ensureTrainingSheetColumns(sheet) {
  const requiredHeaders = ['CourseFee', 'RequisitionFormFileID'];
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header)
        .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      headers.push(header);
    }
  });
  return headers;
}

// ─── ID Generation ─────────────────────────────────────────────────────────────
function generateId(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
}

// ─── Sheet to JSON ─────────────────────────────────────────────────────────────
/**
 * Convert a sheet's data rows into an array of objects keyed by header row.
 */
function sheetToJson(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // skip empty rows
    const obj = {};
    headers.forEach((h, j) => {
      obj[h] = data[i][j] !== undefined ? String(data[i][j]) : '';
    });
    obj._row = i + 1; // 1-indexed sheet row number
    rows.push(obj);
  }
  return rows;
}

// ─── Find Row By ID ─────────────────────────────────────────────────────────────
function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// ─── Date Helpers ───────────────────────────────────────────────────────────────
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d)) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy');
}

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
}

// ─── Response Helpers ───────────────────────────────────────────────────────────
function ok(data) {
  return JSON.stringify({ success: true, data: data });
}

function err(message) {
  return JSON.stringify({ success: false, message: message });
}

// ─── Setup (run once) ───────────────────────────────────────────────────────────
/**
 * Run setupSheets() once from the Apps Script editor to initialise all sheets.
 */
function setupSheets() {
  Object.values(SHEET_NAMES).forEach(name => getSheet(name));
  Logger.log('All sheets initialised successfully.');
}
