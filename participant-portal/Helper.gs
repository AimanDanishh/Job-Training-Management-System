/**
 * Helper.gs — Shared database connection, sheet access, script properties, and response wrappers
 * 
 * Reuses existing Google Spreadsheet DB without modifying database schema or sheet names.
 */

// ─── Script Properties & Config ───────────────────────────────────────────────
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

const SHEET_NAMES = {
  get employees()            { return getConfigProperty('SHEET_EMPLOYEES', 'Employees'); },
  get trainings()            { return getConfigProperty('SHEET_TRAININGS', 'Trainings'); },
  get trainingSessions()     { return getConfigProperty('SHEET_TRAINING_SESSIONS', 'TrainingSessions'); },
  get attendance()           { return getConfigProperty('SHEET_ATTENDANCE', 'Attendance'); },
  get trainingEval()         { return getConfigProperty('SHEET_TRAINING_EVAL', 'TrainingEval'); },
  get postEval()             { return getConfigProperty('SHEET_POST_EVAL', 'PostEval'); },
  get trainingParticipants() { return getConfigProperty('SHEET_TRAINING_PARTICIPANTS', 'TrainingParticipants'); }
};

// ─── Spreadsheet Access ────────────────────────────────────────────────────────
let _cachedSpreadsheet = null;

function getSpreadsheet() {
  if (_cachedSpreadsheet) return _cachedSpreadsheet;
  const spreadsheetId = getConfigProperty('SPREADSHEET_ID', '');
  if (spreadsheetId) {
    _cachedSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    return _cachedSpreadsheet;
  }
  _cachedSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return _cachedSpreadsheet;
}

function getEmployeeSpreadsheetId() {
  return getConfigProperty('EMPLOYEE_SPREADSHEET_ID', getConfigProperty('SPREADSHEET_ID', ''));
}

let _cachedEmployeeSpreadsheet = null;
function getEmployeeSpreadsheet() {
  if (_cachedEmployeeSpreadsheet) return _cachedEmployeeSpreadsheet;
  const empSpreadsheetId = getEmployeeSpreadsheetId();
  if (empSpreadsheetId) {
    try {
      _cachedEmployeeSpreadsheet = SpreadsheetApp.openById(empSpreadsheetId);
      return _cachedEmployeeSpreadsheet;
    } catch(e) {
      Logger.log('Failed to open separate EMPLOYEE_SPREADSHEET_ID: ' + e.message);
    }
  }
  return getSpreadsheet();
}

function getSheet(name) {
  const ss = (name === SHEET_NAMES.employees) ? getEmployeeSpreadsheet() : getSpreadsheet();
  if (!ss) return null;
  
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    const allSheets = ss.getSheets();
    const targetClean = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
    sheet = allSheets.find(s => {
      const sClean = s.getName().toLowerCase().replace(/[^a-z0-9]/g, '');
      return sClean === targetClean ||
             sClean === targetClean + 's' ||
             sClean + 's' === targetClean;
    });

    if (!sheet && name === SHEET_NAMES.employees && ss !== getSpreadsheet() && allSheets.length > 0) {
      sheet = allSheets[0];
    }
    if (!sheet && name === SHEET_NAMES.employees && allSheets.length > 0) {
      sheet = allSheets.find(s => s.getName().toLowerCase().includes('emp') || s.getName().toLowerCase().includes('staff'));
    }
  }

  if (sheet) {
    const cleanName = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
    if (cleanName === 'attendance') {
      ensureAttendanceSheetColumns(sheet);
    }
  }

  return sheet;
}

/**
 * Resolves and opens the single per-training Google Sheet containing tabs:
 * TrainingParticipants, TrainingSessions, Attendance, TrainingEval, PostEval, Summary
 *
 * @param {string} trainingId - Training ID (e.g. TRN-1001) or Training Code (e.g. LM-2026-0001)
 * @returns {Spreadsheet|null} Google Spreadsheet object for the training, or null
 */
function getTrainingDataSpreadsheet(trainingId) {
  if (!trainingId) return null;
  const cleanId = String(trainingId).trim();

  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) return null;

  const trainings = sheetToJson(tSheet);
  const t = trainings.find(r => String(r.ID || r.TrainingID || r.Code || '').trim() === cleanId);
  if (!t) return null;

  // 1. Try sheet ID stored in ParticipantsSheetID, SessionsSheetID, AttendanceSheetID, etc.
  const sheetId = t.ParticipantsSheetID || t.AttendanceSheetID || t.SessionsSheetID || t.EvaluationSheetID || t.PostSheetID;
  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log('Error opening per-training sheet by ID (' + sheetId + '): ' + e.message);
    }
  }

  // 2. Try FolderID
  if (t.FolderID) {
    try {
      const folder = DriveApp.getFolderById(t.FolderID);
      const code = t.Code || t.ID;
      const fileIter = folder.getFilesByName(`${code} Training Data`);
      if (fileIter.hasNext()) {
        return SpreadsheetApp.openById(fileIter.next().getId());
      } else {
        const file = getOrCreateSingleTrainingSheet(folder, code);
        return SpreadsheetApp.openById(file.getId());
      }
    } catch (e) {
      Logger.log('Error opening per-training sheet from FolderID: ' + e.message);
    }
  }

  return null;
}

function getOrCreateSingleTrainingSheet(folder, code) {
  const fileName = `${code} Training Data`;
  let fileIter = folder.getFilesByName(fileName);
  let file;
  let ss;

  if (fileIter.hasNext()) {
    file = fileIter.next();
    ss = SpreadsheetApp.openById(file.getId());
  } else {
    let legacyIter = folder.getFilesByName(`${code} Attendance Sheet`);
    if (legacyIter.hasNext()) {
      file = legacyIter.next();
      ss = SpreadsheetApp.openById(file.getId());
    } else {
      ss = SpreadsheetApp.create(fileName);
      file = DriveApp.getFileById(ss.getId());
      folder.addFile(file);
      try { DriveApp.getRootFolder().removeFile(file); } catch(rErr) {}
    }
  }

  const tabDefs = [
    {
      name: 'TrainingParticipants',
      headers: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt']
    },
    {
      name: 'TrainingSessions',
      headers: ['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate']
    },
    {
      name: 'Attendance',
      headers: ['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName', 'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours', 'Remarks', 'EditedBy', 'EditedAt']
    },
    {
      name: 'TrainingEval',
      headers: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt']
    },
    {
      name: 'PostEval',
      headers: ['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID', 'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply', 'FurtherTraining', 'Comments', 'SubmittedAt']
    }
  ];

  tabDefs.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      const allSheets = ss.getSheets();
      sheet = allSheets.find(s => s.getName().toLowerCase().includes(def.name.toLowerCase()));
      if (!sheet) {
        sheet = ss.insertSheet(def.name);
      } else {
        sheet.setName(def.name);
      }
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(def.headers);
      sheet.getRange(1, 1, 1, def.headers.length)
        .setFontWeight('bold')
        .setBackground('#2563EB')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  });

  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Data');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() <= 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  let summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('Summary');
    summarySheet.getRange('A1:B1').setValues([['Metric / Category', 'Live Formula / Output']]);
    summarySheet.getRange('A1:B1').setFontWeight('bold').setBackground('#1E293B').setFontColor('#FFFFFF');

    summarySheet.getRange('A2:B8').setFormulas([
      ['Total Enrolled Participants', '=IF(ISREF(TrainingParticipants!A2), COUNTA(TrainingParticipants!A2:A), 0)'],
      ['Total Sessions Created', '=IF(ISREF(TrainingSessions!A2), COUNTA(TrainingSessions!A2:A), 0)'],
      ['Total Attendance Logs', '=IF(ISREF(Attendance!A2), COUNTA(Attendance!A2:A), 0)'],
      ['Present Count', '=IF(ISREF(Attendance!H2), COUNTIF(Attendance!H2:H, "Present"), 0)'],
      ['Total Evaluations Submitted', '=IF(ISREF(TrainingEval!A2), COUNTA(TrainingEval!A2:A), 0)'],
      ['Overall Average Score', '=IF(AND(ISREF(TrainingEval!O2), COUNTA(TrainingEval!O2:O)>0), AVERAGE(TrainingEval!O2:O), 0)'],
      ['Total Post-Reviews Completed', '=IF(ISREF(PostEval!A2), COUNTA(PostEval!A2:A), 0)']
    ]);
  }

  SpreadsheetApp.flush();
  return file;
}

/**
 * Helper to look up a training session across all per-training sheets
 * 
 * @param {string} sessionId - Session ID (e.g. SES0001)
 * @returns {Object|null} { session: Object, training: Object, spreadsheet: Spreadsheet, sessionSheet: Sheet }
 */
function findTrainingBySessionId(sessionId) {
  if (!sessionId) return null;
  const cleanSessionId = String(sessionId).trim();

  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) return null;

  const trainings = sheetToJson(tSheet);
  for (const t of trainings) {
    if (!t.ID) continue;
    const ss = getTrainingDataSpreadsheet(t.ID);
    if (!ss) continue;
    const sessSheet = ss.getSheetByName('TrainingSessions');
    if (!sessSheet) continue;
    const sessions = sheetToJson(sessSheet);
    const session = sessions.find(s => String(s.SessionID || '').trim() === cleanSessionId);
    if (session) {
      return { session: session, training: t, spreadsheet: ss, sessionSheet: sessSheet };
    }
  }
  return null;
}

// ─── ID Generation & Header Normalization ─────────────────────────────────────
function generateId(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
}

function normalizeHeader(header) {
  const h = String(header).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  if (['id', 'empid', 'employeeid', 'employeeno', 'staffid', 'badgenumber'].includes(h)) return 'ID';
  if (['name', 'fullname', 'employeename', 'staffname'].includes(h)) return 'Name';
  if (['costcentre', 'costcenter'].includes(h)) return 'CostCentre';
  if (['department', 'dept', 'company', 'division'].includes(h)) return 'Department';
  if (['position', 'positiontitle', 'jobtitle', 'title', 'designation', 'role'].includes(h)) return 'Position';
  if (['email', 'emailaddress'].includes(h)) return 'Email';
  if (['status'].includes(h)) return 'Status';
  return header;
}

/**
 * Convert sheet data rows into JSON objects keyed by column headers
 */
function sheetToJson(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const isRowEmpty = data[i].every(val => val === '' || val === null || val === undefined);
    if (isRowEmpty) continue;

    const obj = {};
    headers.forEach((h, j) => {
      const cleanH = String(h).trim();
      const val = data[i][j] !== undefined ? String(data[i][j]) : '';
      obj[cleanH] = val;

      const normKey = cleanH.replace(/[^a-zA-Z0-9]/g, '');
      if (normKey && !obj[normKey]) {
        obj[normKey] = val;
      }
      const customNormKey = normalizeHeader(cleanH);
      if (customNormKey && !obj[customNormKey]) {
        obj[customNormKey] = val;
      }
    });
    if (!obj.ID && data[i][0] !== undefined) obj.ID = String(data[i][0]);
    if (!obj.Name && data[i][1] !== undefined) obj.Name = String(data[i][1]);

    // Smart fallback for shifted columns
    if (obj.TrainingID && String(obj.TrainingID).trim().toUpperCase().startsWith('SES')) {
      obj.SessionID = obj.TrainingID;
      if (obj.TrainingCode && String(obj.TrainingCode).trim().toUpperCase().startsWith('TRN')) {
        obj.TrainingID = obj.TrainingCode;
      }
    }
    if (!obj.Status && (obj.Date === 'Present' || obj.Date === 'Absent' || obj.Date === 'Late')) {
      obj.Status = obj.Date;
    }
    if (!obj.ScanTime && obj.Day && String(obj.Day).includes(':')) {
      obj.ScanTime = obj.Day;
    }

    obj._row = i + 1;
    rows.push(obj);
  }
  return rows;
}

/** Repair & standardize Attendance sheet columns and legacy shifted data rows. */
function ensureAttendanceSheetColumns(sheet) {
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return;

  const targetHeaders = [
    'AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName',
    'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day',
    'Date', 'Hours', 'Remarks', 'EditedBy', 'EditedAt'
  ];

  // 1. Ensure Row 1 has standard modern headers if old headers exist
  const firstRowStr = data[0].join(',').toLowerCase();
  if (!firstRowStr.includes('sessionid')) {
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders])
      .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  // 2. Fix shifted legacy rows where Col 2 contains SES... and Col 3 contains TRN...
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const col2 = String(row[1] || '').trim();
      const col3 = String(row[2] || '').trim();

      if (col2.toUpperCase().startsWith('SES') && col3.toUpperCase().startsWith('TRN')) {
        const fixedRow = [
          row[0] || generateId('ATT'),                     // AttendanceID
          col2,                                           // SessionID (SES0003)
          col3,                                           // TrainingID (TRN-xxx)
          row[3] || '',                                    // EmployeeNo
          row[4] || '',                                    // EmployeeName
          row[5] || '',                                    // Department
          row[6] || row[14] || now(),                      // ScanTime
          row[7] === 'Present' || row[7] === 'Late' || row[7] === 'Absent' ? row[7] : 'Present', // Status
          row[8] || '',                                    // TrainingCode
          row[9] || '',                                    // Day
          row[10] || '',                                   // Date
          row[11] || 0,                                    // Hours
          row[12] || 'QR Code Public Check-In',            // Remarks
          row[13] || 'Public Portal',                      // EditedBy
          row[14] || now()                                 // EditedAt
        ];
        sheet.getRange(i + 1, 1, 1, fixedRow.length).setValues([fixedRow]);
      }
    }
  }
}

function findRowById(sheet, id) {
  if (!sheet) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) return i + 1;
  }
  return -1;
}

// ─── Date & Utilities ─────────────────────────────────────────────────────────
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

// ─── Standardized JSON Response Helpers ────────────────────────────────────────
function ok(data) {
  return JSON.stringify({ success: true, data: data });
}

function err(message) {
  return JSON.stringify({ success: false, message: message });
}
