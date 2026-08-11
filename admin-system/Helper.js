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
    'EMPLOYEE_SPREADSHEET_ID': '',
    'ALLOWED_DOMAIN':          'company.com',
    'ADMIN_EMAILS':            'admin@company.com',
    'ADMIN_USER':              'admin',
    'ADMIN_PASS':              'admin123',
    'APP_TITLE':               'TrainHub — Training Management System',
    'SHEET_EMPLOYEES':         'Employees',
    'SHEET_HR_EMAIL':          'HR email',
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
    // Google Sheets master copy of AP-HRD-F01-01. Every new programme gets a populated copy.
    'TRAINING_REQUISITION_TEMPLATE_ID': '',
    // Company logo Drive link or public image URL to embed in center of session QR codes
    'COMPANY_LOGO_URL':        '',
    // Public Participant Portal Apps Script Web App Deployment URL
    'PUBLIC_PORTAL_URL':       '',
    // Separate HOD Portal Web App Deployment URL
    'HOD_PORTAL_URL':          ''
  };
  PropertiesService.getScriptProperties().setProperties(defaults, false);
  Logger.log('Default Script Properties successfully initialized in Project Settings.');
}

function getSpreadsheetId() {
  return getConfigProperty('SPREADSHEET_ID', '');
}

function getEmployeeSpreadsheetId() {
  return getConfigProperty('EMPLOYEE_SPREADSHEET_ID', getSpreadsheetId());
}

function getCompanyLogoUrl() {
  return getConfigProperty('COMPANY_LOGO_URL', '');
}

function setCompanyLogoUrl(url) {
  setConfigProperty('COMPANY_LOGO_URL', url);
  Logger.log('COMPANY_LOGO_URL updated in Project Settings: ' + url);
  return 'COMPANY_LOGO_URL set to: ' + url;
}

function getPublicPortalUrl() {
  const portalUrl = getConfigProperty('PUBLIC_PORTAL_URL', '');
  if (portalUrl && portalUrl.trim() !== '') return portalUrl.trim();
  return getAppUrl();
}

function getHodPortalUrl() {
  return getConfigProperty('HOD_PORTAL_URL', '');
}

/**
 * Convert any Google Drive share link, preview link, or File ID into a direct public image CDN URL
 * suitable for external rendering engines like QuickChart.
 * 
 * @param {string} input - Google Drive link (e.g. https://drive.google.com/file/d/FILE_ID/view) or File ID or image URL
 * @returns {string} Direct public image URL (e.g. https://lh3.googleusercontent.com/d/FILE_ID)
 */
function convertDriveLinkToDirectImageUrl(input) {
  if (!input) return '';
  const str = String(input).trim();

  // If already a direct image host URL (not drive.google.com), return directly
  if (!str.includes('drive.google.com') && str.startsWith('http')) {
    return str;
  }

  let fileId = '';
  // Extract File ID from drive.google.com links
  const fileIdMatch = str.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (fileIdMatch && fileIdMatch[1]) {
    fileId = fileIdMatch[1];
  } else {
    const idParamMatch = str.match(/id=([a-zA-Z0-9_-]{25,})/);
    if (idParamMatch && idParamMatch[1]) {
      fileId = idParamMatch[1];
    } else if (/^[a-zA-Z0-9_-]{25,}$/.test(str)) {
      fileId = str;
    }
  }

  if (fileId) {
    // High performance Google Drive image CDN URL format
    return `https://lh3.googleusercontent.com/d/${fileId}`;
  }

  return str;
}

const SHEET_NAMES = {
  get employees()            { return getConfigProperty('SHEET_EMPLOYEES', 'Employees'); },
  get hrEmail()              { return getConfigProperty('SHEET_HR_EMAIL', 'HR email'); },
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
  const spreadsheetId = getSpreadsheetId();
  if (spreadsheetId) {
    _cachedSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    return _cachedSpreadsheet;
  }
  _cachedSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return _cachedSpreadsheet;
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
  const isEmployeeSpreadsheetSheet = (
    name === SHEET_NAMES.employees ||
    name === SHEET_NAMES.hrEmail ||
    name === 'HR email' ||
    name === 'HR Email' ||
    name === 'HOD email' ||
    name === 'HOD Email' ||
    name === 'For IT'
  );
  const ss = isEmployeeSpreadsheetSheet ? getEmployeeSpreadsheet() : getSpreadsheet();
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

    if (!sheet) {
      if (name === SHEET_NAMES.employees && ss !== getSpreadsheet() && allSheets.length > 0) {
        sheet = allSheets[0];
      } else {
        sheet = ss.insertSheet(name);
        initSheetHeaders(sheet, name);
      }
    }
  }

  // Auto-seed initial sample records if sheet is newly created or empty
  if (sheet && sheet.getLastRow() <= 1) {
    seedInitialSheetData(sheet, name);
  }

  // Ensure required columns exist for Trainings and TrainingSessions sheets
  if (sheet) {
    const cleanName = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
    if (cleanName === 'trainings') {
      ensureTrainingSheetColumns(sheet);
    } else if (cleanName === 'trainingsessions' || cleanName === 'sessions') {
      ensureTrainingSessionsSheetColumns(sheet);
    }
  }

  return sheet;
}

// ─── Sheet Headers & Sample Data ─────────────────────────────────────────────
function initSheetHeaders(sheet, name) {
  const headers = {
    Employees:        ['ID', 'Name', 'Department', 'Position', 'Email', 'Phone', 'Status'],
    'HR email':       ['Employee No', 'HR', 'Cost Centre', 'Position Title', 'Email'],
    Trainings:        ['ID', 'Code', 'Name', 'Category', 'Trainer', 'Venue', 'StartDate',
                       'EndDate', 'Duration', 'TotalHours', 'Department', 'Objectives',
                       'Status', 'Stage', 'Participants',
                       'FolderID', 'ParticipantsSheetID', 'SessionsSheetID', 'AttendanceSheetID',
                       'EvaluationSheetID', 'PostSheetID', 'RequisitionFormFileID',
                       'CreatedDate', 'UpdatedDate', 'CourseFee',
                       'ApprovalStatus', 'RequestedBy', 'RequestedByName', 'RequestedByEmail', 'RequestedDate', 'ApprovedBy', 'ApprovedCostCentre', 'ApprovedAt', 'ApprovalRemarks', 'RescheduledDate'],
    TrainingSessions: ['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate'],
    Attendance:       ['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName',
                       'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours',
                       'Remarks', 'EditedBy', 'EditedAt'],
    TrainingEval:     ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName',
                       'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7',
                       'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt'],
    PostEval:         ['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID',
                       'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply',
                       'FurtherTraining', 'Comments', 'SubmittedAt'],
    TrainingParticipants: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt']
  };

  const headerKey = Object.keys(headers).find(k => k === name || getConfigProperty('SHEET_' + k.toUpperCase(), k) === name);
  const matchedHeaders = headers[headerKey] || headers[name];

  if (matchedHeaders && sheet.getLastRow() === 0) {
    sheet.appendRow(matchedHeaders);
    sheet.getRange(1, 1, 1, matchedHeaders.length)
      .setFontWeight('bold')
      .setBackground('#2563EB')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
}

function seedInitialSheetData(sheet, name) {
  if (sheet.getLastRow() === 0) {
    initSheetHeaders(sheet, name);
  }
}

/**
 * Reads all records from the "HR email" tab in EMPLOYEE_SPREADSHEET_ID.
 * Header format: Employee No | HR | Cost Centre | Position Title | Email
 */
function getHrEmailRecords() {
  const getVal = (rowObj, nameList) => {
    if (!rowObj) return '';
    const keys = Object.keys(rowObj);
    for (let n of nameList) {
      const matchKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === n.toLowerCase().replace(/[^a-z0-9]/g, ''));
      if (matchKey && rowObj[matchKey] !== undefined && rowObj[matchKey] !== null) {
        return String(rowObj[matchKey]).trim();
      }
    }
    return '';
  };

  const idAliases = ['Employee No', 'EmployeeNo', 'EmployeeID', 'ID', 'EmpNo', 'Staff ID'];
  const nameAliases = ['HR', 'HRName', 'HR Name', 'Name', 'Employee Name'];
  const deptAliases = ['Cost Centre', 'CostCentre', 'Department', 'Dept'];
  const posAliases = ['Position Title', 'PositionTitle', 'Position', 'JobTitle'];
  const emailAliases = ['Email', 'EmailAddress', 'Email Address', 'HREmail'];

  const records = [];
  try {
    const hrSheet = getSheet(SHEET_NAMES.hrEmail || 'HR email');
    if (hrSheet) {
      const rows = sheetToJson(hrSheet);
      rows.forEach(r => {
        const email = getVal(r, emailAliases);
        if (email) {
          records.push({
            employeeNo: getVal(r, idAliases),
            name: getVal(r, nameAliases),
            costCentre: getVal(r, deptAliases),
            position: getVal(r, posAliases) || 'HR Department',
            email: email.toLowerCase().trim()
          });
        }
      });
    }
  } catch (e) {
    Logger.log('getHrEmailRecords error: ' + e.message);
  }
  return records;
}

/**
 * Resolves HR Profile by email address from "HR email" tab, falling back to "Employees" tab or default admin profile.
 */
function getHrProfileByEmail(emailInput) {
  let emailClean = String(emailInput || '').toLowerCase().trim();
  if (!emailClean) {
    try {
      emailClean = Session.getActiveUser().getEmail().toLowerCase().trim();
    } catch (e) {}
  }

  // 1. Search "HR email" tab
  const hrRecords = getHrEmailRecords();
  const hrMatch = hrRecords.find(r => r.email === emailClean);
  if (hrMatch) {
    return {
      employeeNo: hrMatch.employeeNo || 'HR-001',
      name: hrMatch.name || emailClean.split('@')[0].replace(/[\._]/g, ' ').replace(/\b\w/g, c => c.toUpperCase()),
      position: hrMatch.position || 'HR Department',
      costCentre: hrMatch.costCentre || '',
      email: hrMatch.email
    };
  }

  // 2. Fallback: Search "Employees" tab by Email
  try {
    const empSheet = getSheet(SHEET_NAMES.employees);
    if (empSheet) {
      const empRows = sheetToJson(empSheet);
      const empMatch = empRows.find(e => {
        const empEmail = String(e.Email || e['Email Address'] || '').toLowerCase().trim();
        return empEmail && empEmail === emailClean;
      });

      if (empMatch) {
        return {
          employeeNo: String(empMatch.ID || empMatch.EmployeeID || empMatch.EmployeeNo || '').trim(),
          name: String(empMatch.Name || empMatch.EmployeeName || '').trim(),
          position: String(empMatch.Position || empMatch.JobTitle || empMatch.PositionTitle || 'HR Department').trim(),
          costCentre: String(empMatch.Department || empMatch.CostCentre || '').trim(),
          email: emailClean
        };
      }
    }
  } catch (e) {
    Logger.log('getHrProfileByEmail Employees fallback error: ' + e.message);
  }

  // 3. Fallback: Default formatted profile for configured ADMIN_EMAILS
  const defaultName = emailClean ? emailClean.split('@')[0].replace(/[\._]/g, ' ').replace(/\b\w/g, c => c.toUpperCase()) : 'HR Department';
  return {
    employeeNo: 'ADMIN',
    name: defaultName,
    position: 'HR Department',
    costCentre: '',
    email: emailClean
  };
}

/** Adds fields introduced after the original training sheet was deployed. */
function ensureTrainingSheetColumns(sheet) {
  if (!sheet) return [];
  const requiredHeaders = [
    'ID', 'Code', 'Name', 'Category', 'Trainer', 'Venue', 'StartDate',
    'EndDate', 'Duration', 'TotalHours', 'Department', 'Objectives',
    'Status', 'Stage', 'Participants',
    'FolderID', 'ParticipantsSheetID', 'SessionsSheetID', 'AttendanceSheetID',
    'EvaluationSheetID', 'PostSheetID', 'RequisitionFormFileID',
    'CreatedDate', 'UpdatedDate', 'CourseFee',
    'ApprovalStatus', 'RequestedBy', 'RequestedByName', 'RequestedByEmail', 'RequestedDate', 'ApprovedBy', 'ApprovedCostCentre', 'ApprovedAt', 'ApprovalRemarks', 'RescheduledDate', 'BrochureURL'
  ];
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(header)
        .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      headers.push(header);
    }
  });
  return headers;
}

/** Adds/repairs fields in TrainingSessions sheet if missing. */
function ensureTrainingSessionsSheetColumns(sheet) {
  if (!sheet) return [];
  const requiredHeaders = ['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate'];
  
  if (sheet.getLastRow() === 0) {
    initSheetHeaders(sheet, 'TrainingSessions');
    return requiredHeaders;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const cleanHeaderStrings = headers.map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));

  let headersModified = false;

  requiredHeaders.forEach(reqHeader => {
    const cleanReq = reqHeader.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (!cleanHeaderStrings.includes(cleanReq)) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(reqHeader)
        .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      headers.push(reqHeader);
      cleanHeaderStrings.push(cleanReq);
      headersModified = true;
    }
  });

  if (headersModified) {
    sheet.setFrozenRows(1);
  }

  return headers;
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
      if (typeof getOrCreateSingleTrainingSheet === 'function') {
        const file = getOrCreateSingleTrainingSheet(folder, code);
        return SpreadsheetApp.openById(file.getId());
      }
    } catch (e) {
      Logger.log('Error opening per-training sheet from FolderID: ' + e.message);
    }
  }

  return null;
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

// ─── ID Generation ─────────────────────────────────────────────────────────────
function generateId(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
}

function normalizeEmployeeHeader(header) {
  const h = String(header).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  if (['id', 'empid', 'employeeid', 'employeeno', 'staffid', 'badgenumber', 'no', 'nokp', 'ic'].includes(h)) return 'ID';
  if (['name', 'fullname', 'employeename', 'staffname', 'nama'].includes(h)) return 'Name';
  if (['department', 'dept', 'costcentre', 'company', 'division', 'section', 'jabatan'].includes(h)) return 'Department';
  if (['position', 'positiontitle', 'jobtitle', 'title', 'designation', 'role', 'jawatan', 'jobcategory'].includes(h)) return 'Position';
  if (['email', 'emailaddress', 'emel'].includes(h)) return 'Email';
  if (['phone', 'mobile', 'contact', 'contactnumber', 'telefon'].includes(h)) return 'Phone';
  if (['status', 'employmentstatus', 'employmenttype', 'stat'].includes(h)) return 'Status';
  return header;
}

/** Authoritative Employee-sheet directory used for every participant write. */
function getOfficialEmployeeDirectory() {
  const byId = {};
  const byName = {};
  const sheet = getSheet(SHEET_NAMES.employees);
  const rows = sheet ? sheetToJson(sheet) : [];

  rows.forEach(row => {
    const id = String(row.ID || row.EmployeeID || row.EmployeeNo || row.EmpID || row.StaffID || row['Employee ID'] || row['Employee No'] || row['Staff ID'] || '').trim();
    const name = String(row.Name || row.EmployeeName || row['Employee Name'] || row['Staff Name'] || '').trim();
    if (!id) return;
    const employee = {
      ID: id,
      EmployeeID: id,
      Name: name,
      EmployeeName: name,
      Department: String(row.CostCentre || row['Cost Centre'] || row.Department || row.Dept || '').trim(),
      Position: String(row.PositionTitle || row['Position Title'] || row.Position || row.JobTitle || row['Job Title'] || '').trim(),
      Email: String(row.Email || row['Email Address'] || '').trim()
    };
    byId[id.toLowerCase()] = employee;
    if (name) {
      const nameKey = name.toLowerCase();
      byName[nameKey] = Object.prototype.hasOwnProperty.call(byName, nameKey) ? null : employee;
    }
  });
  return { byId: byId, byName: byName };
}

function canonicalizeTrainingParticipants(participants) {
  const directory = getOfficialEmployeeDirectory();
  const resolved = [];
  const rejected = [];
  const seen = {};

  (Array.isArray(participants) ? participants : []).forEach(participant => {
    const rawId = String(participant.ID || participant.EmployeeID || participant.EmployeeNo || participant.EmpID || participant['Employee ID'] || participant['Employee No'] || '').trim();
    const rawName = String(participant.Name || participant.EmployeeName || participant['Employee Name'] || '').trim();
    const employee = (rawId && directory.byId[rawId.toLowerCase()]) || (rawName && directory.byName[rawName.toLowerCase()]);
    if (!employee) {
      rejected.push(rawId || rawName || 'blank participant');
      return;
    }
    const key = employee.ID.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      resolved.push(employee);
    }
  });
  return { participants: resolved, rejected: rejected };
}

var _sheetDataCache = {};

function sheetToJson(sheet) {
  if (!sheet) return [];

  const sheetName = (sheet && typeof sheet.getName === 'function') ? sheet.getName() : '';
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const isRowEmpty = data[i].every(val => val === '' || val === null || val === undefined);
    if (isRowEmpty) continue;

    const obj = {};
    headers.forEach((h, j) => {
      const cleanH = String(h).trim();
      const normKey = normalizeEmployeeHeader(cleanH);
      const val = data[i][j] !== undefined ? String(data[i][j]) : '';
      obj[cleanH] = val;
      if (normKey && !obj[normKey]) {
        obj[normKey] = val;
      }
    });
    if (!obj.ID && data[i][0] !== undefined) obj.ID = String(data[i][0]);
    if (!obj.Name && data[i][1] !== undefined) obj.Name = String(data[i][1]);

    obj._row = i + 1; // 1-indexed sheet row number
    rows.push(obj);
  }

  if (sheetName) {
    _sheetDataCache[sheetName] = rows;
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
  if (isNaN(d.getTime())) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/**
 * Returns current date formatted as dd/MM/yyyy in project script timezone.
 */
function getFormattedCurrentDate(date) {
  const d = date ? new Date(date) : new Date();
  if (isNaN(d.getTime())) return formatDate(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/**
 * Reusable employee lookup function by Employee No / ID.
 * Searches the existing Employees sheet in EMPLOYEE_SPREADSHEET_ID.
 * Matches Employee No exactly (case-insensitive, trimmed).
 * Returns employee details object or null if not found.
 */
function getEmployeeById(employeeNo) {
  if (!employeeNo || String(employeeNo).trim() === '') return null;
  const cleanId = String(employeeNo).trim().toLowerCase();

  const empSheet = getSheet(SHEET_NAMES.employees);
  if (!empSheet) return null;

  const rows = sheetToJson(empSheet);
  const emp = rows.find(r => {
    const idVal = String(
      r.ID || r.EmployeeID || r.EmployeeNo || r.EmpID || r.StaffID || 
      r['Employee ID'] || r['Employee No'] || r['Staff ID'] || ''
    ).trim().toLowerCase();
    return idVal === cleanId;
  });

  if (!emp) return null;

  return {
    ID: String(emp.ID || emp.EmployeeID || emp.EmployeeNo || String(employeeNo).trim()).trim(),
    Name: String(emp.Name || emp.EmployeeName || emp['Employee Name'] || emp['Staff Name'] || '').trim(),
    Department: String(emp.Department || emp.CostCentre || emp['Cost Centre'] || 'N/A').trim(),
    Position: String(emp.Position || emp.JobTitle || emp.PositionTitle || emp['Position Title'] || emp['Job Title'] || 'Staff').trim(),
    Email: String(emp.Email || emp['Email Address'] || '').trim()
  };
}

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

// ─── Response Helpers ───────────────────────────────────────────────────────────
function ok(data) {
  return JSON.stringify({ success: true, data: data });
}

function err(message) {
  return JSON.stringify({ success: false, message: message });
}

function setupSheets() {
  getSheet(SHEET_NAMES.trainings);
  getSheet(SHEET_NAMES.employees);

  try {
    const masterSS = getSpreadsheet();
    if (masterSS) {
      const allowedNames = [SHEET_NAMES.trainings, SHEET_NAMES.employees];
      const sheets = masterSS.getSheets();
      if (sheets.length > 1) {
        sheets.forEach(s => {
          const sName = s.getName();
          if (!allowedNames.includes(sName)) {
            try { masterSS.deleteSheet(s); } catch(e) {}
          }
        });
      }
    }
  } catch(e) {}

  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet && tSheet.getLastRow() > 1) {
      ensureTrainingSheetColumns(tSheet);
      const rows = sheetToJson(tSheet);
      rows.forEach(t => {
        if (t.ID && t.Code && t.Name) {
          const workspace = createTrainingWorkspace(t.Code, t.Name);
          const reqForm = createTrainingRequisitionForm(t.Code, t, workspace.folderId);

          if (t._row) {
            const h = ensureTrainingSheetColumns(tSheet);
            const setColVal = (colName, val) => {
              const idx = h.indexOf(colName) + 1;
              if (idx > 0 && val) tSheet.getRange(t._row, idx).setValue(val);
            };
            setColVal('FolderID', workspace.folderId);
            setColVal('ParticipantsSheetID', workspace.partSheetId);
            setColVal('SessionsSheetID', workspace.sessionSheetId);
            setColVal('AttendanceSheetID', workspace.attendanceSheetId);
            setColVal('EvaluationSheetID', workspace.evaluationSheetId);
            setColVal('PostSheetID', workspace.postSheetId);
            if (reqForm.fileId) setColVal('RequisitionFormFileID', reqForm.fileId);
          }

          try { syncParticipantsToTrainingDriveSheet(t.ID); } catch(e) {}
          try { syncTrainingRequisitionParticipants(t.ID); } catch(e) {}
        }
      });
    }
  } catch (e) {
    Logger.log('setupSheets workspace sync error: ' + e.message);
  }

  Logger.log('Master database cleaned and per-training single-sheet Drive workspaces initialised successfully.');
}
