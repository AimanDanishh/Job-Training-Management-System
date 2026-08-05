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
    Trainings:        ['ID', 'Code', 'Name', 'Category', 'Trainer', 'Venue', 'StartDate',
                       'EndDate', 'Duration', 'TotalHours', 'Department', 'Objectives',
                       'Status', 'Stage', 'Participants',
                       'FolderID', 'ParticipantsSheetID', 'SessionsSheetID', 'AttendanceSheetID',
                       'EvaluationSheetID', 'PostSheetID', 'RequisitionFormFileID',
                       'CreatedDate', 'UpdatedDate', 'CourseFee',
                       'ApprovalStatus', 'RequestedBy', 'RequestedDate', 'ApprovedBy', 'ApprovedCostCentre', 'ApprovedAt', 'ApprovalRemarks', 'RescheduledDate'],
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
  if (sheet.getLastRow() > 1) return;

  const key = Object.keys(SHEET_NAMES).find(k => SHEET_NAMES[k] === name || k === name || name.toLowerCase().includes(k.toLowerCase())) || name;
  const timeNow = now();

  if (key === 'employees' || name === 'Employees' || name === 'Employees') {
    const sampleEmployees = [
      ['EMP-1001', 'Ahmad Razak', 'Engineering / Cost Centre 101', 'Senior Engineer', 'ahmad.razak@company.com', '+60 12-3456789', 'Active'],
      ['EMP-1002', 'Siti Nurhaliza', 'HR / Cost Centre 102', 'HR Executive', 'siti.nurhaliza@company.com', '+60 13-9876543', 'Active'],
      ['EMP-1003', 'Tan Wei Liang', 'Finance / Cost Centre 103', 'Financial Analyst', 'wei.liang@company.com', '+60 16-4567890', 'Active'],
      ['EMP-1004', 'Muthu Kumar', 'Operations / Cost Centre 104', 'Operations Supervisor', 'muthu.k@company.com', '+60 17-2345678', 'Active'],
      ['EMP-1005', 'Lee Jia Hui', 'IT / Cost Centre 105', 'System Admin', 'jiahui.lee@company.com', '+60 19-8765432', 'Active'],
      ['EMP-1006', 'Faridah Hashim', 'Sales / Cost Centre 106', 'Sales Manager', 'faridah.h@company.com', '+60 11-1234567', 'Active'],
      ['EMP-1007', 'Chong Jin Hoe', 'Engineering / Cost Centre 101', 'Software Engineer', 'jinhoe.c@company.com', '+60 14-5678901', 'Active'],
      ['EMP-1008', 'Nadia Azman', 'HR / Cost Centre 102', 'Talent Acquisition Specialist', 'nadia.a@company.com', '+60 18-9012345', 'Active']
    ];
    sampleEmployees.forEach(r => sheet.appendRow(r));
  } else if (key === 'trainings' || name === 'Trainings') {
    const sampleTrainings = [
      [
        'TRN-1001', 'LM-2026-0001', 'Leadership Excellence & Strategic Management',
        'Leadership & Management', 'Dr. Aris Thorne', 'Grand Ballroom / Online',
        '2026-08-10', '2026-08-12', 3, 24, 'HR / Cost Centre 102',
        'Enhance strategic leadership capabilities and team management skills.',
        'Upcoming', 'Created', 4, '', '', '', '', '', '', '', '', timeNow, timeNow, '1500.00', ''
      ],
      [
        'TRN-1002', 'CR-2026-0002', 'ISO 27001 Cybersecurity & Data Compliance',
        'Compliance & Regulatory', 'Sarah Jenkins', 'Training Room A',
        '2026-08-01', '2026-08-02', 2, 16, 'IT / Cost Centre 105',
        'Comprehensive security protocols and compliance training.',
        'In Progress', 'Attendance In Progress', 2, '', '', '', '', '', '', '', '', timeNow, timeNow, '1200.00', ''
      ]
    ];
    sampleTrainings.forEach(r => sheet.appendRow(r));
  } else if (key === 'trainingParticipants' || name === 'TrainingParticipants') {
    const sampleParticipants = [
      ['TP-1001', 'TRN-1001', 'EMP-1001', 'Ahmad Razak', 'Engineering / Cost Centre 101', 'Senior Engineer', timeNow],
      ['TP-1002', 'TRN-1001', 'EMP-1002', 'Siti Nurhaliza', 'HR / Cost Centre 102', 'HR Executive', timeNow],
      ['TP-1003', 'TRN-1001', 'EMP-1003', 'Tan Wei Liang', 'Finance / Cost Centre 103', 'Financial Analyst', timeNow],
      ['TP-1004', 'TRN-1001', 'EMP-1004', 'Muthu Kumar', 'Operations / Cost Centre 104', 'Operations Supervisor', timeNow],
      ['TP-1005', 'TRN-1002', 'EMP-1005', 'Lee Jia Hui', 'IT / Cost Centre 105', 'System Admin', timeNow],
      ['TP-1006', 'TRN-1002', 'EMP-1007', 'Chong Jin Hoe', 'Engineering / Cost Centre 101', 'Software Engineer', timeNow]
    ];
    sampleParticipants.forEach(r => sheet.appendRow(r));
  } else if (key === 'trainingSessions' || name === 'TrainingSessions') {
    const sampleSessions = [
      ['SES0001', 'TRN-1001', 'Day 1 - Morning', '2026-08-10', '09:00', '12:00', '', '', 'Active', timeNow],
      ['SES0002', 'TRN-1001', 'Day 1 - Afternoon', '2026-08-10', '13:30', '17:00', '', '', 'Active', timeNow],
      ['SES0003', 'TRN-1002', 'Day 1', '2026-08-01', '09:00', '17:00', '', '', 'Active', timeNow],
      ['SES0004', 'TRN-1002', 'Day 2', '2026-08-02', '09:00', '17:00', '', '', 'Active', timeNow]
    ];
    sampleSessions.forEach(r => sheet.appendRow(r));
  }
}

/** Adds fields introduced after the original training sheet was deployed. */
function ensureTrainingSheetColumns(sheet) {
  const requiredHeaders = [
    'FolderID', 'ParticipantsSheetID', 'SessionsSheetID', 'AttendanceSheetID',
    'EvaluationSheetID', 'PostSheetID', 'RequisitionFormFileID', 'CourseFee'
  ];
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

/**
 * Convert a sheet's data rows into an array of objects keyed by header row.
 */
function sheetToJson(sheet) {
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
 * Run setupSheets() once from the Apps Script editor to initialise all sheets
 * and auto-create single-sheet formula-equipped Drive workspaces for all training programmes.
 */
function setupSheets() {
  Object.values(SHEET_NAMES).forEach(name => getSheet(name));

  // Ensure Master Database Spreadsheet (SPREADSHEET_ID) contains only the Trainings tab (and Employees tab if stored in master SS)
  try {
    const masterSS = getSpreadsheet();
    if (masterSS) {
      const allowedNames = [SHEET_NAMES.trainings, SHEET_NAMES.employees];
      const sheets = masterSS.getSheets();
      if (sheets.length > 1) {
        sheets.forEach(s => {
          const sName = s.getName();
          if (!allowedNames.includes(sName) && !sName.toLowerCase().includes('training') && !sName.toLowerCase().includes('emp')) {
            try { masterSS.deleteSheet(s); } catch(e) {}
          }
        });
      }
    }
  } catch(e) {}

  // Sync per-training Google Drive workspaces and single sheet tabs
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

          // Sync sessions to single sheet
          try {
            const sessSheet = getSheet(SHEET_NAMES.trainingSessions);
            if (sessSheet) {
              const sessions = sheetToJson(sessSheet).filter(s => String(s.TrainingID) === String(t.ID));
              sessions.forEach(s => {
                const sRow = [s.SessionID, s.TrainingID, s.SessionName, s.SessionDate, s.StartTime, s.EndTime, s.AttendanceURL, s.QRCodeURL, s.QRStatus, s.CreatedDate];
                syncSessionToTrainingDriveSheet(t.ID, sRow);
              });
            }
          } catch(e) {}

          // Sync attendance to single sheet
          try {
            const attSheet = getSheet(SHEET_NAMES.attendance);
            if (attSheet) {
              const atts = sheetToJson(attSheet).filter(a => String(a.TrainingID) === String(t.ID));
              atts.forEach(a => {
                const aRow = [a.AttendanceID, a.SessionID, a.TrainingID, a.EmployeeNo, a.EmployeeName, a.Department, a.ScanTime, a.Status, a.TrainingCode, a.Day, a.Date, a.Hours, a.Remarks, a.EditedBy, a.EditedAt];
                syncAttendanceToTrainingDriveSheet(t.ID, aRow);
              });
            }
          } catch(e) {}

          // Sync evaluations to single sheet
          try {
            const evalSheet = getSheet(SHEET_NAMES.trainingEval);
            if (evalSheet) {
              const evals = sheetToJson(evalSheet).filter(ev => String(ev.TrainingID) === String(t.ID));
              evals.forEach(ev => {
                const evRow = [ev.ID, ev.TrainingID, ev.EmployeeID, ev.EmployeeName, ev.Q1, ev.Q2, ev.Q3, ev.Q4, ev.Q5, ev.Q6, ev.Q7, ev.SectionB1, ev.SectionB2, ev.SectionB3, ev.AvgScore, ev.SubmittedAt];
                syncEvaluationToTrainingDriveSheet(t.ID, evRow);
              });
            }
          } catch(e) {}

          // Sync post evaluations to single sheet
          try {
            const postSheet = getSheet(SHEET_NAMES.postEval);
            if (postSheet) {
              const posts = sheetToJson(postSheet).filter(p => String(p.TrainingID) === String(t.ID));
              posts.forEach(p => {
                const pRow = [p.ID, p.TrainingID, p.EmployeeID, p.EvaluatorName, p.EvaluatorID, p.CompetencyBefore, p.CompetencyAfter, p.Improvement, p.CanApply, p.FurtherTraining, p.Comments, p.SubmittedAt];
                syncPostEvalToTrainingDriveSheet(t.ID, pRow);
              });
            }
          } catch(e) {}
        }
      });
    }
  } catch (e) {
    Logger.log('setupSheets workspace sync error: ' + e.message);
  }

  Logger.log('All database sheets and per-training single-sheet Drive workspaces initialised successfully.');
}
