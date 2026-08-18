/**
 * Helper.gs - Shared utilities, sheet access, script properties, and initial setup
 */

// --- Script Properties Configuration ------------------------------------------
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
 * Safely populates default Script Properties without overwriting existing configured properties.
 * If overwriteExisting is true, non-empty default values will update existing keys, but empty defaults are never saved.
 */
function initDefaultScriptProperties(overwriteExisting) {
  const isOverwrite = overwriteExisting === true;
  let existingProps = {};
  try {
    existingProps = PropertiesService.getScriptProperties().getProperties() || {};
  } catch (e) {
    Logger.log('Warning: Could not fetch existing script properties: ' + e.message);
  }

  const defaults = {
    'SPREADSHEET_ID':          '',
    'EMPLOYEE_SPREADSHEET_ID': '',
    'ALLOWED_DOMAIN':          '',
    'ADMIN_EMAILS':            '',
    'APP_TITLE':               'TrainHub - Training Management System',
    'SHEET_EMPLOYEES':         'For IT',
    'SHEET_HR_EMAIL':          'HR email',
    'SHEET_TRAININGS':         'Trainings',
    'ROOT_FOLDER_ID':          '',
    'TRAINING_FOLDER':         '',
    'TEMPLATE_FOLDER_ID':      '',
    'ATTENDANCE_TEMPLATE_ID':  '',
    'EVALUATION_TEMPLATE_ID':  '',
    'CERTIFICATE_TEMPLATE_ID': '',
    'REPORT_TEMPLATE_ID':      '',
    'TRAINING_REQUISITION_TEMPLATE_ID': '',
    'COMPANY_LOGO_URL':        '',
    'PUBLIC_PORTAL_URL':       '',
    'HOD_PORTAL_URL':          ''
  };

  const toUpdate = {};
  for (const key in defaults) {
    const existingVal = existingProps[key];
    const defaultVal = defaults[key];

    if (isOverwrite) {
      if (defaultVal !== '') {
        toUpdate[key] = defaultVal;
      }
    } else {
      if ((existingVal === undefined || existingVal === null || String(existingVal).trim() === '') && defaultVal !== '') {
        toUpdate[key] = defaultVal;
      }
    }
  }

  if (Object.keys(toUpdate).length > 0) {
    PropertiesService.getScriptProperties().setProperties(toUpdate, false);
    Logger.log('Safely updated missing Script Properties: ' + Object.keys(toUpdate).join(', '));
  } else {
    Logger.log('No Script Properties needed initialization. Existing settings preserved.');
  }

  PropertiesService.getScriptProperties().deleteProperty('ADMIN_USER');
  PropertiesService.getScriptProperties().deleteProperty('ADMIN_PASS');
}

/**
 * Safely updates database-dependent configuration without modifying unrelated properties.
 */
function updateDatabaseConfiguration(spreadsheetId, employeeSpreadsheetId) {
  if (!spreadsheetId || String(spreadsheetId).trim() === '') {
    throw new Error('Spreadsheet ID is required.');
  }
  const cleanSsId = String(spreadsheetId).trim();
  setConfigProperty('SPREADSHEET_ID', cleanSsId);

  if (employeeSpreadsheetId && String(employeeSpreadsheetId).trim() !== '') {
    setConfigProperty('EMPLOYEE_SPREADSHEET_ID', String(employeeSpreadsheetId).trim());
  }

  // Invalidate in-memory spreadsheet handle caches so newly configured IDs take effect immediately
  _cachedSpreadsheet = null;
  _cachedEmployeeSpreadsheet = null;

  Logger.log('Database configuration safely updated. SPREADSHEET_ID: ' + cleanSsId);
  return true;
}

/**
 * Diagnostic utility to return status of all system properties without revealing secrets.
 */
function getSystemConfigurationDiagnostics() {
  let props = {};
  try {
    props = PropertiesService.getScriptProperties().getProperties() || {};
  } catch (e) {}

  const keys = [
    'SPREADSHEET_ID',
    'EMPLOYEE_SPREADSHEET_ID',
    'ROOT_FOLDER_ID',
    'TRAINING_FOLDER',
    'TRAINING_REQUISITION_TEMPLATE_ID',
    'ALLOWED_DOMAIN',
    'ADMIN_EMAILS',
    'APP_TITLE',
    'SHEET_EMPLOYEES',
    'SHEET_HR_EMAIL',
    'SHEET_TRAININGS',
    'COMPANY_LOGO_URL',
    'PUBLIC_PORTAL_URL',
    'HOD_PORTAL_URL'
  ];

  const result = {};
  keys.forEach(k => {
    const val = props[k];
    result[k] = (val !== undefined && val !== null && String(val).trim() !== '') ? 'Configured' : 'Not configured';
  });
  return result;
}

/**
 * Database setup and initialization function.
 * Connects to a new/existing database, validates the Spreadsheet ID, creates missing standard sheets,
 * updates database settings, and preserves all unrelated configuration.
 */
function setupDatabase(spreadsheetId, employeeSpreadsheetId) {
  if (!spreadsheetId || String(spreadsheetId).trim() === '') {
    return {
      success: false,
      message: 'Spreadsheet ID is required for Database Setup.'
    };
  }

  const cleanSsId = String(spreadsheetId).trim();
  const cleanEmpSsId = employeeSpreadsheetId ? String(employeeSpreadsheetId).trim() : cleanSsId;

  try {
    // 1. Connect and validate main Spreadsheet
    const ss = SpreadsheetApp.openById(cleanSsId);
    if (!ss) {
      return { success: false, message: 'Failed to access Google Spreadsheet with ID: ' + cleanSsId };
    }

    // 2. Connect optional separate Employee Spreadsheet
    let empSS = ss;
    if (cleanEmpSsId && cleanEmpSsId !== cleanSsId) {
      try {
        empSS = SpreadsheetApp.openById(cleanEmpSsId);
      } catch (e) {
        return { success: false, message: 'Failed to access separate Employee Spreadsheet with ID: ' + cleanEmpSsId };
      }
    }

    // 3. Main Database contains ONLY 'Trainings' tab
    const mainSheets = [
      {
        name: getConfigProperty('SHEET_TRAININGS', 'Trainings'),
        targetSS: ss,
        headers: ['ID', 'Code', 'Name', 'Type', 'Organizer', 'StartDate', 'EndDate', 'Location', 'Status', 'ApprovalStatus', 'FolderID', 'ParticipantsSheetID', 'SessionsSheetID', 'AttendanceSheetID', 'EvaluationSheetID', 'PostSheetID', 'RequisitionFormFileID']
      }
    ];

    const sheetStatuses = [];
    mainSheets.forEach(def => {
      let sheet = def.targetSS.getSheetByName(def.name);
      if (!sheet) {
        sheet = def.targetSS.insertSheet(def.name);
        if (def.headers && def.headers.length > 0) {
          sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
        }
        sheetStatuses.push({ name: def.name, status: 'Created', connected: true });
      } else {
        sheetStatuses.push({ name: def.name, status: 'Connected', connected: true });
      }
    });

    // 4. Employee Spreadsheet READ-ONLY Connection Check (never modify or insert sheets)
    if (empSS) {
      ['For IT', 'HOD email', 'Csuite email', 'HOHR email', 'HR email', 'Cost Centre'].forEach(tabName => {
        const foundTab = empSS.getSheetByName(tabName);
        sheetStatuses.push({
          name: tabName,
          status: foundTab ? 'Connected (Read-Only)' : 'Not found (Optional)',
          connected: !!foundTab
        });
      });
    }

    // 5. Note per-training Drive Sheets
    sheetStatuses.push({
      name: 'Individual Training Data (${code} Training Data)',
      status: 'Managed in Drive per training (Attendance & Evals)',
      connected: true
    });

    // 4. Update ONLY database properties without wiping unrelated configuration
    updateDatabaseConfiguration(cleanSsId, cleanEmpSsId);

    // 5. Initialize defaults safely (filling missing keys only)
    initDefaultScriptProperties(false);

    // 6. Validate configuration
    const validation = (typeof validateSystemConfiguration === 'function') ? validateSystemConfiguration() : null;

    return {
      success: true,
      message: 'Database setup completed successfully.',
      summary: {
        databaseConnected: true,
        databaseId: cleanSsId,
        databaseName: ss.getName(),
        sheets: sheetStatuses,
        configSaved: true,
        unrelatedConfigPreserved: true,
        validation: validation
      }
    };
  } catch (e) {
    Logger.log('setupDatabase error: ' + e.message);
    return {
      success: false,
      message: 'Database Setup failed: ' + e.message
    };
  }
}

function getSpreadsheetId() {
  return getConfigProperty('SPREADSHEET_ID', '');
}

function getEmployeeSpreadsheetId() {
  return getConfigProperty('EMPLOYEE_SPREADSHEET_ID', getSpreadsheetId());
}

const DEFAULT_APOLLO_LOGO_URL = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTAwIj48cmVjdCB3aWR0aD0iMTAwIiBoZWlnaHQ9IjEwMCIgcng9IjIyIiBmaWxsPSIjRTkxQzJEIi8+PGNpcmNsZSBjeD0iNTAiIGN5PSI1MCIgcj0iMjgiIGZpbGw9IiNGRkY4Q0YiIHN0cm9rZT0iI0ZFQzMwRCIgc3Ryb2tlLXdpZHRoPSIyIi8+PHRleHQgeD0iNTAiIHk9IjYxIiBmb250LWZhbWlseT0iQXJpYWwsIHNhbnMtc2VyaWYiIGZvbnQtd2VpZ2h0PSI5MDAiIGZvbnQtc2l6ZT0iMzIiIGZpbGw9IiNFOTFDMkQiIHRleHQtYW5jaG9yPSJtaWRkbGUiPkE8L3RleHQ+PC9zdmc+';

function getCompanyLogoUrl() {
  const url = getConfigProperty('COMPANY_LOGO_URL', '');
  if (url && String(url).trim() !== '') return String(url).trim();
  return DEFAULT_APOLLO_LOGO_URL;
}

function setCompanyLogoUrl(url) {
  setConfigProperty('COMPANY_LOGO_URL', url);
  Logger.log('COMPANY_LOGO_URL updated in Project Settings: ' + url);
  return 'COMPANY_LOGO_URL set to: ' + url;
}

function isSameEmployeeId(id1, id2) {
  if (id1 === null || id1 === undefined || id2 === null || id2 === undefined) return false;
  const s1 = String(id1).trim().toLowerCase();
  const s2 = String(id2).trim().toLowerCase();
  if (s1 === s2) return true;
  const z1 = s1.replace(/^0+/, '');
  const z2 = s2.replace(/^0+/, '');
  if (z1 !== '' && z1 === z2) return true;
  if (s1.replace(/0/g, '') === '' && s2.replace(/0/g, '') === '') return true;
  return false;
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
  get employees()            { return getConfigProperty('SHEET_EMPLOYEES', 'For IT'); },
  get hrEmail()              { return getConfigProperty('SHEET_HR_EMAIL', 'HR email'); },
  get trainings()            { return getConfigProperty('SHEET_TRAININGS', 'Trainings'); },
  get trainingSessions()     { return getConfigProperty('SHEET_TRAINING_SESSIONS', 'Sessions'); },
  get attendance()           { return getConfigProperty('SHEET_ATTENDANCE', 'Attendance'); },
  get trainingEval()         { return getConfigProperty('SHEET_TRAINING_EVAL', 'Evaluation'); },
  get postEval()             { return getConfigProperty('SHEET_POST_EVAL', 'Post Evaluation'); },
  get trainingParticipants() { return getConfigProperty('SHEET_TRAINING_PARTICIPANTS', 'Participants'); }
};

// --- Spreadsheet Access --------------------------------------------------------
let _cachedSpreadsheet = null;
function getSpreadsheet() {
  if (_cachedSpreadsheet) return _cachedSpreadsheet;

  const spreadsheetId = getSpreadsheetId();
  if (spreadsheetId && String(spreadsheetId).trim() !== '') {
    try {
      _cachedSpreadsheet = SpreadsheetApp.openById(String(spreadsheetId).trim());
      if (_cachedSpreadsheet) return _cachedSpreadsheet;
    } catch (e) {
      Logger.log('Warning: Failed to open configured SPREADSHEET_ID (' + spreadsheetId + '): ' + e.message);
    }
  }

  // Fallback 1: Container active spreadsheet
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      _cachedSpreadsheet = active;
      try { setConfigProperty('SPREADSHEET_ID', active.getId()); } catch(pErr) {}
      return _cachedSpreadsheet;
    }
  } catch(aErr) {}

  // Fallback 2: Search Google Drive for existing 'TrainHub Database' or 'TrainHub Master Database'
  try {
    let files = DriveApp.getFilesByName('TrainHub Database');
    if (!files.hasNext()) files = DriveApp.getFilesByName('TrainHub Master Database');
    if (files.hasNext()) {
      const file = files.next();
      _cachedSpreadsheet = SpreadsheetApp.openById(file.getId());
      if (_cachedSpreadsheet) {
        try { setConfigProperty('SPREADSHEET_ID', file.getId()); } catch(pErr) {}
        return _cachedSpreadsheet;
      }
    }
  } catch(dErr) {}

  return null;
}

let _cachedEmployeeSpreadsheet = null;
function getEmployeeSpreadsheet() {
  if (_cachedEmployeeSpreadsheet) return _cachedEmployeeSpreadsheet;

  const empSpreadsheetId = getEmployeeSpreadsheetId();
  if (empSpreadsheetId && String(empSpreadsheetId).trim() !== '') {
    try {
      _cachedEmployeeSpreadsheet = SpreadsheetApp.openById(String(empSpreadsheetId).trim());
      if (_cachedEmployeeSpreadsheet) return _cachedEmployeeSpreadsheet;
    } catch(e) {
      Logger.log('Failed to open separate EMPLOYEE_SPREADSHEET_ID (' + empSpreadsheetId + '): ' + e.message);
    }
  }

  // Fallback: Search Google Drive for 'Namelist' or 'TrainHub Employee Master Database'
  try {
    let files = DriveApp.getFilesByName('Namelist');
    if (!files.hasNext()) files = DriveApp.getFilesByName('TrainHub Employee Master Database');
    if (files.hasNext()) {
      const file = files.next();
      _cachedEmployeeSpreadsheet = SpreadsheetApp.openById(file.getId());
      if (_cachedEmployeeSpreadsheet) return _cachedEmployeeSpreadsheet;
    }
  } catch(dErr) {}

  return getSpreadsheet();
}

function getSheet(name) {
  const cleanNameLower = String(name || '').toLowerCase().trim();
  const isEmployeeSpreadsheetSheet = (
    cleanNameLower === 'employees' ||
    cleanNameLower === 'for it' ||
    cleanNameLower === 'hr email' ||
    cleanNameLower === 'hod email' ||
    cleanNameLower === 'csuite email' ||
    cleanNameLower === 'c-suite email' ||
    cleanNameLower === 'hohr email' ||
    cleanNameLower === 'cost centre' ||
    cleanNameLower === 'costcentre'
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
             (targetClean === 'forit' && (sClean === 'employees' || sClean.includes('employee'))) ||
             (targetClean === 'employees' && sClean === 'forit') ||
             (targetClean === 'participants' && sClean === 'trainingparticipants') ||
             (targetClean === 'trainingparticipants' && sClean === 'participants') ||
             (targetClean === 'sessions' && sClean === 'trainingsessions') ||
             (targetClean === 'trainingsessions' && sClean === 'sessions') ||
             (targetClean === 'evaluation' && sClean === 'trainingeval') ||
             (targetClean === 'trainingeval' && sClean === 'evaluation') ||
             (targetClean === 'postevaluation' && sClean === 'posteval') ||
             (targetClean === 'posteval' && sClean === 'postevaluation');
    });

    if (!sheet) {
      if (isEmployeeSpreadsheetSheet) {
        // Employee spreadsheet is READ-ONLY - do not modify or insert sheets
        if (allSheets.length > 0) {
          sheet = allSheets.find(s => s.getName().toLowerCase().includes('for it') || s.getName().toLowerCase().includes('employee')) || allSheets[0];
        }
      } else {
        sheet = ss.insertSheet(name);
        initSheetHeaders(sheet, name);
      }
    }
  }

  // Ensure required columns exist for Trainings and Sessions sheets
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

// --- Sheet Headers -------------------------------------------------------------
function initSheetHeaders(sheet, name) {
  const headers = {
    'For IT':             ['ID', 'Name', 'Department', 'Position Title', 'Email', 'Phone', 'Status'],
    Employees:            ['ID', 'Name', 'Department', 'Position Title', 'Email', 'Phone', 'Status'],
    'HR email':           ['Employee No', 'HR', 'Cost Centre', 'Position Title', 'Email'],
    Trainings:            ['ID', 'Code', 'Name', 'Category', 'Trainer', 'Venue', 'StartDate',
                           'EndDate', 'Duration', 'TotalHours', 'Department', 'Objectives',
                           'Status', 'Stage', 'Participants',
                           'FolderID', 'ParticipantsSheetID', 'SessionsSheetID', 'AttendanceSheetID',
                           'EvaluationSheetID', 'PostSheetID', 'RequisitionFormFileID',
                           'CreatedDate', 'UpdatedDate', 'CourseFee',
                           'ApprovalStatus', 'RequestedBy', 'RequestedByName', 'RequestedByEmail', 'RequestedDate', 'ApprovedBy', 'ApprovedCostCentre', 'ApprovedAt', 'ApprovalRemarks', 'RescheduledDate'],
    Sessions:             ['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate'],
    TrainingSessions:     ['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate'],
    Attendance:           ['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName',
                           'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours',
                           'Remarks', 'EditedBy', 'EditedAt'],
    Evaluation:           ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName',
                           'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7',
                           'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt'],
    TrainingEval:         ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName',
                           'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7',
                           'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt'],
    'Post Evaluation':    ['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID',
                           'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply',
                           'FurtherTraining', 'Comments', 'SubmittedAt'],
    PostEval:             ['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID',
                           'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply',
                           'FurtherTraining', 'Comments', 'SubmittedAt'],
    Participants:         ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt', 'SupervisorID', 'SupervisorEmail', 'SupervisorName'],
    TrainingParticipants: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt', 'SupervisorID', 'SupervisorEmail', 'SupervisorName']
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
    'ApprovalStatus', 'RequestedBy', 'RequestedByName', 'RequestedByEmail', 'RequestedDate', 'ApprovedBy', 'ApprovedCostCentre', 'ApprovedAt', 'ApprovalRemarks', 'RescheduledDate', 'BrochureURL',
    'TrainingProvider', 'ExpiryDate', 'CertExpiryDate',
    'HOD', 'Csuite', 'HOHR', 'HODStatus', 'CsuiteStatus', 'HOHRStatus'
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
  const t = trainings.find(r => 
    String(r.ID || '').trim() === cleanId ||
    String(r.Code || '').trim() === cleanId ||
    String(r.TrainingID || '').trim() === cleanId ||
    (r.Name && String(r.Name).trim().toLowerCase() === cleanId.toLowerCase())
  );
  if (!t) return null;

  // 1. Direct Resolution: Open by stored ParticipantsSheetID first if available
  const storedSheetId = String(t.ParticipantsSheetID || t.singleSheetId || t.TrainingDataSheetID || '').trim();
  if (storedSheetId) {
    try {
      const ss = SpreadsheetApp.openById(storedSheetId);
      if (ss) return ss;
    } catch(e) {
      Logger.log('Could not open spreadsheet directly via stored ParticipantsSheetID (' + storedSheetId + '): ' + e.message);
    }
  }

  // 2. Drive Workspace Resolution: Search via FolderID or Root Folder
  try {
    let targetFolder = null;

    if (t.FolderID) {
      try {
        targetFolder = DriveApp.getFolderById(String(t.FolderID).trim());
      } catch(e) {}
    }

    if (!targetFolder && typeof getOrCreateTrainingRootFolder === 'function') {
      try {
        const trainingRoot = getOrCreateTrainingRootFolder();
        if (trainingRoot) {
          const code = t.Code || t.ID || cleanId;
          const folderName = `${code} ${t.Name || ''}`.trim();
          let folderIter = trainingRoot.getFoldersByName(folderName);
          if (folderIter.hasNext()) targetFolder = folderIter.next();
        }
      } catch(fErr) {}
    }

    if (targetFolder) {
      const code = t.Code || t.ID || cleanId;
      let fileIter = targetFolder.getFilesByName('Training Data');
      if (!fileIter.hasNext()) fileIter = targetFolder.getFilesByName(`${code} Training Data`);
      if (fileIter.hasNext()) {
        const file = fileIter.next();
        const ss = SpreadsheetApp.openById(file.getId());
        if (ss) {
          // Auto-persist resolved ParticipantsSheetID to row if missing
          try {
            if (t._row) {
              const headers = ensureTrainingSheetColumns(tSheet);
              const colIdx = headers.indexOf('ParticipantsSheetID') + 1;
              if (colIdx > 0) tSheet.getRange(t._row, colIdx).setValue(file.getId());
            }
          } catch(persistErr) {}
          return ss;
        }
      }
    }
  } catch (e) {
    Logger.log('Error opening per-training sheet for ' + cleanId + ': ' + e.message);
  }

  return null;
}

/**
 * Enriches a training object with direct URLs to its Google Sheets & Drive folders:
 * - trainingDataSheetUrl: Direct link to per-training 'Training Data' spreadsheet
 * - requisitionFormUrl: Direct link to AP-HRD-F01-01 requisition sheet
 * - folderUrl: Direct link to Google Drive workspace folder
 *
 * @param {Object} t - Training record object
 * @returns {Object} Enriched training object
 */
function enrichTrainingWithUrls(t) {
  if (!t || typeof t !== 'object') return t;

  let partSheetId = String(t.ParticipantsSheetID || t.singleSheetId || '').trim();
  let reqFileId   = String(t.RequisitionFormFileID || t.RequisitionFormId || '').trim();
  let folderId    = String(t.FolderID || '').trim();

  // If partSheetId is missing, attempt to resolve via getTrainingDataSpreadsheet
  if (!partSheetId && t.ID) {
    try {
      const ss = getTrainingDataSpreadsheet(t.ID);
      if (ss) {
        partSheetId = ss.getId();
        t.ParticipantsSheetID = partSheetId;
      }
    } catch(e) {}
  }

  if (partSheetId) {
    t.trainingDataSheetUrl = 'https://docs.google.com/spreadsheets/d/' + partSheetId + '/edit';
  } else {
    t.trainingDataSheetUrl = '';
  }

  if (reqFileId) {
    t.requisitionFormUrl = 'https://docs.google.com/spreadsheets/d/' + reqFileId + '/edit';
  } else if (t.RequisitionUrl || t.BrochureURL) {
    t.requisitionFormUrl = t.RequisitionUrl || t.BrochureURL;
  } else {
    t.requisitionFormUrl = '';
  }

  if (folderId) {
    t.folderUrl = 'https://drive.google.com/drive/folders/' + folderId;
  } else {
    t.folderUrl = '';
  }

  return t;
}

/**
 * Helper to look up a training session across central database sheet and all per-training sheets
 * 
 * @param {string} sessionId - Session ID (e.g. SES0001)
 * @returns {Object|null} { session: Object, training: Object, spreadsheet: Spreadsheet, sessionSheet: Sheet }
 */
function findTrainingBySessionId(sessionId) {
  if (!sessionId) return null;
  const cleanSessionId = String(sessionId).trim().toLowerCase();

  // Tier 1: Check central Main Database spreadsheet first
  try {
    const mainSs = getSpreadsheet();
    if (mainSs) {
      const centralSessSheet = mainSs.getSheetByName('TrainingSessions') || 
                               mainSs.getSheetByName('Sessions') || 
                               mainSs.getSheetByName('Training Sessions') || 
                               mainSs.getSheetByName('Session');
      if (centralSessSheet) {
        const sessions = sheetToJson(centralSessSheet);
        const session = sessions.find(s => {
          const sId = String(s.SessionID || s.ID || s.SessionCode || '').trim().toLowerCase();
          return sId === cleanSessionId;
        });
        if (session) {
          const tSheet = getSheet(SHEET_NAMES.trainings);
          const trainings = tSheet ? sheetToJson(tSheet) : [];
          const t = trainings.find(r => {
            const id = String(r.ID || '').trim().toLowerCase();
            const code = String(r.Code || '').trim().toLowerCase();
            const tId = String(r.TrainingID || '').trim().toLowerCase();
            return id === String(session.TrainingID || '').trim().toLowerCase() ||
                   code === String(session.TrainingID || '').trim().toLowerCase() ||
                   tId === String(session.TrainingID || '').trim().toLowerCase();
          }) || { ID: session.TrainingID };

          const perTrainingSs = getTrainingDataSpreadsheet(t) || mainSs;
          return { session: session, training: t, spreadsheet: perTrainingSs, sessionSheet: centralSessSheet };
        }
      }
    }
  } catch(eCentral) {}

  // Tier 2: Check per-training spreadsheets
  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) return null;

  const trainings = sheetToJson(tSheet);
  for (const t of trainings) {
    if (!t.ID && !t.Code) continue;
    const ss = getTrainingDataSpreadsheet(t);
    if (!ss) continue;
    const sessSheet = ss.getSheetByName('TrainingSessions') || 
                      ss.getSheetByName('Sessions') || 
                      ss.getSheetByName('Training Sessions') || 
                      ss.getSheetByName('Session');
    if (!sessSheet) continue;
    const sessions = sheetToJson(sessSheet);
    const session = sessions.find(s => {
      const sId = String(s.SessionID || s.ID || s.SessionCode || '').trim().toLowerCase();
      return sId === cleanSessionId;
    });
    if (session) {
      return { session: session, training: t, spreadsheet: ss, sessionSheet: sessSheet };
    }
  }
  return null;
}

// --- ID Generation -------------------------------------------------------------
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
      if (normKey) {
        obj[normKey] = val;
      }
    });

    // Only fallback obj.ID and obj.Name if column 0 isn't an ID for another entity (like AttendanceID or SessionID)
    const firstHeaderClean = String(headers[0] || '').trim().toLowerCase();
    if (!obj.ID && data[i][0] !== undefined && !firstHeaderClean.includes('attendance') && !firstHeaderClean.includes('session')) {
      obj.ID = String(data[i][0]);
    }
    if (!obj.Name && data[i][1] !== undefined) obj.Name = String(data[i][1]);

    obj._row = i + 1; // 1-indexed sheet row number
    rows.push(obj);
  }

  if (sheetName) {
    _sheetDataCache[sheetName] = rows;
  }
  return rows;
}

// --- Find Row By ID -------------------------------------------------------------
function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// --- Date Helpers ---------------------------------------------------------------
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


function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

// --- Response Helpers -----------------------------------------------------------
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
