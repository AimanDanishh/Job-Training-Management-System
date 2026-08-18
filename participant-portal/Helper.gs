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

const DEFAULT_APOLLO_LOGO_URL = 'data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" rx="22" fill="%232563EB"/><circle cx="50" cy="50" r="28" fill="%23FFFFFF"/><text x="50" y="61" font-family="Arial, sans-serif" font-weight="900" font-size="32" fill="%232563EB" text-anchor="middle">A</text></svg>';

function getCompanyLogoUrl() {
  const url = getConfigProperty('COMPANY_LOGO_URL', '');
  if (url && String(url).trim() !== '') return String(url).trim();
  return DEFAULT_APOLLO_LOGO_URL;
}

function convertDriveLinkToDirectImageUrl(input) {
  if (!input) return '';
  const str = String(input).trim();
  if (!str.includes('drive.google.com') && str.startsWith('http')) {
    return str;
  }
  let fileId = '';
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
    return `https://lh3.googleusercontent.com/d/${fileId}`;
  }
  return str;
}

function formatMinimalistDate(dateVal) {
  if (!dateVal) return '';
  const str = String(dateVal).replace(/GMT.*$/, '').replace(/\(.*\)$/, '').trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
  const dmyMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmyMatch) return `${dmyMatch[1].padStart(2, '0')}/${dmyMatch[2].padStart(2, '0')}/${dmyMatch[3]}`;
  const ymdMatch = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymdMatch) return `${ymdMatch[3].padStart(2, '0')}/${ymdMatch[2].padStart(2, '0')}/${ymdMatch[1]}`;
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
  }
  return str.split('T')[0] || '';
}

function formatDisplayTime(val, defaultTime = '09:00 AM') {
  if (!val) return defaultTime;
  const str = String(val).replace(/GMT.*$/, '').replace(/\(.*\)$/, '').trim();
  const ampmMatch = str.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
  if (ampmMatch && !str.includes('1899') && !str.includes('Singapore') && !str.includes('Standard') && !str.includes('GMT')) {
    let hh = parseInt(ampmMatch[1], 10);
    const mm = ampmMatch[2];
    const ampm = ampmMatch[3] ? ampmMatch[3].toUpperCase() : (hh >= 12 ? 'PM' : 'AM');
    if (hh > 12) hh -= 12;
    if (hh === 0) hh = 12;
    return `${String(hh).padStart(2, '0')}:${mm} ${ampm}`;
  }
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    let hh = d.getHours();
    const mm = String(d.getMinutes()).padStart(2, '0');
    const ampm = hh >= 12 ? 'PM' : 'AM';
    if (hh > 12) hh -= 12;
    if (hh === 0) hh = 12;
    return `${String(hh).padStart(2, '0')}:${mm} ${ampm}`;
  }
  return defaultTime;
}

function formatSessionTimeRange(startVal, endVal) {
  const start = formatDisplayTime(startVal, '09:00 AM');
  const end = formatDisplayTime(endVal, '05:00 PM');
  return `${start} - ${end}`;
}

const SHEET_NAMES = {
  get employees()            { return getConfigProperty('SHEET_EMPLOYEES', 'Employees'); },
  get trainings()            { return getConfigProperty('SHEET_TRAININGS', 'Trainings'); },
  get trainingSessions()     { return getConfigProperty('SHEET_TRAINING_SESSIONS', 'TrainingSessions'); },
  get trainingParticipants() { return getConfigProperty('SHEET_TRAINING_PARTICIPANTS', 'TrainingParticipants'); }
};

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
             sClean === targetClean + 's' ||
             sClean + 's' === targetClean;
    });

    if (!sheet && isEmployeeSpreadsheetSheet && allSheets.length > 0) {
      sheet = allSheets.find(s => s.getName().toLowerCase().includes('for it') || s.getName().toLowerCase().includes('emp') || s.getName().toLowerCase().includes('staff')) || allSheets[0];
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
function getTrainingDataSpreadsheet(trainingOrId) {
  if (!trainingOrId) return null;
  let t = (typeof trainingOrId === 'object' && trainingOrId !== null) ? trainingOrId : null;
  const cleanId = String(t ? (t.ID || t.TrainingID || t.Code || '') : trainingOrId).trim().toLowerCase();

  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) return null;

  if (!t) {
    const trainings = sheetToJson(tSheet);
    t = trainings.find(r => {
      const id = String(r.ID || '').trim().toLowerCase();
      const code = String(r.Code || '').trim().toLowerCase();
      const tId = String(r.TrainingID || '').trim().toLowerCase();
      return id === cleanId || code === cleanId || tId === cleanId;
    });
  }
  if (!t) return null;

  // 1. Direct Resolution: Open by stored ParticipantsSheetID / SessionsSheetID / singleSheetId
  const storedSheetId = String(t.ParticipantsSheetID || t.SessionsSheetID || t.singleSheetId || t.TrainingDataSheetID || t.TrainingDataID || '').trim();
  if (storedSheetId) {
    try {
      const ss = SpreadsheetApp.openById(storedSheetId);
      if (ss) return ss;
    } catch(e) {
      Logger.log('Could not open spreadsheet directly via stored ID (' + storedSheetId + '): ' + e.message);
    }
  }

  // 2. Direct Folder Resolution: Open from t.FolderID if present
  if (t.FolderID) {
    try {
      const folder = DriveApp.getFolderById(String(t.FolderID).trim());
      if (folder) {
        const code = t.Code || t.ID || cleanId;
        let fileIter = folder.getFilesByName('Training Data');
        if (!fileIter.hasNext()) fileIter = folder.getFilesByName(`${code} Training Data`);
        if (fileIter.hasNext()) {
          return SpreadsheetApp.openById(fileIter.next().getId());
        }
      }
    } catch(fErr) {}
  }

  // 3. Fallback: Drive search by folder name (if TRAINING_FOLDER or ROOT_FOLDER_ID configured)
  try {
    const configuredFolderId = getConfigProperty('TRAINING_FOLDER', '') || getConfigProperty('TRAINING_FOLDER_ID', '');
    let trainingRoot = null;
    if (configuredFolderId) {
      try { trainingRoot = DriveApp.getFolderById(configuredFolderId); } catch(e) {}
    }

    if (!trainingRoot) {
      const rootId = getConfigProperty('ROOT_FOLDER_ID', '');
      if (rootId) {
        try {
          const systemRoot = DriveApp.getFolderById(rootId);
          let trainingRootIter = systemRoot.getFoldersByName('Training Folder');
          if (trainingRootIter.hasNext()) trainingRoot = trainingRootIter.next();
        } catch(rErr) {}
      }
    }

    if (trainingRoot) {
      const code = t.Code || t.ID || cleanId;
      const folderName = `${code} ${t.Name || ''}`.trim();
      let folderIter = trainingRoot.getFoldersByName(folderName);
      let targetFolder = folderIter.hasNext() ? folderIter.next() : null;

      if (!targetFolder) {
        const allSubFolders = trainingRoot.getFolders();
        while (allSubFolders.hasNext()) {
          const f = allSubFolders.next();
          if (f.getName().startsWith(code) || (t.Name && f.getName().includes(t.Name))) {
            targetFolder = f;
            break;
          }
        }
      }

      if (targetFolder) {
        let fileIter = targetFolder.getFilesByName('Training Data');
        if (!fileIter.hasNext()) fileIter = targetFolder.getFilesByName(`${code} Training Data`);
        if (fileIter.hasNext()) {
          return SpreadsheetApp.openById(fileIter.next().getId());
        }
      }
    }
  } catch (e) {
    Logger.log('Error opening per-training sheet: ' + e.message);
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
      file.moveTo(folder);
    }
  }

  const tabDefs = [
    {
      name: 'Participants',
      headers: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt', 'SupervisorID', 'SupervisorEmail', 'SupervisorName']
    },
    {
      name: 'Sessions',
      headers: ['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate']
    },
    {
      name: 'Attendance',
      headers: ['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName', 'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours', 'Remarks', 'EditedBy', 'EditedAt']
    },
    {
      name: 'Evaluation',
      headers: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt']
    },
    {
      name: 'Post Evaluation',
      headers: ['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID', 'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply', 'FurtherTraining', 'Comments', 'SubmittedAt']
    }
  ];

  tabDefs.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      const allSheets = ss.getSheets();
      sheet = allSheets.find(s => s.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === def.name.toLowerCase().replace(/[^a-z0-9]/g, ''));
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
      ['Total Enrolled Participants', '=IF(ISREF(Participants!A2), COUNTA(Participants!A2:A), 0)'],
      ['Total Sessions Created', '=IF(ISREF(Sessions!A2), COUNTA(Sessions!A2:A), 0)'],
      ['Total Attendance Logs', '=IF(ISREF(Attendance!A2), COUNTA(Attendance!A2:A), 0)'],
      ['Present Count', '=IF(ISREF(Attendance!H2), COUNTIF(Attendance!H2:H, "Present"), 0)'],
      ['Total Evaluations Submitted', '=IF(ISREF(Evaluation!A2), COUNTA(Evaluation!A2:A), 0)'],
      ['Overall Average Score', '=IF(AND(ISREF(Evaluation!O2), COUNTA(Evaluation!O2:O)>0), AVERAGE(Evaluation!O2:O), 0)'],
      ['Total Post-Reviews Completed', '=IF(ISREF(\'Post Evaluation\'!A2), COUNTA(\'Post Evaluation\'!A2:A), 0)']
    ]);
  }

  SpreadsheetApp.flush();
  return file;
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
