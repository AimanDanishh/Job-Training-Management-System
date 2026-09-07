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

function getSystemLogoUrl() {
  const url = getConfigProperty('SYSTEM_LOGO_URL', '');
  if (url && String(url).trim() !== '') return String(url).trim();
  return '';
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

/**
 * Safely extracts a 25+ character Google Drive / Spreadsheet ID from raw IDs or full URLs.
 */
function extractSpreadsheetId(input) {
  if (!input) return '';
  const str = String(input).trim();
  const urlMatch = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (urlMatch && urlMatch[1]) return urlMatch[1];
  const dMatch = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (dMatch && dMatch[1]) return dMatch[1];
  const idMatch = str.match(/id=([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) return idMatch[1];
  if (/^[a-zA-Z0-9_-]{25,}$/.test(str)) return str;
  return str;
}

/**
 * Safely extracts a Google Drive folder ID from raw IDs or full URLs.
 */
function extractFolderId(input) {
  if (!input) return '';
  const str = String(input).trim();
  const folderMatch = str.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch && folderMatch[1]) return folderMatch[1];
  const dMatch = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (dMatch && dMatch[1]) return dMatch[1];
  const idMatch = str.match(/id=([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) return idMatch[1];
  if (/^[a-zA-Z0-9_-]{25,}$/.test(str)) return str;
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
  if (!s1 || !s2) return false;

  // 1. Strip leading zeros: e.g. "00123" vs "123"
  const z1 = s1.replace(/^0+/, '');
  const z2 = s2.replace(/^0+/, '');
  if (z1 !== '' && z1 === z2) return true;
  if (s1.replace(/0/g, '') === '' && s2.replace(/0/g, '') === '') return true;

  // 2. Strip common prefixes: e.g. "EMP-00123", "STAFF-00123" vs "00123"
  const cleanPrefix = (str) => str.replace(/^(emp|staff|no|e)[\-_:\s]*/i, '').replace(/^0+/, '');
  const p1 = cleanPrefix(s1);
  const p2 = cleanPrefix(s2);
  if (p1 !== '' && p1 === p2) return true;

  return false;
}

function isSameSessionId(id1, id2) {
  if (id1 === null || id1 === undefined || id2 === null || id2 === undefined) return false;
  let s1 = String(id1).trim().toLowerCase();
  let s2 = String(id2).trim().toLowerCase();

  // If passed a URL or query string, extract session value
  if (s1.includes('session=') || s1.includes('sessionid=')) {
    const m = s1.match(/[?&](?:session|sessionid|id|s)=([^&#\s]+)/i);
    if (m && m[1]) s1 = decodeURIComponent(m[1]).trim().toLowerCase();
  }
  if (s2.includes('session=') || s2.includes('sessionid=')) {
    const m = s2.match(/[?&](?:session|sessionid|id|s)=([^&#\s]+)/i);
    if (m && m[1]) s2 = decodeURIComponent(m[1]).trim().toLowerCase();
  }

  if (s1 === s2) return true;
  if (!s1 || !s2) return false;

  // 1. Compare alphanumeric only (removes dashes, spaces, underscores: e.g. "SES-0001" vs "SES0001" vs "SES 0001")
  const c1 = s1.replace(/[^a-z0-9]/g, '');
  const c2 = s2.replace(/[^a-z0-9]/g, '');
  if (c1 !== '' && c1 === c2) return true;

  // 2. Strip "SES" prefix and leading zeros (e.g. "SES0001" vs "1" or "SES-1")
  const n1 = c1.replace(/^ses/i, '').replace(/^0+/, '');
  const n2 = c2.replace(/^ses/i, '').replace(/^0+/, '');
  if (n1 !== '' && n1 === n2) return true;

  // 3. Both are zero (e.g. "SES0000" vs "0")
  if (c1.replace(/^ses/i, '').replace(/^0+/, '') === '' && c2.replace(/^ses/i, '').replace(/^0+/, '') === '') return true;

  return false;
}

function isSameTrainingId(id1, id2) {
  if (id1 === null || id1 === undefined || id2 === null || id2 === undefined) return false;
  const s1 = String(id1).trim().toLowerCase();
  const s2 = String(id2).trim().toLowerCase();
  if (s1 === s2) return true;
  if (!s1 || !s2) return false;

  const c1 = s1.replace(/[^a-z0-9]/g, '');
  const c2 = s2.replace(/[^a-z0-9]/g, '');
  if (c1 !== '' && c1 === c2) return true;

  const n1 = c1.replace(/^(trn|tr|training)[\-_:\s]*/i, '').replace(/^0+/, '');
  const n2 = c2.replace(/^(trn|tr|training)[\-_:\s]*/i, '').replace(/^0+/, '');
  if (n1 !== '' && n1 === n2) return true;

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
  const isEmpSheet = ['employees', 'cost centre', 'costcentre', 'hod email', 'hodemail', 'csuite email', 'csuiteemail', 'c-suite email', 'hohr email', 'hohremail', 'for it', 'forit', 'for_it', 'staff', 'staff list'].includes(String(name).toLowerCase().trim());
  const primarySs = isEmpSheet ? getEmployeeSpreadsheet() : getSpreadsheet();
  const secondarySs = isEmpSheet ? getSpreadsheet() : getEmployeeSpreadsheet();

  const findInSs = (ssObj) => {
    if (!ssObj) return null;
    let sheet = ssObj.getSheetByName(name);
    if (!sheet) {
      const allSheets = ssObj.getSheets();
      const targetClean = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
      sheet = allSheets.find(s => {
        const sClean = s.getName().toLowerCase().replace(/[^a-z0-9]/g, '');
        return sClean === targetClean ||
               sClean === targetClean + 's' ||
               sClean + 's' === targetClean;
      });

      if (!sheet && isEmpSheet && allSheets.length > 0) {
        sheet = allSheets.find(s => {
          const n = s.getName().toLowerCase();
          return n.includes('for it') || n.includes('emp') || n.includes('staff') || n.includes('master');
        });
      }
    }
    return sheet;
  };

  let found = findInSs(primarySs);
  if (!found) found = findInSs(secondarySs);

  if (found) {
    const cleanName = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
    if (cleanName === 'attendance') {
      ensureAttendanceSheetColumns(found);
    }
  }

  return found;
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

  // Direct Resolution: If passed an existing 25+ character Drive Spreadsheet ID directly or full URL
  if (typeof trainingOrId === 'string') {
    const extractedDirectId = extractSpreadsheetId(trainingOrId);
    if (extractedDirectId && /^[a-zA-Z0-9_-]{25,}$/.test(extractedDirectId)) {
      try {
        const directSs = SpreadsheetApp.openById(extractedDirectId);
        if (directSs) return directSs;
      } catch(e) {}
    }
  }

  let t = (typeof trainingOrId === 'object' && trainingOrId !== null) ? trainingOrId : null;
  const cleanId = String(t ? (t.ID || t.TrainingID || t.Code || '') : trainingOrId).trim();

  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) return null;

  if (!t) {
    const trainings = sheetToJson(tSheet);
    t = trainings.find(r => {
      const id = String(r.ID || r['Training ID'] || r.TrainingID || '').trim();
      const code = String(r.Code || r['Training Code'] || r.TrainingCode || '').trim();
      const tId = String(r.TrainingID || '').trim();
      return isSameTrainingId(id, cleanId) || isSameTrainingId(code, cleanId) || isSameTrainingId(tId, cleanId);
    });
  }
  if (!t) return null;

  // 1. Direct Resolution: Open by stored ParticipantsSheetID / SessionsSheetID / singleSheetId / AttendanceSheetID
  const rawSheetId = String(
    t.ParticipantsSheetID || t['ParticipantsSheetID'] || t['Participants Sheet ID'] || t['Participants Sheet'] ||
    t.SessionsSheetID || t['SessionsSheetID'] || t['Sessions Sheet ID'] || t['Sessions Sheet'] ||
    t.AttendanceSheetID || t['AttendanceSheetID'] || t['Attendance Sheet ID'] ||
    t.singleSheetId || t['singleSheetId'] ||
    t.TrainingDataSheetID || t['TrainingDataSheetID'] || t['Training Data Sheet ID'] ||
    t.TrainingDataID || t['TrainingDataID'] ||
    t.SheetID || t['SheetID'] || t['Spreadsheet ID'] || t['SpreadsheetID'] ||
    ''
  ).trim();
  const storedSheetId = extractSpreadsheetId(rawSheetId);
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
      const folderId = extractFolderId(t.FolderID) || String(t.FolderID).trim();
      const folder = DriveApp.getFolderById(folderId);
      if (folder) {
        const code = t.Code || t.ID || cleanId;
        let fileIter = folder.getFilesByName('Training Data');
        if (!fileIter.hasNext()) fileIter = folder.getFilesByName(`${code} Training Data`);
        if (fileIter.hasNext()) {
          return SpreadsheetApp.openById(fileIter.next().getId());
        }
        // Fallback: Check any Google Sheet file in the folder
        const allFiles = folder.getFiles();
        while (allFiles.hasNext()) {
          const f = allFiles.next();
          if (f.getMimeType() === MimeType.GOOGLE_SHEETS || f.getName().includes('Training') || f.getName().includes('Data')) {
            return SpreadsheetApp.openById(f.getId());
          }
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
  let cleanSessionId = String(sessionId).trim();

  // If a full URL or query string was passed as sessionId
  if (cleanSessionId.includes('http://') || cleanSessionId.includes('https://') || cleanSessionId.includes('?')) {
    const urlMatch = cleanSessionId.match(/[?&](?:session|sessionId|id|s)=([^&#\s]+)/i);
    if (urlMatch && urlMatch[1]) {
      cleanSessionId = decodeURIComponent(urlMatch[1]).trim();
    }
  }

  const extractSessionIdFromRow = (r) => {
    if (!r || typeof r !== 'object') return '';
    return String(
      r.SessionID || r['Session ID'] || r['SessionId'] || r['Session_ID'] ||
      r.ID || r['ID'] || r.SessionCode || r['Session Code'] || r.Code || r.Session || ''
    ).trim();
  };

  const findSessionSheetInSpreadsheet = (ss) => {
    if (!ss) return null;
    const targetNames = ['trainingsessions', 'sessions', 'trainingsession', 'session'];
    const allSheets = ss.getSheets();
    for (const name of targetNames) {
      const found = allSheets.find(s => s.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === name);
      if (found) return found;
    }
    // Check if sheet name contains "session"
    const fuzzyFound = allSheets.find(s => s.getName().toLowerCase().includes('session'));
    if (fuzzyFound) return fuzzyFound;

    // Fallback: check if row 1 headers contain 'SessionID' or 'Session ID'
    for (const s of allSheets) {
      try {
        if (s.getLastRow() >= 1 && s.getLastColumn() >= 1) {
          const firstRow = s.getRange(1, 1, 1, Math.min(s.getLastColumn(), 15)).getValues()[0];
          const hasSessionCol = firstRow.some(h => {
            const clean = String(h).toLowerCase().replace(/[^a-z0-9]/g, '');
            return clean === 'sessionid' || clean === 'sessioncode';
          });
          if (hasSessionCol) return s;
        }
      } catch(e) {}
    }
    return null;
  };

  const searchSessionInSheet = (sheet) => {
    if (!sheet) return null;
    try {
      // 1. Structured JSON row lookup
      const rows = sheetToJson(sheet);
      for (const r of rows) {
        const sId = extractSessionIdFromRow(r);
        if (isSameSessionId(sId, cleanSessionId)) {
          return r;
        }
      }

      // 2. Raw 2D grid fallback: scans every cell in data range if header parsing missed it
      const data = sheet.getDataRange().getValues();
      if (data.length >= 2) {
        let headerRowIdx = 0;
        for (let r = 0; r < Math.min(data.length, 5); r++) {
          const stringCells = data[r].filter(c => typeof c === 'string' && c.trim().length > 0);
          if (stringCells.length >= 2) {
            headerRowIdx = r;
            break;
          }
        }
        const headers = data[headerRowIdx].map(h => String(h || '').trim());
        for (let r = headerRowIdx + 1; r < data.length; r++) {
          const rowValues = data[r];
          if (rowValues.every(c => c === '' || c === null || c === undefined)) continue;
          
          let matched = false;
          for (let c = 0; c < rowValues.length; c++) {
            const cellVal = String(rowValues[c] || '').trim();
            if (isSameSessionId(cellVal, cleanSessionId)) {
              matched = true;
              break;
            }
          }
          if (matched) {
            const rowObj = {};
            headers.forEach((h, idx) => {
              if (h) rowObj[h] = rowValues[idx] !== undefined ? String(rowValues[idx]) : '';
            });
            if (!rowObj.SessionID) rowObj.SessionID = cleanSessionId;
            if (!rowObj.SessionName) rowObj.SessionName = rowObj['Session Name'] || 'Session Check-In';
            return rowObj;
          }
        }
      }
    } catch(err) {
      Logger.log('searchSessionInSheet warning: ' + err.message);
    }
    return null;
  };

  // Tier 1: Check Central Database Spreadsheet & Employee Spreadsheet
  try {
    const mainSs = getSpreadsheet();
    const candidateSpreadsheets = [mainSs];
    const empSs = getEmployeeSpreadsheet();
    if (empSs && (!mainSs || empSs.getId() !== mainSs.getId())) {
      candidateSpreadsheets.push(empSs);
    }

    for (const ss of candidateSpreadsheets) {
      if (!ss) continue;
      const centralSessSheet = findSessionSheetInSpreadsheet(ss);
      if (centralSessSheet) {
        const session = searchSessionInSheet(centralSessSheet);
        if (session) {
          // Resolve parent training (if available)
          let t = null;
          try {
            const tSheet = getSheet(SHEET_NAMES.trainings);
            const trainings = tSheet ? sheetToJson(tSheet) : [];
            const sessionTId = String(session.TrainingID || session['Training ID'] || session.trainingId || '').trim();
            t = trainings.find(r => {
              const id = String(r.ID || r['Training ID'] || r.TrainingID || '').trim();
              const code = String(r.Code || r['Training Code'] || r.TrainingCode || '').trim();
              const tId = String(r.TrainingID || '').trim();
              return isSameTrainingId(id, sessionTId) ||
                     isSameTrainingId(code, sessionTId) ||
                     isSameTrainingId(tId, sessionTId);
            });
          } catch(tErr) {
            Logger.log('Parent training lookup warning in findTrainingBySessionId: ' + tErr.message);
          }

          // CRITICAL: Even if parent training record was not found in Trainings sheet,
          // NEVER return null! We already found the verified session!
          const trainingObj = t || {
            ID: session.TrainingID || session['Training ID'] || 'TRN-0000',
            Name: session.SessionName || session['Session Name'] || session.TrainingTitle || 'Training Programme',
            TrainingTitle: session.SessionName || session['Session Name'] || session.TrainingTitle || 'Training Programme',
            Code: session.TrainingCode || session['Training Code'] || session.TrainingID || session['Training ID'] || ''
          };

          const perTrainingSs = (t && (t.ID || t.Code)) ? (getTrainingDataSpreadsheet(t) || ss) : ss;
          return { session: session, training: trainingObj, spreadsheet: perTrainingSs, sessionSheet: centralSessSheet };
        }
      }
    }
  } catch(eCentral) {
    Logger.log('findTrainingBySessionId Tier 1 error: ' + eCentral.message);
  }

  // Tier 2: Check per-training spreadsheets
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet) {
      const trainings = sheetToJson(tSheet);
      for (const t of trainings) {
        if (!t.ID && !t.Code && !t.TrainingID && !t['Training ID']) continue;
        const ss = getTrainingDataSpreadsheet(t);
        if (!ss) continue;
        const sessSheet = findSessionSheetInSpreadsheet(ss);
        if (!sessSheet) continue;
        const session = searchSessionInSheet(sessSheet);
        if (session) {
          return { session: session, training: t, spreadsheet: ss, sessionSheet: sessSheet };
        }
      }
    }
  } catch(eTier2) {
    Logger.log('findTrainingBySessionId Tier 2 error: ' + eTier2.message);
  }

  // Tier 3: Drive Search for recent Training Data spreadsheets
  try {
    const fileIter = DriveApp.searchFiles('mimeType = "application/vnd.google-apps.spreadsheet" and title contains "Training Data"');
    let count = 0;
    while (fileIter.hasNext() && count < 8) {
      count++;
      const file = fileIter.next();
      try {
        const ss = SpreadsheetApp.openById(file.getId());
        const sessSheet = findSessionSheetInSpreadsheet(ss);
        if (sessSheet) {
          const session = searchSessionInSheet(sessSheet);
          if (session) {
            const inferredName = file.getName().replace(/ Training Data.*/i, '').trim();
            return {
              session: session,
              training: {
                ID: session.TrainingID || inferredName,
                Name: session.SessionName || inferredName,
                TrainingTitle: session.SessionName || inferredName,
                Code: session.TrainingCode || inferredName
              },
              spreadsheet: ss,
              sessionSheet: sessSheet
            };
          }
        }
      } catch(fErr) {}
    }
  } catch(eTier3) {
    Logger.log('findTrainingBySessionId Tier 3 error: ' + eTier3.message);
  }

  Logger.log('findTrainingBySessionId: Session not found for ID: ' + cleanSessionId);
  return null;
}

// ─── ID Generation & Header Normalization ─────────────────────────────────────
function generateId(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
}

function normalizeHeader(header) {
  const h = String(header).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  if (['empid', 'employeeid', 'employeeno', 'staffid', 'badgenumber', 'staffno', 'employeenumber', 'badgeno'].includes(h)) return 'EmployeeID';
  if (h === 'id') return 'ID';
  if (['employeename', 'staffname'].includes(h)) return 'EmployeeName';
  if (['name', 'fullname'].includes(h)) return 'Name';
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

    // Employee ID alias cross-population
    const detectedEmpId = obj.EmployeeID || obj['Employee ID'] || obj.EmployeeNo || obj['Employee No'] || obj['Employee Number'] || obj.EmpID || obj['Emp ID'] || obj.StaffID || obj['Staff ID'] || obj.StaffNo || obj.EmpNo || obj['Emp No'] || obj.BadgeNo || obj['Badge No'] || '';
    if (detectedEmpId) {
      if (!obj.EmployeeID) obj.EmployeeID = detectedEmpId;
      if (!obj.EmployeeNo) obj.EmployeeNo = detectedEmpId;
      if (!obj.EmpID) obj.EmpID = detectedEmpId;
      if (!obj.ID && !headers.some(h => String(h).trim().toLowerCase() === 'id')) obj.ID = detectedEmpId;
    } else if (obj.ID && !String(obj.ID).startsWith('PRT') && !String(obj.ID).startsWith('ATT') && !String(obj.ID).startsWith('EVL') && !String(obj.ID).startsWith('PEV') && !String(obj.ID).startsWith('SES') && !String(obj.ID).startsWith('TRN')) {
      if (!obj.EmployeeID) obj.EmployeeID = obj.ID;
      if (!obj.EmployeeNo) obj.EmployeeNo = obj.ID;
      if (!obj.EmpID) obj.EmpID = obj.ID;
    }

    // Name alias cross-population
    const detectedName = obj.EmployeeName || obj['Employee Name'] || obj.Name || obj['Staff Name'] || obj.FullName || obj['Full Name'] || '';
    if (detectedName) {
      if (!obj.Name) obj.Name = detectedName;
      if (!obj.EmployeeName) obj.EmployeeName = detectedName;
    }

    // Cost Centre <-> Department cross-mapping
    if (obj.CostCentre && !obj.Department) obj.Department = obj.CostCentre;
    if (obj.Department && !obj.CostCentre) obj.CostCentre = obj.Department;
    if (obj['Cost Centre'] && !obj.Department) obj.Department = obj['Cost Centre'];
    if (obj['Cost Centre'] && !obj.CostCentre) obj.CostCentre = obj['Cost Centre'];

    // Position <-> Job Title cross-mapping
    if (obj.Position && !obj.JobTitle) obj.JobTitle = obj.Position;
    if (obj.JobTitle && !obj.Position) obj.Position = obj.JobTitle;
    if (obj.PositionTitle && !obj.Position) obj.Position = obj.PositionTitle;
    if (obj['Position Title'] && !obj.Position) obj.Position = obj['Position Title'];
    if (obj['Job Title'] && !obj.Position) obj.Position = obj['Job Title'];

    // Session ID & Training ID alias cross-population
    const detectedSessionId = obj.SessionID || obj['Session ID'] || obj['Session Id'] || obj['Session_ID'] || obj.SessionCode || obj['Session Code'] || (String(obj.ID || '').toUpperCase().startsWith('SES') ? obj.ID : '');
    if (detectedSessionId) {
      if (!obj.SessionID) obj.SessionID = detectedSessionId;
      if (!obj['Session ID']) obj['Session ID'] = detectedSessionId;
    }

    const detectedTrainingId = obj.TrainingID || obj['Training ID'] || obj['Training Id'] || obj['Training_ID'] || (String(obj.ID || '').toUpperCase().startsWith('TRN') ? obj.ID : '');
    if (detectedTrainingId) {
      if (!obj.TrainingID) obj.TrainingID = detectedTrainingId;
      if (!obj['Training ID']) obj['Training ID'] = detectedTrainingId;
    }

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
