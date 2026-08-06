/**
 * Helper.gs — HOD Portal Shared Utilities & Spreadsheet Services
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


function getSpreadsheetId() {
  return getConfigProperty('SPREADSHEET_ID', '');
}

function getSpreadsheet() {
  const ssId = getSpreadsheetId();
  if (ssId) {
    return SpreadsheetApp.openById(ssId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getEmployeeSpreadsheetId() {
  return getConfigProperty('EMPLOYEE_SPREADSHEET_ID', getSpreadsheetId());
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
  const isEmpSheet = ['employees', 'cost centre', 'costcentre', 'hod email', 'hodemail', 'for it', 'forit', 'for_it'].includes(String(name).toLowerCase().trim());
  const ss = isEmpSheet ? getEmployeeSpreadsheet() : getSpreadsheet();
  if (!ss) return null;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    const allSheets = ss.getSheets();
    const targetClean = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
    sheet = allSheets.find(s => s.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === targetClean);
    if (!sheet && allSheets.length > 0 && isEmpSheet) {
      sheet = allSheets[0];
    }
  }
  return sheet;
}

let _sheetDataCache = {};
function sheetToJson(sheet) {
  if (!sheet) return [];
  const sheetName = sheet.getName();
  if (_sheetDataCache[sheetName]) return _sheetDataCache[sheetName];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map(h => String(h).trim());
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row.some(cell => cell !== '' && cell !== null)) continue;

    const obj = {};
    headers.forEach((h, colIndex) => {
      let val = row[colIndex];
      if (val instanceof Date) {
        val = formatDate(val);
      }
      obj[h] = val;
    });
    obj._row = i + 1;
    rows.push(obj);
  }

  _sheetDataCache[sheetName] = rows;
  return rows;
}

function findRowById(sheet, id) {
  if (!sheet) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) return i + 1;
  }
  return -1;
}

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateTime(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss');
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

  const tSheet = getSheet('Trainings');
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
      }
    } catch (e) {
      Logger.log('Error opening per-training sheet from FolderID: ' + e.message);
    }
  }

  return null;
}

function ok(data) {
  return JSON.stringify({ success: true, data: data });
}

function err(message) {
  return JSON.stringify({ success: false, message: message });
}
