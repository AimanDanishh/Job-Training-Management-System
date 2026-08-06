/**
 * Helper.gs — Employee Requisition Shared Utilities & Database Helper
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

let _cachedSpreadsheet = null;
function getSpreadsheet() {
  if (_cachedSpreadsheet) return _cachedSpreadsheet;
  const ssId = getSpreadsheetId();
  if (ssId) {
    try {
      _cachedSpreadsheet = SpreadsheetApp.openById(ssId);
      return _cachedSpreadsheet;
    } catch (e) {
      Logger.log('Error opening spreadsheet by ID: ' + e.message);
    }
  }
  _cachedSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return _cachedSpreadsheet;
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

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
}

function generateId(prefix) {
  const p = prefix || 'TRN';
  const num = Math.floor(1000 + Math.random() * 9000);
  return `${p}-${num}`;
}

function generateTrainingCode(category) {
  const cat = String(category || 'General').toUpperCase();
  let prefix = 'TR';
  if (cat.includes('LEAD') || cat.includes('MANAG')) prefix = 'LM';
  else if (cat.includes('COMPL') || cat.includes('REGUL')) prefix = 'CR';
  else if (cat.includes('TECH') || cat.includes('IT')) prefix = 'TS';
  else if (cat.includes('SAFE') || cat.includes('QUAL')) prefix = 'SQ';

  const year = new Date().getFullYear();
  const num = String(Math.floor(1 + Math.random() * 9999)).padStart(4, '0');
  return `${prefix}-${year}-${num}`;
}

function ok(data) {
  return JSON.stringify({ success: true, data: data });
}

function err(message) {
  return JSON.stringify({ success: false, message: message });
}
