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

function getSpreadsheet() {
  const ssId = getSpreadsheetId();
  if (ssId) {
    return SpreadsheetApp.openById(ssId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  if (!ss) return null;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    const allSheets = ss.getSheets();
    const targetClean = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
    sheet = allSheets.find(s => s.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === targetClean);
  }
  return sheet;
}

function sheetToJson(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, colIndex) => {
      let val = data[i][colIndex];
      if (val instanceof Date) {
        val = formatDate(val);
      }
      obj[h] = val;
    });
    obj._row = i + 1;
    rows.push(obj);
  }
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
