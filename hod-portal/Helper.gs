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

/**
 * Returns configured Employee Portal URL from Script Properties
 */
function getEmployeePortalUrl() {
  let url = getConfigProperty('EMPLOYEE_PORTAL_URL', '') || getConfigProperty('EMPLOYEE_REQUISITION_URL', '');
  return url;
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
  const isEmpSheet = ['employees', 'cost centre', 'costcentre', 'hod email', 'hodemail', 'csuite email', 'csuiteemail', 'c-suite email', 'hohr email', 'hohremail', 'for it', 'forit', 'for_it'].includes(String(name).toLowerCase().trim());
  const primarySs = isEmpSheet ? getEmployeeSpreadsheet() : getSpreadsheet();

  const secondarySs = isEmpSheet ? getSpreadsheet() : getEmployeeSpreadsheet();

  const findInSs = (ssObj) => {
    if (!ssObj) return null;
    let sheet = ssObj.getSheetByName(name);
    if (!sheet) {
      const allSheets = ssObj.getSheets();
      const targetClean = String(name).toLowerCase().replace(/[^a-z0-9]/g, '');
      sheet = allSheets.find(s => s.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === targetClean);
    }
    return sheet;
  };

  let found = findInSs(primarySs);
  if (!found) found = findInSs(secondarySs);
  return found;
}


let _sheetDataCache = {};
function sheetToJson(sheet) {
  if (!sheet) return [];

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

  return rows;
}

function findRowById(sheet, id) {
  if (!sheet || id === null || id === undefined) return -1;
  const cleanId = String(id).trim().toLowerCase();
  if (!cleanId) return -1;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return -1;

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idColIdxs = [];
  headers.forEach((h, idx) => {
    if (['id', 'code', 'trainingid', 'training id', 'requisitionid', 'requisition id', 'trainingcode', 'training code'].indexOf(h) !== -1) {
      idColIdxs.push(idx);
    }
  });
  if (idColIdxs.length === 0) idColIdxs.push(0);

  for (let i = 1; i < data.length; i++) {
    for (let c = 0; c < idColIdxs.length; c++) {
      const colIdx = idColIdxs[c];
      if (String(data[i][colIdx]).trim().toLowerCase() === cleanId) {
        return i + 1;
      }
    }
  }
  return -1;
}

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


function formatDateTime(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
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

  // If trainingId is already a training object with ParticipantsSheetID / FolderID
  if (typeof trainingId === 'object' && trainingId !== null) {
    if (trainingId.ParticipantsSheetID) {
      try {
        return SpreadsheetApp.openById(String(trainingId.ParticipantsSheetID).trim());
      } catch(e) {}
    }
    if (trainingId.FolderID) {
      try {
        const folder = DriveApp.getFolderById(String(trainingId.FolderID).trim());
        let fileIter = folder.getFilesByName('Training Data');
        if (!fileIter.hasNext()) fileIter = folder.getFilesByName(`${trainingId.Code || trainingId.ID} Training Data`);
        if (fileIter.hasNext()) {
          return SpreadsheetApp.openById(fileIter.next().getId());
        }
      } catch(e) {}
    }
  }

  const cleanId = String(trainingId).trim().toLowerCase();
  const tSheet = getSheet('Trainings');
  if (!tSheet) return null;

  const trainings = sheetToJson(tSheet);
  const t = trainings.find(r => {
    const id = String(r.ID || '').trim().toLowerCase();
    const code = String(r.Code || '').trim().toLowerCase();
    const tId = String(r.TrainingID || '').trim().toLowerCase();
    return id === cleanId || code === cleanId || tId === cleanId;
  });
  if (!t) return null;

  // 1. Direct open by ParticipantsSheetID / SpreadsheetID if present
  if (t.ParticipantsSheetID) {
    try {
      return SpreadsheetApp.openById(String(t.ParticipantsSheetID).trim());
    } catch(e) {}
  }

  // 2. Direct open from t.FolderID if present
  if (t.FolderID) {
    try {
      const folder = DriveApp.getFolderById(String(t.FolderID).trim());
      let fileIter = folder.getFilesByName('Training Data');
      if (!fileIter.hasNext()) fileIter = folder.getFilesByName(`${t.Code || t.ID} Training Data`);
      if (fileIter.hasNext()) {
        return SpreadsheetApp.openById(fileIter.next().getId());
      }
    } catch(e) {}
  }

  // 3. Fallback: Drive search by folder name
  try {
      const configuredFolderId = getConfigProperty('TRAINING_FOLDER', '') || getConfigProperty('TRAINING_FOLDER_ID', '');
      let trainingRoot = null;
      if (configuredFolderId) {
        try { trainingRoot = DriveApp.getFolderById(configuredFolderId); } catch(e) {}
      }
      if (!trainingRoot) {
        const rootId = getConfigProperty('ROOT_FOLDER_ID', '');
        if (!rootId) throw new Error('ROOT_FOLDER_ID is required.');
        const root = DriveApp.getFolderById(rootId);
        const trainingRoots = root.getFoldersByName('Training Folder');
        if (trainingRoots.hasNext()) trainingRoot = trainingRoots.next();
      }
      if (!trainingRoot) return null;

      const code = t.Code || t.ID || cleanId;
      const folderName = `${code} ${t.Name || ''}`.trim();
      let folders = trainingRoot.getFoldersByName(folderName);
      if (!folders.hasNext()) {
        const allSubFolders = trainingRoot.getFolders();
        while (allSubFolders.hasNext()) {
          const f = allSubFolders.next();
          if (f.getName().startsWith(code) || (t.Name && f.getName().includes(t.Name))) {
            let fileIter = f.getFilesByName('Training Data');
            if (!fileIter.hasNext()) fileIter = f.getFilesByName(`${code} Training Data`);
            if (fileIter.hasNext()) {
              return SpreadsheetApp.openById(fileIter.next().getId());
            }
          }
        }
        return null;
      }
      const folder = folders.next();
      let fileIter = folder.getFilesByName('Training Data');
      if (!fileIter.hasNext()) fileIter = folder.getFilesByName(`${code} Training Data`);
      if (fileIter.hasNext()) {
        return SpreadsheetApp.openById(fileIter.next().getId());
      }
  } catch (e) {
    Logger.log('Error resolving per-training sheet: ' + e.message);
  }

  return null;
}

/**
 * Resolves participant list for a training requisition using robust multi-tier fallbacks:
 * 1. Direct JSON ParticipantList / participants on the training record
 * 2. Direct open of ParticipantsSheetID / Training Data Spreadsheet (Sheet: Participants or TrainingParticipants)
 * 3. Requisition Form Google Spreadsheet (A15:I38 or A23:F46)
 * 4. Central Database Participants sheet
 * 5. Full Employee Master Sheet enrichment (Name, Department, Job Position)
 */
function getParticipantsForRequisition(training, cleanId) {
  let participants = [];
  const tObj = (typeof training === 'object' && training !== null) ? training : {};
  const tId = String(tObj.ID || tObj.TrainingID || tObj.Code || cleanId || '').trim();
  const tCode = String(tObj.Code || tObj.ID || '').trim();

  // Tier 1: Try Direct JSON ParticipantList / participants on the training object
  const rawList = tObj.ParticipantList || tObj.participants || tObj.ParticipantsList;
  if (rawList) {
    try {
      const parsed = Array.isArray(rawList) ? rawList : (typeof rawList === 'string' ? JSON.parse(rawList) : null);
      if (Array.isArray(parsed) && parsed.length > 0) {
        participants = parsed.map(p => ({
          ID: String(p.EmployeeID || p.EmployeeNo || p['Employee No'] || p.ID || '').trim(),
          EmployeeID: String(p.EmployeeID || p.EmployeeNo || p['Employee No'] || p.ID || '').trim(),
          Name: String(p.EmployeeName || p.Name || p['Employee Name'] || '').trim(),
          EmployeeName: String(p.EmployeeName || p.Name || p['Employee Name'] || '').trim(),
          Department: String(p.Department || p.CostCentre || p['Cost Centre'] || '').trim(),
          CostCentre: String(p.Department || p.CostCentre || p['Cost Centre'] || '').trim(),
          Position: String(p.Position || p.JobTitle || p['Job Position'] || '').trim(),
          SupervisorID: String(p.SupervisorID || p.SupervisorId || '').trim(),
          SupervisorName: String(p.SupervisorName || '').trim(),
          SupervisorEmail: String(p.SupervisorEmail || '').trim()
        })).filter(p => p.EmployeeID || p.EmployeeName);
      }
    } catch(eJson) {}
  }

  // Tier 2: Try per-training Training Data Spreadsheet
  if (participants.length === 0) {
    try {
      const ss = getTrainingDataSpreadsheet(training || tId || tCode);
      if (ss) {
        const partSheet = ss.getSheetByName('Participants') || 
                          ss.getSheetByName('TrainingParticipants') || 
                          ss.getSheetByName('ParticipantList') ||
                          ss.getSheets()[0];
        if (partSheet) {
          const rows = sheetToJson(partSheet);
          if (Array.isArray(rows) && rows.length > 0) {
            participants = rows.map(p => ({
              ID: String(p.EmployeeID || p.EmployeeNo || p['Employee No'] || p['Employee ID'] || (p.ID && !String(p.ID).startsWith('TP-') ? p.ID : '') || '').trim(),
              EmployeeID: String(p.EmployeeID || p.EmployeeNo || p['Employee No'] || p['Employee ID'] || (p.ID && !String(p.ID).startsWith('TP-') ? p.ID : '') || '').trim(),
              Name: String(p.EmployeeName || p.Name || p['Employee Name'] || p['Name'] || '').trim(),
              EmployeeName: String(p.EmployeeName || p.Name || p['Employee Name'] || p['Name'] || '').trim(),
              Department: String(p.Department || p.CostCentre || p['Cost Centre'] || p['Department'] || p.Dept || '').trim(),
              CostCentre: String(p.Department || p.CostCentre || p['Cost Centre'] || p['Department'] || p.Dept || '').trim(),
              Position: String(p.Position || p.JobTitle || p['Job Position'] || p['Position'] || p.Designation || '').trim(),
              SupervisorID: String(p.SupervisorID || p.SupervisorId || '').trim(),
              SupervisorName: String(p.SupervisorName || '').trim(),
              SupervisorEmail: String(p.SupervisorEmail || '').trim()
            })).filter(p => p.EmployeeID || p.EmployeeName);
          }
        }
      }
    } catch (e1) {
      Logger.log('Error reading from Training Data spreadsheet: ' + e1.message);
    }
  }

  // Tier 3: Try Requisition Form Spreadsheet (AP-HRD-F01-01) across all page tabs
  if (participants.length === 0 && (tObj.RequisitionFormFileID || tObj.RequisitionFormId)) {
    const formFileId = String(tObj.RequisitionFormFileID || tObj.RequisitionFormId).trim();
    try {
      const formSs = SpreadsheetApp.openById(formFileId);
      const allFormSheets = formSs.getSheets();
      const extracted = [];

      allFormSheets.forEach(formSheet => {
        const sName = formSheet.getName();
        if (sName.startsWith('Training Form') || sName.includes('Form') || allFormSheets.length === 1) {
          const rangeValues = formSheet.getRange('A14:I40').getValues();
          rangeValues.forEach(row => {
            let colA = String(row[0] || '').trim();
            let colB = String(row[1] || '').trim();
            let colC = String(row[2] || '').trim();
            let colD = String(row[3] || '').trim();
            let colE = String(row[4] || '').trim();
            let colH = String(row[7] || '').trim();

            let empId = '';
            let empName = '';
            let dept = '';
            let pos = '';

            if (/^\d{1,2}$/.test(colA) && (colB || colC)) {
              empId = colB;
              empName = colC;
              dept = colD;
              pos = colE || colH;
            } else if (colA && !/^\d{1,2}$/.test(colA)) {
              empId = colA;
              empName = colC || colB;
              dept = colD;
              pos = colH || colE;
            } else if (colC) {
              empName = colC;
              empId = colB || colA;
              dept = colD;
              pos = colH;
            }

            if (empId || empName) {
              const lowerId = empId.toLowerCase();
              const lowerName = empName.toLowerCase();
              const isHeader = lowerId.includes('employee') || lowerId.includes('bil') || lowerId.includes('no') ||
                               lowerName.includes('name') || lowerName.includes('nama') || lowerId.includes('name');
              if (!isHeader && (empId || empName)) {
                extracted.push({
                  ID: empId,
                  EmployeeID: empId,
                  Name: empName,
                  EmployeeName: empName,
                  Department: dept,
                  CostCentre: dept,
                  Position: pos
                });
              }
            }
          });
        }
      });

      if (extracted.length > 0) {
        participants = extracted;
      }
    } catch (e2) {
      Logger.log('Error reading participants from Requisition Form: ' + e2.message);
    }
  }

  // Tier 4: Try Central Database sheet 'Participants' or 'TrainingParticipants'
  if (participants.length === 0) {
    try {
      const dbPartSheet = getSheet('Participants') || getSheet('TrainingParticipants');
      if (dbPartSheet) {
        const allDbParts = sheetToJson(dbPartSheet);
        const matched = allDbParts.filter(p => {
          const pTid = String(p.TrainingID || p.TrainingCode || p.ID || '').trim().toLowerCase();
          const target1 = tId.toLowerCase();
          const target2 = tCode.toLowerCase();
          const target3 = cleanId.toLowerCase();
          return pTid && (pTid === target1 || pTid === target2 || pTid === target3);
        });
        if (matched.length > 0) {
          participants = matched.map(p => ({
            ID: String(p.EmployeeID || p.EmployeeNo || p['Employee No'] || (p.ID && !String(p.ID).startsWith('TP-') ? p.ID : '') || '').trim(),
            EmployeeID: String(p.EmployeeID || p.EmployeeNo || p['Employee No'] || (p.ID && !String(p.ID).startsWith('TP-') ? p.ID : '') || '').trim(),
            Name: String(p.EmployeeName || p.Name || p['Employee Name'] || '').trim(),
            EmployeeName: String(p.EmployeeName || p.Name || p['Employee Name'] || '').trim(),
            Department: String(p.Department || p.CostCentre || p['Cost Centre'] || '').trim(),
            CostCentre: String(p.Department || p.CostCentre || p['Cost Centre'] || '').trim(),
            Position: String(p.Position || p.JobTitle || p['Job Position'] || '').trim(),
            SupervisorID: String(p.SupervisorID || p.SupervisorId || '').trim(),
            SupervisorName: String(p.SupervisorName || '').trim(),
            SupervisorEmail: String(p.SupervisorEmail || '').trim()
          })).filter(p => p.EmployeeID || p.EmployeeName);
        }
      }
    } catch (e3) {
      Logger.log('Error reading from central database Participants sheet: ' + e3.message);
    }
  }

  // Tier 5: Enrich with Employees master sheet to guarantee full details
  try {
    const empSheet = getSheet('Employees');
    if (empSheet) {
      const employees = sheetToJson(empSheet);
      if (participants.length > 0) {
        participants = participants.map(p => {
          const empId = p.EmployeeID || p.ID;
          const empName = p.EmployeeName || p.Name;
          const matchedEmp = employees.find(e => {
            const eId = String(e['Employee No'] || e.EmployeeNo || e.ID || e.EmployeeID || '').toLowerCase().trim();
            const eName = String(e.Name || e.EmployeeName || '').toLowerCase().trim();
            return (empId && eId && eId === String(empId).toLowerCase().trim()) ||
                   (empName && eName && eName === String(empName).toLowerCase().trim());
          });
          if (matchedEmp) {
            const resolvedId = empId || String(matchedEmp['Employee No'] || matchedEmp.EmployeeNo || matchedEmp.ID || '').trim();
            const resolvedName = p.EmployeeName || String(matchedEmp.Name || matchedEmp.EmployeeName || '').trim();
            const resolvedDept = p.Department || String(matchedEmp['Cost Centre'] || matchedEmp.Department || '').trim();
            const resolvedPos = p.Position || String(matchedEmp['Job Position'] || matchedEmp.Position || 'Participant').trim();
            return {
              ID: resolvedId,
              EmployeeID: resolvedId,
              Name: resolvedName,
              EmployeeName: resolvedName,
              Department: resolvedDept,
              CostCentre: resolvedDept,
              Position: resolvedPos,
              SupervisorID: p.SupervisorID || '',
              SupervisorName: p.SupervisorName || '',
              SupervisorEmail: p.SupervisorEmail || ''
            };
          }
          return p;
        });
      }
    }
  } catch (e5) {}

  return participants;
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

function ok(data) {
  return JSON.stringify({ success: true, data: data });
}

function err(message) {
  return JSON.stringify({ success: false, message: message });
}

/**
 * Updates digital approval signature blocks on the AP-HRD-F01-01 Training Requisition Form Google Sheet.
 * Format: (request status, employee no., name, job position, date)
 */
function updateTrainingRequisitionSignatures(trainingId, step, sigData, targetFormId) {
  try {
    SpreadsheetApp.flush();
    const trainingSheet = getSheet('Trainings');
    if (!trainingSheet) return;
    const headers = trainingSheet.getDataRange().getValues()[0].map(h => String(h).trim());
    let row = -1;
    const data = trainingSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(trainingId).trim()) { row = i + 1; break; }
    }
    if (row === -1) return;

    let formId = targetFormId || '';
    if (!formId) {
      const formCol = headers.indexOf('RequisitionFormFileID') + 1;
      formId = formCol ? trainingSheet.getRange(row, formCol).getValue() : '';
    }
    if (!formId) return;

    const ss = SpreadsheetApp.openById(formId);
    const allSheets = ss.getSheets();
    const targetSheets = allSheets.filter(s => {
      const sName = s.getName();
      return sName.startsWith('Training Form') || sName.includes('Form') || allSheets.length === 1;
    });

    // Preserve original template header rows 39 and 40 (columns A to I) completely untouched

    const empNo    = sigData.employeeNo || sigData.EmployeeNo || sigData.EmployeeID || sigData.ID || '';
    const empName  = sigData.name || sigData.EmployeeName || sigData.Name || '';
    const position = sigData.position || sigData.JobPosition || sigData.Position || sigData.JobTitle || '';
    const sigDate  = getFormattedCurrentDate(sigData.date || sigData.Date || sigData.Timestamp);
    const status   = sigData.status || sigData.RequestStatus || sigData.ApprovalStatus || '';

    const formatSingleColCell = (prefix, val) => {
      if (!val) return prefix;
      const sVal = String(val).trim();
      if (sVal.toUpperCase().startsWith(prefix.toUpperCase())) return sVal;
      return `${prefix} ${sVal}`;
    };

    const stepNorm = String(step || '').trim().toLowerCase();
    
    targetSheets.forEach(sheet => {
      if (stepNorm === 'request' || stepNorm === 'requested by') {
        sheet.getRange('A41').setValue(formatSingleColCell('EMPLOYEE NO:', empNo));
        sheet.getRange('A42').setValue(formatSingleColCell('NAME:', empName));
        sheet.getRange('A44').setValue(formatSingleColCell('JOB POSITION:', position));
        sheet.getRange('A46').setValue(formatSingleColCell('DATE:', sigDate));

        sheet.getRange('B41').setValue(empNo);
        sheet.getRange('B42').setValue(empName);
        sheet.getRange('B43').setValue(empName);
        sheet.getRange('B44').setValue(position);
        sheet.getRange('B46').setValue(sigDate);
      } else if (stepNorm === 'hod' || stepNorm === 'head of department' || stepNorm === 'verified by head of department') {
        sheet.getRange('C41').setValue(formatSingleColCell('STATUS:', status || 'Verified'));
        sheet.getRange('C42').setValue(formatSingleColCell('EMPLOYEE NO:', empNo));
        sheet.getRange('C43').setValue(formatSingleColCell('NAME:', empName));
        sheet.getRange('C44').setValue(formatSingleColCell('JOB POSITION:', position));
        sheet.getRange('C46').setValue(formatSingleColCell('DATE:', sigDate));
      } else if (stepNorm === 'csuite' || stepNorm === 'c-suite' || stepNorm === 'approved by c-suite') {
        sheet.getRange('D41').setValue(formatSingleColCell('STATUS:', status || 'Approved'));
        sheet.getRange('D42').setValue(formatSingleColCell('EMPLOYEE NO:', empNo));
        sheet.getRange('D43').setValue(formatSingleColCell('NAME:', empName));
        sheet.getRange('D44').setValue(formatSingleColCell('JOB POSITION:', position));
        sheet.getRange('D46').setValue(formatSingleColCell('DATE:', sigDate));

        sheet.getRange('E41').setValue(status || 'Approved');
        sheet.getRange('E42').setValue(empNo);
        sheet.getRange('E43').setValue(empName);
        sheet.getRange('E44').setValue(position);
        sheet.getRange('E46').setValue(sigDate);
      } else if (stepNorm === 'hohr' || stepNorm === 'head of hr' || stepNorm === 'approved by hohr') {
        sheet.getRange('F41').setValue(formatSingleColCell('STATUS:', status || 'Approved'));
        sheet.getRange('F42').setValue(formatSingleColCell('EMPLOYEE NO:', empNo));
        sheet.getRange('F43').setValue(formatSingleColCell('NAME:', empName));
        sheet.getRange('F44').setValue(formatSingleColCell('JOB POSITION:', position));
        sheet.getRange('F46').setValue(formatSingleColCell('DATE:', sigDate));

        sheet.getRange('G41').setValue(status || 'Approved');
        sheet.getRange('G42').setValue(empNo);
        sheet.getRange('G43').setValue(empName);
        sheet.getRange('G44').setValue(position);
        sheet.getRange('G46').setValue(sigDate);
      } else if (stepNorm === 'hr' || stepNorm === 'arina' || stepNorm === 'hr department' || stepNorm === 'acknowledged by hr department') {
        sheet.getRange('H41').setValue(formatSingleColCell('EMPLOYEE NO:', empNo));
        sheet.getRange('H42').setValue(formatSingleColCell('NAME:', empName));
        sheet.getRange('H44').setValue(formatSingleColCell('JOB POSITION:', position));
        sheet.getRange('H46').setValue(formatSingleColCell('DATE:', sigDate));

        sheet.getRange('I41').setValue(empNo);
        sheet.getRange('I42').setValue(empName);
        sheet.getRange('I43').setValue(empName);
        sheet.getRange('I44').setValue(position);
        sheet.getRange('I46').setValue(sigDate);
      }

      // Auto-check: If step is NOT 'request' but cell B41 is empty, ensure requester signature is populated too
      if (stepNorm !== 'request' && stepNorm !== 'requested by') {
        const currentReqVal = sheet.getRange('B41').getValue() || sheet.getRange('B42').getValue();
        if (!currentReqVal || String(currentReqVal).trim() === '') {
          const reqIdCol = headers.indexOf('RequestedBy') + 1;
          const reqNameCol = headers.indexOf('RequestedByName') + 1;
          const createdDateCol = headers.indexOf('CreatedDate') + 1;
          const reqId = reqIdCol > 0 ? trainingSheet.getRange(row, reqIdCol).getValue() : '';
          const reqName = reqNameCol > 0 ? trainingSheet.getRange(row, reqNameCol).getValue() : '';
          const reqDate = (createdDateCol > 0 ? trainingSheet.getRange(row, createdDateCol).getValue() : '') || sigDate;
          let reqPos = 'Requester';
          if (reqId && getSheet('Employees')) {
            const emps = sheetToJson(getSheet('Employees'));
            const m = emps.find(e => String(e.ID || e.EmployeeID).toLowerCase() === String(reqId).toLowerCase());
            if (m) reqPos = m.Position || m.JobTitle || m.PositionTitle || 'Requester';
          }
          if (reqId || reqName) {
            sheet.getRange('B41').setValue(reqId);
            sheet.getRange('B42').setValue(reqName);
            sheet.getRange('B43').setValue(reqName);
            sheet.getRange('B44').setValue(reqPos);
            sheet.getRange('B46').setValue(getFormattedCurrentDate(reqDate));
          }
        }
      }
    });

    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log('updateTrainingRequisitionSignatures error: ' + e.message);
  }
}

/**
 * Utility: Format dates cleanly for corporate email template (e.g. "7 August 2026")
 */
function formatDateForEmail(dateVal) {
  if (!dateVal) return '';
  try {
    let d;
    if (dateVal instanceof Date) {
      d = dateVal;
    } else {
      const str = String(dateVal).trim();
      if (str.includes('/')) {
        const parts = str.split('/');
        if (parts.length === 3) {
          d = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
        }
      }
      if (!d || isNaN(d.getTime())) {
        d = new Date(str);
      }
    }
    if (isNaN(d.getTime())) return String(dateVal);
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
  } catch (e) {
    return String(dateVal);
  }
}

/**
 * Utility: Format course fee for corporate email template (e.g. "RM 2,000.00")
 */
function formatFeeForEmail(feeVal) {
  if (feeVal === undefined || feeVal === null || String(feeVal).trim() === '') return 'RM 0.00';
  let str = String(feeVal).replace(/RM/gi, '').trim();
  let num = parseFloat(str.replace(/,/g, ''));
  if (isNaN(num)) return String(feeVal).startsWith('RM') ? String(feeVal) : `RM ${feeVal}`;
  return 'RM ' + num.toLocaleString('en-MY', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Utility: Format duration for corporate email template (e.g. "1 day · 0.5 hours")
 */
function formatDurationForEmail(durationVal, hoursVal) {
  const d = durationVal || 1;
  const h = hoursVal !== undefined && hoursVal !== null ? hoursVal : (d * 8);
  return `${d} day${d > 1 ? 's' : ''} · ${h} hour${h > 1 ? 's' : ''}`;
}

/**
 * Generates a modern corporate HTML email for Training Requisition approval workflow.
 *
 * @param {Object} params
 * @param {string} params.requestId - Request ID (e.g. TRN-7190)
 * @param {string} params.trainingTitle - Training Name
 * @param {string} params.requesterName - Requester Name
 * @param {string} params.employeeId - Employee ID
 * @param {string} params.department - Department / Cost Centre
 * @param {string} params.category - Training Category
 * @param {string} params.proposedDate - Proposed Date(s)
 * @param {string} params.duration - Duration text
 * @param {string} params.estimatedFee - Formatted fee text
 * @param {string} params.status - Current status text
 * @param {string} params.reviewUrl - Direct review link
 * @param {string} [params.badgeText] - Badge label (default: ACTION REQUIRED)
 * @param {string} [params.headlineText] - Email main headline
 * @param {string} [params.greetingText] - Greeting text
 * @param {string} [params.introText] - Short intro description
 * @returns {string} Fully styled HTML string for email body
 */
function buildTrainingRequisitionEmailHtml(params) {
  const requestId = params.requestId || 'TRN-0000';
  const trainingTitle = params.trainingTitle || 'Training Request';
  const requesterName = params.requesterName || 'Employee Requester';
  const employeeId = params.employeeId || 'N/A';
  const department = params.department || 'N/A';
  const category = params.category || 'General';
  const proposedDate = params.proposedDate || 'N/A';
  const duration = params.duration || '1 day';
  const estimatedFee = params.estimatedFee || 'RM 0.00';
  const status = params.status || 'Pending HOD Approval';
  const reviewUrl = params.reviewUrl || '';
  
  const badgeText = params.badgeText || 'ACTION REQUIRED';
  const headlineText = params.headlineText || 'Training Requisition Requires Your Review';
  const greetingText = params.greetingText || 'Dear Approver / Manager,';
  const introText = params.introText || 'A new training requisition has been submitted and is currently awaiting your review and approval.';

  // Status background & border colors depending on status
  let statusBg = '#FFFBEB';
  let statusBorder = '#F59E0B';
  let statusText = '#78350F';
  let badgeBg = '#FEF3C7';
  let badgeBorder = '#FCD34D';
  let badgeColor = '#92400E';

  const statusLower = String(status).toLowerCase();
  if (statusLower.includes('approved')) {
    statusBg = '#ECFDF5';
    statusBorder = '#10B981';
    statusText = '#065F46';
    badgeBg = '#D1FAE5';
    badgeBorder = '#6EE7B7';
    badgeColor = '#065F46';
  } else if (statusLower.includes('reject')) {
    statusBg = '#FEF2F2';
    statusBorder = '#EF4444';
    statusText = '#991B1B';
    badgeBg = '#FEE2E2';
    badgeBorder = '#FCA5A5';
    badgeColor = '#991B1B';
  } else if (statusLower.includes('return')) {
    statusBg = '#EFF6FF';
    statusBorder = '#3B82F6';
    statusText = '#1E40AF';
    badgeBg = '#DBEAFE';
    badgeBorder = '#93C5FD';
    badgeColor = '#1E40AF';
  }

  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Training Requisition</title>
</head>
<body style="margin: 0; padding: 0; background-color: #EEF1F5; font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif; -webkit-font-smoothing: antialiased; -webkit-text-size-adjust: none; width: 100% !important;">
    <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #EEF1F5; table-layout: fixed; padding: 20px 0;">
        <tr>
            <td align="center" style="padding: 10px;">
                <!-- Main Email Card -->
                <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 680px; background-color: #FFFFFF; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08); border: 1px solid #E3E7EC;">
                    
                    <!-- Header -->
                    <tr>
                        <td style="background-color: #17365D; padding: 28px 32px; text-align: left;">
                            <div style="font-size: 26px; font-weight: 800; color: #FFFFFF; letter-spacing: 1.5px; line-height: 1.2; text-transform: uppercase; margin: 0;">TRAINHUB</div>
                            <div style="font-size: 11px; font-weight: 600; color: rgba(255, 255, 255, 0.75); letter-spacing: 2px; text-transform: uppercase; margin-top: 4px;">Training Management System</div>
                        </td>
                    </tr>

                    <!-- Body Content -->
                    <tr>
                        <td style="padding: 32px 32px 24px 32px;">
                            
                            <!-- Badge -->
                            <div style="margin-bottom: 16px;">
                                <span style="display: inline-block; background-color: ${badgeBg}; color: ${badgeColor}; border: 1px solid ${badgeBorder}; border-radius: 12px; padding: 4px 12px; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;">${badgeText}</span>
                            </div>

                            <!-- Title Headline -->
                            <h1 style="color: #101828; font-size: 20px; font-weight: 700; margin: 0 0 16px 0; line-height: 1.3;">${headlineText}</h1>

                            <!-- Intro Text -->
                            <p style="color: #475467; font-size: 14px; line-height: 1.6; margin: 0 0 8px 0;">${greetingText}</p>
                            <p style="color: #475467; font-size: 14px; line-height: 1.6; margin: 0 0 24px 0;">${introText}</p>

                            <!-- Requisition Card -->
                            <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #FFFFFF; border: 1px solid #E3E7EC; border-radius: 8px; border-collapse: separate; margin-bottom: 20px;">
                                <tr>
                                    <td style="padding: 24px;">
                                        <!-- Header Label -->
                                        <div style="color: #667085; font-size: 11px; font-weight: 700; letter-spacing: 1.2px; text-transform: uppercase; margin-bottom: 8px;">TRAINING REQUISITION</div>
                                        
                                        <!-- Main Prominent Training Title -->
                                        <div style="color: #17365D; font-size: 20px; font-weight: 700; line-height: 1.3; margin-bottom: 4px;">${trainingTitle}</div>
                                        
                                        <!-- Request ID -->
                                        <div style="color: #667085; font-size: 13px; font-weight: 500; margin-bottom: 20px;">Request ID: ${requestId}</div>

                                        <!-- Divider line -->
                                        <div style="border-bottom: 1px solid #E3E7EC; margin-bottom: 20px;"></div>

                                        <!-- Details Table -->
                                        <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr>
                                                <td style="color: #667085; font-size: 13px; font-weight: 600; width: 140px; padding: 6px 12px 6px 0; vertical-align: top;">Requester</td>
                                                <td style="color: #101828; font-size: 14px; font-weight: 500; padding: 6px 0; vertical-align: top;">${requesterName}</td>
                                            </tr>
                                            <tr>
                                                <td style="color: #667085; font-size: 13px; font-weight: 600; padding: 6px 12px 6px 0; vertical-align: top;">Employee ID</td>
                                                <td style="color: #101828; font-size: 14px; font-weight: 500; padding: 6px 0; vertical-align: top;">${employeeId}</td>
                                            </tr>
                                            <tr>
                                                <td style="color: #667085; font-size: 13px; font-weight: 600; padding: 6px 12px 6px 0; vertical-align: top;">Department</td>
                                                <td style="color: #101828; font-size: 14px; font-weight: 500; padding: 6px 0; vertical-align: top;">${department}</td>
                                            </tr>
                                            <tr>
                                                <td style="color: #667085; font-size: 13px; font-weight: 600; padding: 6px 12px 6px 0; vertical-align: top;">Category</td>
                                                <td style="color: #101828; font-size: 14px; font-weight: 500; padding: 6px 0; vertical-align: top;">${category}</td>
                                            </tr>
                                            <tr>
                                                <td style="color: #667085; font-size: 13px; font-weight: 600; padding: 6px 12px 6px 0; vertical-align: top;">Proposed Date</td>
                                                <td style="color: #101828; font-size: 14px; font-weight: 500; padding: 6px 0; vertical-align: top;">${proposedDate}</td>
                                            </tr>
                                            <tr>
                                                <td style="color: #667085; font-size: 13px; font-weight: 600; padding: 6px 12px 6px 0; vertical-align: top;">Duration</td>
                                                <td style="color: #101828; font-size: 14px; font-weight: 500; padding: 6px 0; vertical-align: top;">${duration}</td>
                                            </tr>
                                        </table>

                                        <!-- Estimated Training Fee Card -->
                                        <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #F8FAFC; border: 1px solid #E3E7EC; border-radius: 6px; border-collapse: separate; margin-top: 20px;">
                                            <tr>
                                                <td style="padding: 16px 20px;">
                                                    <div style="color: #667085; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">ESTIMATED TRAINING FEE</div>
                                                    <div style="color: #17365D; font-size: 22px; font-weight: 800;">${estimatedFee}</div>
                                                </td>
                                            </tr>
                                        </table>

                                    </td>
                                </tr>
                            </table>

                            <!-- Current Status Section -->
                            <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: ${statusBg}; border-left: 4px solid ${statusBorder}; border-top: 1px solid #E3E7EC; border-right: 1px solid #E3E7EC; border-bottom: 1px solid #E3E7EC; border-radius: 0 6px 6px 0; border-collapse: separate; margin-bottom: 28px;">
                                <tr>
                                    <td style="padding: 14px 18px;">
                                        <div style="color: ${statusText}; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">CURRENT STATUS</div>
                                        <div style="color: ${statusText}; font-size: 16px; font-weight: 700;">${status}</div>
                                    </td>
                                </tr>
                            </table>

                            <!-- Review Button (if reviewUrl provided) -->
                            ${reviewUrl ? `
                            <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="margin: 0 auto 16px auto;">
                                <tr>
                                    <td align="center" style="border-radius: 6px; background-color: #17365D;">
                                        <a href="${reviewUrl}" target="_blank" style="display: inline-block; padding: 14px 32px; font-family: Arial, sans-serif; font-size: 14px; font-weight: 700; color: #FFFFFF; text-decoration: none; border-radius: 6px; text-transform: uppercase; letter-spacing: 0.5px;">REVIEW TRAINING REQUEST</a>
                                    </td>
                                </tr>
                            </table>
                            ` : ''}

                            <!-- Instruction Text -->
                            <p style="color: #667085; font-size: 12px; text-align: center; margin: 12px 0 0 0; line-height: 1.5;">
                                Review the request details and take the appropriate action through the TrainHub Training Management System.
                            </p>

                        </td>
                    </tr>

                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #F8FAFC; border-top: 1px solid #E3E7EC; padding: 24px 32px; text-align: center;">
                            <div style="font-size: 13px; font-weight: 700; color: #475467; margin-bottom: 6px;">TrainHub Training Management System</div>
                            <div style="font-size: 12px; color: #98A2B3; line-height: 1.5;">
                                This is an automated notification from the Training Management System.<br>
                                Please do not reply directly to this email.
                            </div>
                        </td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
</body>
</html>`;
}

// ─── Targeted Event-Based Incremental Synchronization Engine ──────────────────

const REPORT_CONFIG = {
  'Training Hours': {
    sheetName: 'Training Hours',
    headerRowIndex: 2,
    dataStartRow: 3,
    keyColumnIndex: 1,
    keyFieldName: 'costCentre',
    columns: [
      { name: 'Cost Centre/Month', field: 'costCentre', isKey: true },
      { name: 'Jan', field: 'm0' },
      { name: 'Feb', field: 'm1' },
      { name: 'Mar', field: 'm2' },
      { name: 'Apr', field: 'm3' },
      { name: 'May', field: 'm4' },
      { name: 'Jun', field: 'm5' },
      { name: 'Jul', field: 'm6' },
      { name: 'Aug', field: 'm7' },
      { name: 'Sep', field: 'm8' },
      { name: 'Oct', field: 'm9' },
      { name: 'Nov', field: 'm10' },
      { name: 'Dec', field: 'm11' },
      { name: 'Total Year', field: 'totalYear' },
      { name: 'Total Training Hours', field: 'totalHours' }
    ]
  },
  'Training Cost': {
    sheetName: 'Training Cost',
    headerRowIndex: 2,
    dataStartRow: 3,
    keyColumnIndex: 1,
    keyFieldName: 'trainingTitle',
    columns: [
      { name: 'Training Title', field: 'trainingTitle', isKey: true },
      { name: 'Training Date (From)', field: 'dateFrom' },
      { name: 'Training Date (To)', field: 'dateTo' },
      { name: 'Total Participant', field: 'totalParticipant' },
      { name: 'Training Fees', field: 'trainingFees' },
      { name: 'Meal', field: 'meal' },
      { name: 'Subsistance Allowance', field: 'subsistanceAllowance' },
      { name: 'Hotel Fees', field: 'hotelFees' },
      { name: 'Mileage Claim', field: 'mileageClaim' },
      { name: 'Taxi Fees', field: 'taxiFees' },
      { name: 'Toll Fees', field: 'tollFees' },
      { name: 'Flight', field: 'flight' },
      { name: 'Total Cost', field: 'totalCost' },
      { name: 'Total HRDF Grant (RM)', field: 'totalHrdfGrant' }
    ]
  },
  'Training Title': {
    sheetName: 'Training Title',
    headerRowIndex: 1,
    dataStartRow: 2,
    keyColumnIndex: 2,
    keyFieldName: 'empNoKey',
    columns: [
      { name: 'Name', field: 'name' },
      { name: 'Emp. No', field: 'empNo', isKey: true },
      { name: 'Training Title', field: 'trainingTitle' },
      { name: '(EE)/Cost Centre Description', field: 'costCentreDesc' },
      { name: 'Training Type', field: 'trainingType' },
      { name: 'Date (From)', field: 'dateFrom' },
      { name: 'Date (Until)', field: 'dateUntil' },
      { name: 'Month', field: 'month' },
      { name: 'Total Hours', field: 'totalHours' },
      { name: 'Trainer', field: 'trainer' },
      { name: 'Training Provide', field: 'trainingProvider' },
      { name: 'Expiry Date (if applicable)', field: 'expiryDate' }
    ]
  },
  'Employee Report': {
    sheetName: 'Employee Report',
    headerRowIndex: 1,
    dataStartRow: 2,
    keyColumnIndex: 2,
    keyFieldName: 'empNo',
    columns: [
      { name: 'No', field: 'no' },
      { name: 'Emp. No', field: 'empNo', isKey: true },
      { name: 'Employee Name', field: 'name' },
      { name: 'Cost Centre / Department', field: 'costCentre' },
      { name: 'Position', field: 'position' },
      { name: 'Total Trainings', field: 'totalTrainings' },
      { name: 'Total Hours', field: 'totalHours' },
      { name: 'Attended Trainings', field: 'attendedList' },
      { name: 'Status', field: 'status' }
    ]
  },
  'Annual Training Plan (ATP)': {
    sheetName: 'Annual Training Plan (ATP)',
    headerRowIndex: 1,
    dataStartRow: 2,
    keyColumnIndex: 2,
    keyFieldName: 'trainingTitle',
    columns: [
      { name: 'No', field: 'no' },
      { name: 'Training Title', field: 'trainingTitle', isKey: true },
      { name: 'Training Category', field: 'trainingCategory' },
      { name: 'TNA Source', field: 'tnaSource' },
      { name: 'Training Mode', field: 'trainingMode' },
      { name: 'Training Duration (hrs)', field: 'durationHours' },
      { name: 'Trainer', field: 'trainer' },
      { name: 'Department', field: 'department' },
      { name: 'Position / Employee No', field: 'positionFormula' },
      { name: 'Total Pax', field: 'totalPax' },
      { name: 'Planned Date', field: 'plannedDate' },
      { name: 'Actual Date (From)', field: 'actualDateFrom' },
      { name: 'Actual Date (To)', field: 'actualDateTo' },
      { name: 'Training Status', field: 'trainingStatus' },
      { name: 'Remarks', field: 'remarks' }
    ]
  }
};

function generateSyncId() {
  return 'SYNC-' + String(Date.now()).slice(-6);
}

function getOrCreateReportsFolder() {
  const rootFolderId = getConfigProperty('ROOT_FOLDER_ID', '');
  if (!rootFolderId) throw new Error('ROOT_FOLDER_ID is required.');
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  let folderIter = rootFolder.getFoldersByName('Reports');
  if (folderIter.hasNext()) return folderIter.next();
  throw new Error("Required folder 'Reports' was not found under ROOT_FOLDER_ID.");
}

function getOrCreateSyncHistorySheet(ss) {
  let historySheet = ss.getSheetByName('_SYNC_HISTORY');
  if (!historySheet) {
    historySheet = ss.insertSheet('_SYNC_HISTORY');
    historySheet.appendRow(['Sync ID', 'Timestamp', 'Training ID', 'Action', 'Trigger', 'Reports', 'Status', 'Duration', 'Message', 'Error']);
    historySheet.getRange('A1:J1').setFontWeight('bold').setBackground('#1E293B').setFontColor('#FFFFFF');
    try { historySheet.setFrozenRows(1); } catch (e) {}
  }
  return historySheet;
}

function logSyncHistory(ss, rec) {
  try {
    if (!ss) {
      const repFolder = getOrCreateReportsFolder();
      const fileName = `Master Annual Training Report (2026)`;
      let fileIter = repFolder.getFilesByName(fileName);
      if (fileIter.hasNext()) ss = SpreadsheetApp.openById(fileIter.next().getId());
    }
    if (!ss) return;

    const historySheet = getOrCreateSyncHistorySheet(ss);
    const syncId = rec.syncId || generateSyncId();
    const timeStamp = rec.timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    historySheet.appendRow([
      syncId,
      timeStamp,
      rec.trainingId || '',
      rec.action || 'UPDATE',
      rec.trigger || 'System Action',
      rec.reports || '',
      rec.status || 'SUCCESS',
      rec.duration || '0.00s',
      rec.message || '',
      rec.error || ''
    ]);
  } catch (e) {
    Logger.log('logSyncHistory error: ' + e.message);
  }
}

function getOrCreateSyncMetadataSheet(ss) {
  let metaSheet = ss.getSheetByName('_SYNC_METADATA');
  if (!metaSheet) {
    metaSheet = ss.insertSheet('_SYNC_METADATA');
    metaSheet.appendRow(['Report Sheet', 'Record Key', 'Column Name', 'Last System Value', 'Last Sync Timestamp']);
    metaSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#334155').setFontColor('#FFFFFF');
    try { metaSheet.hideSheet(); } catch (e) {}
  } else {
    try { metaSheet.hideSheet(); } catch (e) {}
  }
  return metaSheet;
}

function loadSyncMetadata(ss) {
  const metaSheet = getOrCreateSyncMetadataSheet(ss);
  const data = metaSheet.getDataRange().getValues();
  const map = {};

  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      const sheetName = String(data[i][0] || '').trim();
      const recKey    = String(data[i][1] || '').trim();
      const colName   = String(data[i][2] || '').trim();
      const lastVal   = data[i][3];

      if (sheetName && recKey && colName) {
        if (!map[sheetName]) map[sheetName] = {};
        if (!map[sheetName][recKey]) map[sheetName][recKey] = {};
        map[sheetName][recKey][colName] = lastVal;
      }
    }
  }
  return map;
}

function saveSyncMetadata(ss, metadataMap) {
  const metaSheet = getOrCreateSyncMetadataSheet(ss);
  metaSheet.clearContents();
  metaSheet.appendRow(['Report Sheet', 'Record Key', 'Column Name', 'Last System Value', 'Last Sync Timestamp']);

  const timeStamp = new Date().toISOString();
  const rows = [];

  Object.keys(metadataMap).forEach(sheetName => {
    Object.keys(metadataMap[sheetName]).forEach(recKey => {
      Object.keys(metadataMap[sheetName][recKey]).forEach(colName => {
        const val = metadataMap[sheetName][recKey][colName];
        rows.push([sheetName, recKey, colName, val, timeStamp]);
      });
    });
  });

  if (rows.length > 0) {
    metaSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  metaSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#334155').setFontColor('#FFFFFF');
  try { metaSheet.hideSheet(); } catch (e) {}
}

function normalizeSyncValue(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return val.toISOString().slice(0, 10);
  }
  let str = String(val).trim();
  if (str === '—' || str === 'N/A' || str === 'null' || str === 'undefined') {
    return '';
  }
  return str;
}

function isCellAdminModified(currentSheetValue, lastSystemValue) {
  const normCurrent = normalizeSyncValue(currentSheetValue);
  const normLast = normalizeSyncValue(lastSystemValue);
  return normCurrent !== normLast;
}

function ensureReportHeaderExists(sheet, reportKey, config, extraParams) {
  if (sheet.getLastRow() > 0) return;
  const year = (extraParams && extraParams.year) ? extraParams.year : '2026';

  if (reportKey === 'Training Hours') {
    const yrSuffix = year && year.length === 4 ? year.slice(-2) : '26';
    const monthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].map(m => `${m}-${yrSuffix}`);
    const row1 = ['Cost Centre/Month', 'Training Hours', '', '', '', '', '', '', '', '', '', '', '', `Total ${year}`, `Total Training Hours ${year}`];
    const row2 = ['', ...monthLabels, '', ''];
    sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
    sheet.getRange(2, 1, 1, row2.length).setValues([row2]);
    sheet.getRange('A1:A2').merge().setValue('Cost Centre/Month');
    sheet.getRange('B1:M1').merge().setValue('Training Hours');
    sheet.getRange('N1:N2').merge().setValue(`Total ${year}`);
    sheet.getRange('O1:O2').merge().setValue(`Total Training Hours ${year}`);
    sheet.getRange('A1:O2').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('center');
  } else if (reportKey === 'Training Cost') {
    const row1 = ['Training Title', 'Training Date (From)', 'Training Date (To)', 'Total Participant', 'Training Cost (RM)', '', '', '', '', '', '', '', `Total Cost ${year}`, 'Total HRDF Grant (RM)'];
    const row2 = ['', '', '', '', 'Training Fees', 'Meal', 'Subsistance Allowance', 'Hotel Fees', 'Mileage Claim', 'Taxi Fees', 'Toll Fees', 'Flight', '', ''];
    sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
    sheet.getRange(2, 1, 1, row2.length).setValues([row2]);
    sheet.getRange('A1:A2').merge().setValue('Training Title');
    sheet.getRange('B1:B2').merge().setValue('Training Date (From)');
    sheet.getRange('C1:C2').merge().setValue('Training Date (To)');
    sheet.getRange('D1:D2').merge().setValue('Total Participant');
    sheet.getRange('E1:L1').merge().setValue('Training Cost (RM)');
    sheet.getRange('M1:M2').merge().setValue(`Total Cost ${year}`);
    sheet.getRange('N1:N2').merge().setValue('Total HRDF Grant (RM)');
    sheet.getRange('A1:N2').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('center');
  } else {
    const headerCols = config.columns.map(c => c.name);
    sheet.getRange(1, 1, 1, headerCols.length).setValues([headerCols]);
    sheet.getRange(1, 1, 1, headerCols.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  }
}

function getRecordValueForColumn(record, colDef, rowNumber) {
  const field = colDef.field;
  if (!field) return '';
  if (field.startsWith('m') && field.length <= 3 && !isNaN(parseInt(field.slice(1)))) {
    const idx = parseInt(field.slice(1), 10);
    const months = Array.isArray(record.months) ? record.months : [];
    return months[idx] !== undefined ? months[idx] : 0;
  }
  if (field === 'positionFormula') {
    const reqUrl = record.requisitionFormUrl || '#';
    return (reqUrl && reqUrl !== '#') ? `=HYPERLINK("${reqUrl}", "Name List")` : 'Name List';
  }
  if (field === 'totalCost' || field === 'totalCostFormula') {
    const rNum = rowNumber || 3;
    return `=SUM(E${rNum}:L${rNum})`;
  }
  if (['trainingFees', 'meal', 'subsistanceAllowance', 'hotelFees', 'mileageClaim', 'taxiFees', 'tollFees', 'flight', 'totalHrdfGrant'].includes(field)) {
    const val = record[field];
    if (val === 0 || val === '0' || val === 'RM 0.00' || val === 'RM 0' || val === null || val === undefined) {
      return '';
    }
    return val;
  }
  if (field === 'attendedList' && Array.isArray(record.attendedList)) {
    return record.attendedList.join(', ');
  }
  return record[field] !== undefined ? record[field] : '';
}

function syncReportSheetIncrementally(ss, reportKey, reportData, extraParams) {
  const config = REPORT_CONFIG[reportKey];
  if (!config) return;

  let sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) sheet = ss.insertSheet(config.sheetName);

  ensureReportHeaderExists(sheet, reportKey, config, extraParams);

  const metadataMap = loadSyncMetadata(ss);
  if (!metadataMap[config.sheetName]) metadataMap[config.sheetName] = {};
  const sheetMeta = metadataMap[config.sheetName];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  let existingSheetValues = [];
  if (lastRow >= config.dataStartRow && lastCol >= 1) {
    existingSheetValues = sheet.getRange(config.dataStartRow, 1, lastRow - config.dataStartRow + 1, Math.max(lastCol, config.columns.length)).getValues();
  }

  const rowMap = {};
  existingSheetValues.forEach((rowVals, idx) => {
    const sheetRowNumber = config.dataStartRow + idx;
    const rawKeyVal = rowVals[config.keyColumnIndex - 1];
    const keyVal = normalizeSyncValue(rawKeyVal);

    if (keyVal && keyVal.toLowerCase() !== 'total' && keyVal.toLowerCase() !== 'summary') {
      rowMap[keyVal.toLowerCase()] = {
        rowNumber: sheetRowNumber,
        rowValues: rowVals
      };
    }
  });

  const incomingRows = reportData.rows || [];

  incomingRows.forEach(record => {
    let recKey = '';
    if (config.keyFieldName === 'empNoKey') {
      recKey = normalizeSyncValue(`${record.empNo || record.name || ''}_${record.trainingTitle || ''}`);
    } else {
      recKey = normalizeSyncValue(record[config.keyFieldName] || record.trainingTitle || record.costCentre || record.empNo || record.ID || record.no);
    }
    if (!recKey) return;

    if (!sheetMeta[recKey]) sheetMeta[recKey] = {};
    const recordMeta = sheetMeta[recKey];
    const existingRow = rowMap[recKey.toLowerCase()];

    if (existingRow) {
      const rowNum = existingRow.rowNumber;
      const currentVals = existingRow.rowValues;

      config.columns.forEach((colDef, cIdx) => {
        const colNumber = cIdx + 1;
        const colName = colDef.name;
        const incomingVal = getRecordValueForColumn(record, colDef, rowNum);

        const currentSheetVal = currentVals[cIdx] !== undefined ? currentVals[cIdx] : '';
        let lastSystemVal     = recordMeta[colName];

        if (lastSystemVal === undefined) {
          lastSystemVal = currentSheetVal;
          if (normalizeSyncValue(currentSheetVal) !== '') {
            recordMeta[colName] = currentSheetVal;
          }
        }

        if (!isCellAdminModified(currentSheetVal, lastSystemVal)) {
          if (normalizeSyncValue(currentSheetVal) !== normalizeSyncValue(incomingVal)) {
            sheet.getRange(rowNum, colNumber).setValue(incomingVal);
            recordMeta[colName] = incomingVal;
          }
        }
      });
    } else {
      const targetRowNum = sheet.getLastRow() + 1;
      const newRowVals = config.columns.map(colDef => getRecordValueForColumn(record, colDef, targetRowNum));
      sheet.appendRow(newRowVals);

      config.columns.forEach((colDef, cIdx) => {
        recordMeta[colDef.name] = newRowVals[cIdx];
      });
    }
  });

  saveSyncMetadata(ss, metadataMap);
}

function buildAnnualTrainingPlanData(trainingsInput) {
  const trainings = Array.isArray(trainingsInput) ? trainingsInput : [];
  const rows = trainings.map((t, idx) => {
    const feeNum = parseFloat(String(t.CourseFee || '0').replace(/[^0-9\.-]/g, '')) || 0;
    const paxNum = parseInt(String(t.Participants || 20), 10) || 20;
    let reqUrl = t.RequisitionUrl || t.RequisitionFormUrl || '';
    if (!reqUrl && t.RequisitionFormFileID) {
      reqUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(t.RequisitionFormFileID)}/edit`;
    }

    return {
      no: idx + 1,
      trainingTitle: t.Name || t['Training Name'] || `Training ${idx+1}`,
      trainingCategory: t.Category || 'General',
      tnaSource: t.TnaSource || 'Training Requisition Form',
      trainingMode: t.TrainingMode || t.Mode || 'In-house',
      durationHours: t.TotalHours || 8,
      trainer: t.Trainer || 'Certified Trainer',
      department: t.Department || 'All Departments',
      positionEmpNo: 'Name List',
      requisitionFormUrl: reqUrl,
      totalPax: paxNum,
      plannedDate: t.StartDate || '',
      actualDateFrom: t.StartDate || '',
      actualDateTo: t.EndDate || t.StartDate || '',
      trainingStatus: t.ApprovalStatus || t.Status || 'Pending',
      remarks: t.Stage || 'Planned'
    };
  });
  return { rows: rows };
}

function buildCostReportData(trainingsInput, year) {
  const trainings = Array.isArray(trainingsInput) ? trainingsInput : [];
  const rows = trainings.map(t => {
    const feeNum = parseFloat(String(t.CourseFee || '0').replace(/[^0-9\.-]/g, '')) || 0;
    const paxNum = parseInt(String(t.Participants || 20), 10) || 20;
    return {
      trainingTitle: t.Name || t['Training Name'] || 'Training Programme',
      dateFrom: t.StartDate || '',
      dateTo: t.EndDate || t.StartDate || '',
      totalParticipant: paxNum,
      trainingFees: feeNum > 0 ? feeNum : '',
      meal: '',
      subsistanceAllowance: '',
      hotelFees: '',
      mileageClaim: '',
      taxiFees: '',
      tollFees: '',
      flight: '',
      totalCost: feeNum > 0 ? feeNum : '',
      totalHrdfGrant: ''
    };
  });
  return { year: year, rows: rows };
}

function syncTrainingById(trainingId, triggerName, actionName) {
  const lock = LockService.getScriptLock();
  let hasLock = false;
  try {
    hasLock = lock.tryLock(5000);
  } catch (lErr) {}

  const startTime = Date.now();
  const trigger = triggerName || 'System Trigger';
  const action = actionName || 'UPDATE';
  const displayYear = '2026';

  try {
    if (!trainingId || String(trainingId).trim() === '') {
      return { success: false, message: 'Invalid Training ID.' };
    }

    const tId = String(trainingId).trim();
    const tSheet = getSheet('Trainings');
    if (!tSheet) return { success: false, message: 'Trainings sheet unavailable.' };
    const tRows = sheetToJson(tSheet);
    const tRecord = tRows.find(r => String(r.ID || '').trim() === tId || String(r.Code || '').trim() === tId);

    if (!tRecord) {
      const durationStr = ((Date.now() - startTime) / 1000).toFixed(2) + 's';
      logSyncHistory(null, {
        syncId: generateSyncId(),
        timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
        trainingId: tId,
        action: action,
        trigger: trigger,
        reports: 'None',
        status: 'FAILED',
        duration: durationStr,
        message: `Unable to find Training ID: ${tId}`,
        error: `Training ID ${tId} not found in database.`
      });
      return { success: false, message: `Unable to find Training ID: ${tId}` };
    }

    const repFolder = getOrCreateReportsFolder();
    const fileName = `Master Annual Training Report (${displayYear})`;

    let fileIter = repFolder.getFilesByName(fileName);
    let ss;
    if (fileIter.hasNext()) {
      ss = SpreadsheetApp.openById(fileIter.next().getId());
    } else {
      ss = SpreadsheetApp.create(fileName);
      const file = DriveApp.getFileById(ss.getId());
      repFolder.addFile(file);
      try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
    }

    const atpData = buildAnnualTrainingPlanData([tRecord]);
    const costData = buildCostReportData([tRecord], displayYear);

    const reportsUpdated = [];

    if (atpData && atpData.rows && atpData.rows.length > 0) {
      syncReportSheetIncrementally(ss, 'Annual Training Plan (ATP)', atpData);
      reportsUpdated.push('Annual Training Plan (ATP)');
    }

    if (costData && costData.rows && costData.rows.length > 0) {
      syncReportSheetIncrementally(ss, 'Training Cost', costData, { year: displayYear });
      reportsUpdated.push('Training Cost');
    }

    SpreadsheetApp.flush();

    const durationStr = ((Date.now() - startTime) / 1000).toFixed(2) + 's';
    const syncId = generateSyncId();
    const message = `Successfully synchronized Training ID ${tId} across ${reportsUpdated.length} report tab(s).`;
    const timeStampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    logSyncHistory(ss, {
      syncId: syncId,
      timestamp: timeStampStr,
      trainingId: tId,
      action: action,
      trigger: trigger,
      reports: reportsUpdated.join(', '),
      status: 'SUCCESS',
      duration: durationStr,
      message: message,
      error: ''
    });

    return {
      success: true,
      syncId: syncId,
      trainingId: tId,
      action: action,
      trigger: trigger,
      reports: reportsUpdated,
      status: 'SUCCESS',
      duration: durationStr,
      message: message,
      fileUrl: ss.getUrl()
    };

  } catch (e) {
    const durationStr = ((Date.now() - startTime) / 1000).toFixed(2) + 's';
    Logger.log('syncTrainingById error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    if (hasLock) {
      try { lock.releaseLock(); } catch(rErr) {}
    }
  }
}
