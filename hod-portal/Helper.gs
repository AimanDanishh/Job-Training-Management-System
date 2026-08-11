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

/**
 * Reusable employee lookup function by Employee No / ID.
 * Searches the existing Employees sheet in EMPLOYEE_SPREADSHEET_ID.
 * Matches Employee No exactly (case-insensitive, trimmed).
 * Returns employee details object or null if not found.
 */
function getEmployeeById(employeeNo) {
  if (!employeeNo || String(employeeNo).trim() === '') return null;
  const cleanId = String(employeeNo).trim().toLowerCase();

  const empSheet = getSheet('Employees');
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
    const sheet = ss.getSheetByName('Training Form') || ss.getSheets()[0];

    // Ensure template sign headers in Row 40 remain intact
    sheet.getRange('A40').setValue('REQUEST BY');
    sheet.getRange('C40').setValue('VERIFIED BY HEAD OF DEPARTMENT');
    sheet.getRange('D40').setValue('APPROVED BY C-SUITE');
    sheet.getRange('F40').setValue('APPROVED BY HOHR');
    sheet.getRange('H40').setValue('ACKNOWLEDGED BY HR DEPARTMENT');

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

    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log('updateTrainingRequisitionSignatures error: ' + e.message);
  }
}
