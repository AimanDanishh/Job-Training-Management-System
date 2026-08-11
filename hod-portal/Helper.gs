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
