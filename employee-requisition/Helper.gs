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

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function generateId(prefix) {
  const p = prefix || 'TRN';
  const num = Math.floor(1000 + Math.random() * 9000);
  return `${p}-${num}`;
}

/**
 * Builds the authoritative participant directory from EMPLOYEE_SPREADSHEET_ID.
 * Training participant rows are copies only; they must never be treated as the
 * source of an employee number, name, cost centre, or job title.
 */
function getOfficialEmployeeDirectory() {
  const byId = {};
  const byName = {};
  const empSheet = getSheet('Employees');
  const rows = empSheet ? sheetToJson(empSheet) : [];

  rows.forEach(row => {
    const id = String(row.ID || row.EmployeeID || row.EmployeeNo || row.EmpID || row.StaffID || row['Employee ID'] || row['Employee No'] || row['Staff ID'] || '').trim();
    const name = String(row.Name || row.EmployeeName || row['Employee Name'] || row['Staff Name'] || '').trim();
    if (!id) return;

    const employee = {
      ID: id,
      Name: name,
      Department: String(row.CostCentre || row['Cost Centre'] || row.Department || row.Dept || '').trim(),
      Position: String(row.PositionTitle || row['Position Title'] || row.Position || row.JobTitle || row['Job Title'] || '').trim(),
      Email: String(row.Email || row['Email Address'] || '').trim()
    };
    byId[id.toLowerCase()] = employee;
    if (name) {
      const nameKey = name.toLowerCase();
      // Do not use a name that belongs to more than one employee.
      byName[nameKey] = Object.prototype.hasOwnProperty.call(byName, nameKey) ? null : employee;
    }
  });
  return { byId: byId, byName: byName };
}

function canonicalizeTrainingParticipants(participants) {
  const directory = getOfficialEmployeeDirectory();
  const canonical = [];
  const rejected = [];
  const seen = {};

  (Array.isArray(participants) ? participants : []).forEach(participant => {
    let rawId = String(participant.EmployeeID || participant.EmployeeNo || participant.EmpID || participant['Employee ID'] || participant['Employee No'] || '').trim();
    if (!rawId && participant.ID && !String(participant.ID).startsWith('TP-')) {
      rawId = String(participant.ID).trim();
    }
    const rawName = String(participant.EmployeeName || participant.Name || participant['Employee Name'] || '').trim();
    const employee = (rawId && directory.byId[rawId.toLowerCase()]) || (rawName && directory.byName[rawName.toLowerCase()]);

    if (!employee) {
      rejected.push(rawId || rawName || 'blank participant');
      return;
    }
    const idKey = employee.ID.toLowerCase();
    if (!seen[idKey]) {
      seen[idKey] = true;
      canonical.push(employee);
    }
  });
  return { participants: canonical, rejected: rejected };
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

/** Utility: Clean & Extract raw Google Drive / Sheet ID from URL or ID string */
function extractCleanDriveId(val) {
  if (!val) return '';
  const str = String(val).trim();
  const match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

/** Google Drive Workspace & Requisition Form Helpers */
function getSystemRootFolder() {
  const rootFolderId = extractCleanDriveId(getConfigProperty('ROOT_FOLDER_ID', ''));
  if (!rootFolderId) {
    throw new Error('ROOT_FOLDER_ID is required. Configure the system root folder before submitting a requisition.');
  }

  const systemRoot = DriveApp.getFolderById(rootFolderId);
  const trainingFolders = systemRoot.getFoldersByName('Training Folder');
  if (!trainingFolders.hasNext()) {
    throw new Error("Required folder 'Training Folder' was not found under ROOT_FOLDER_ID.");
  }
  return trainingFolders.next();
}

function createTrainingWorkspace(code, trainingName) {
  try {
    const parentFolder = getSystemRootFolder();
    const folderName = `${code} ${trainingName}`.substring(0, 100).trim();
    let folderIter = parentFolder.getFoldersByName(folderName);
    let folder = folderIter.hasNext() ? folderIter.next() : parentFolder.createFolder(folderName);

    let singleSheetFile = getOrCreateSingleTrainingSheet(folder, code);
    const singleSheetId = singleSheetFile.getId();

    return {
      folderId:            folder.getId(),
      folderUrl:           folder.getUrl(),
      partSheetId:         singleSheetId,
      sessionSheetId:      singleSheetId,
      attendanceSheetId:   singleSheetId,
      evaluationSheetId:   singleSheetId,
      postSheetId:         singleSheetId
    };
  } catch(e) {
    Logger.log('createTrainingWorkspace error: ' + e.message);
    const rootFolder = getSystemRootFolder();
    return {
      folderId: rootFolder ? rootFolder.getId() : '', folderUrl: rootFolder ? rootFolder.getUrl() : '',
      partSheetId: '', sessionSheetId: '', attendanceSheetId: '', evaluationSheetId: '', postSheetId: ''
    };
  }
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
      // SpreadsheetApp creates the file in My Drive first; move it into the
      // training workspace so no data file is left at a random Drive location.
      file.moveTo(folder);
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

/** Opens the one Training Data spreadsheet for a training requisition. */
function getTrainingDataSpreadsheet(trainingId) {
  if (!trainingId) return null;
  const trainingSheet = getSheet('Trainings');
  if (!trainingSheet) return null;
  const rows = sheetToJson(trainingSheet);
  const training = rows.find(r => String(r.ID || r.TrainingID || r.Code || '').trim() === String(trainingId).trim());
  if (!training) return null;

  try {
    const trainingRoot = getSystemRootFolder();
    const code = training.Code || training.ID;
    const folders = trainingRoot.getFoldersByName(`${code} ${training.Name || ''}`.trim());
    if (!folders.hasNext()) return null;
    const folder = folders.next();
    const fileIter = folder.getFilesByName(`${code} Training Data`);
    return fileIter.hasNext() ? SpreadsheetApp.openById(fileIter.next().getId()) : null;
  } catch (e) {
    Logger.log('Could not open Training Data sheet from folder: ' + e.message);
    return null;
  }
}

function syncParticipantsToTrainingDriveSheet(trainingId, directParticipantsList) {
  try {
    const tSheet = getSheet('Trainings');
    if (!tSheet) return;
    const headers = tSheet.getDataRange().getValues()[0].map(h => String(h).trim());
    let row = -1;
    const data = tSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(trainingId).trim()) { row = i + 1; break; }
    }
    if (row === -1) return;

    const tData = data[row - 1];
    const folderId = tData[15]; // Column P = FolderID
    const code     = tData[1] || trainingId;

    if (!folderId) return;

    let participantsToSync = [];

    if (Array.isArray(directParticipantsList) && directParticipantsList.length > 0) {
      participantsToSync = directParticipantsList.map(p => ({
        ID: p.ID || p.EmployeeID || generateId('TP'),
        TrainingID: trainingId,
        EmployeeID: p.EmployeeID || p.ID || p.EmployeeNo || '',
        EmployeeName: p.EmployeeName || p.Name || '',
        Department: p.Department || p.CostCentre || '',
        Position: p.Position || p.JobTitle || p.PositionTitle || '',
        AddedAt: p.AddedAt || now()
      }));
    }

    // Always rebuild Training Data from the Employee-sheet source of truth.
    const resolution = canonicalizeTrainingParticipants(participantsToSync);
    if (resolution.rejected.length > 0) {
      Logger.log('Skipped Training Data participant rows not found in Employees for ' + trainingId + ': ' + resolution.rejected.join(', '));
    }
    participantsToSync = resolution.participants.map(p => ({
      ID: generateId('TP'), TrainingID: trainingId, EmployeeID: p.ID,
      EmployeeName: p.Name, Department: p.Department, Position: p.Position, AddedAt: now()
    }));

    const folder = DriveApp.getFolderById(folderId);
    const headersList = ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt'];
    const file = getOrCreateSingleTrainingSheet(folder, code);

    const participantSheetIdColumn = headers.indexOf('ParticipantsSheetID') + 1;
    if (participantSheetIdColumn > 0) tSheet.getRange(row, participantSheetIdColumn).setValue(file.getId());

    const ss = SpreadsheetApp.openById(file.getId());
    let sheet = ss.getSheetByName('TrainingParticipants') || ss.getActiveSheet();

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, headersList.length).clearContent();
    }

    if (participantsToSync.length > 0) {
      const dataToAppend = participantsToSync.map(p => [
        p.ID || generateId('TP'),
        trainingId,
        p.EmployeeID || p.ID || '',
        p.EmployeeName || p.Name || '',
        p.Department || p.CostCentre || '',
        p.Position || p.JobTitle || '',
        p.AddedAt || now()
      ]);
      sheet.getRange(2, 1, dataToAppend.length, headersList.length).setValues(dataToAppend);
    }
    SpreadsheetApp.flush();

    // Clear memory cache so subsequent reads see latest data
    _sheetDataCache = {};
  } catch (e) {
    Logger.log('syncParticipantsToTrainingDriveSheet error: ' + e.message);
  }
}

function syncAllTrainingDataSheets() {
  try {
    const tSheet = getSheet('Trainings');
    if (!tSheet || tSheet.getLastRow() <= 1) return ok('No trainings to sync.');
    const rows = sheetToJson(tSheet);
    let count = 0;
    rows.forEach(t => {
      if (t.ID && t.Code && t.Name) {
        const workspace = createTrainingWorkspace(t.Code, t.Name);
        if (t._row) {
          const headers = ensureTrainingSheetColumns(tSheet);
          const setColVal = (colName, val) => {
            const idx = headers.indexOf(colName) + 1;
            if (idx > 0 && val) tSheet.getRange(t._row, idx).setValue(val);
          };
          setColVal('FolderID', workspace.folderId);
          setColVal('ParticipantsSheetID', workspace.partSheetId);
          setColVal('SessionsSheetID', workspace.sessionSheetId);
          setColVal('AttendanceSheetID', workspace.attendanceSheetId);
          setColVal('EvaluationSheetID', workspace.evaluationSheetId);
          setColVal('PostSheetID', workspace.postSheetId);
        }
        syncParticipantsToTrainingDriveSheet(t.ID);
        syncTrainingRequisitionParticipants(t.ID);
        count++;
      }
    });
    SpreadsheetApp.flush();
    return ok(`Successfully synced per-training Drive workspaces and Training Data sheets for ${count} trainings.`);
  } catch (e) {
    Logger.log('syncAllTrainingDataSheets error: ' + e.message);
    return err('Failed to sync training data sheets: ' + e.message);
  }
}

function createTrainingRequisitionForm(code, training, targetFolderId, requesterSigData) {
  let rawFolderId = extractCleanDriveId(targetFolderId);
  let targetFolder = null;

  if (rawFolderId) {
    try { targetFolder = DriveApp.getFolderById(rawFolderId); } catch(fErr) {
      Logger.log('Error opening targetFolderId (' + rawFolderId + '): ' + fErr.message);
    }
  }
  if (!targetFolder) {
    return { message: 'Requisition form skipped: the dedicated training workspace is unavailable.' };
  }

  const rawTemplateId = extractCleanDriveId(getConfigProperty('TRAINING_REQUISITION_TEMPLATE_ID', ''));
  const fileName = `${code} Training Requisition Form`;
  let copiedFile = null;

  if (rawTemplateId) {
    try {
      const templateFile = DriveApp.getFileById(rawTemplateId);
      copiedFile = templateFile.makeCopy(fileName, targetFolder);
      Logger.log(`Successfully copied template ID (${rawTemplateId}) into target folder`);
    } catch (tmplErr) {
      Logger.log('Template copy error from ID (' + rawTemplateId + '): ' + tmplErr.message);
    }
  }

  // Fallback: If no template or copy failed, create a new Google Sheet directly in targetFolder
  if (!copiedFile) {
    try {
      const newSs = SpreadsheetApp.create(fileName);
      copiedFile = DriveApp.getFileById(newSs.getId());
      copiedFile.moveTo(targetFolder);
      Logger.log(`Created new clean form Google Sheet in target folder`);
    } catch(createErr) {
      Logger.log('SpreadsheetApp.create error: ' + createErr.message);
    }
  }

  if (!copiedFile) {
    return { message: 'Requisition form creation failed completely.' };
  }

  try {
    copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(pErr) {}

  const ss = SpreadsheetApp.openById(copiedFile.getId());
  let sheet = ss.getSheetByName('Training Form') || ss.getSheets()[0];

  const tData = training || {};
  const startDate = formatDate(tData.StartDate);
  const endDate = formatDate(tData.EndDate);
  const duration = Number(tData.Duration) || 1;
  const hours = tData.TotalHours ? ` (${tData.TotalHours} hours)` : '';

  try {
    const setTemplateValue = (area, value) => sheet.getRange(area.split(':')[0]).setValue(value == null ? '' : value);
    sheet.getRange('A5').setValue('Training Title:');
    setTemplateValue('C5:I5', tData.Name || tData.TrainingName || '');
    sheet.getRange('A6').setValue('Course Fee (RM):');
    setTemplateValue('C6:E6', tData.CourseFee !== undefined ? tData.CourseFee : '0.00');
    sheet.getRange('F6').setValue('Date:');
    setTemplateValue('G6:I6', endDate && endDate !== startDate ? `${startDate} - ${endDate}` : (startDate || ''));
    sheet.getRange('A7').setValue('Duration:');
    setTemplateValue('C7:E7', `${duration} day(s)${hours}`);
    sheet.getRange('F7').setValue('Venue:');
    setTemplateValue('G7:I7', tData.Venue || '');
    sheet.getRange('A8').setValue('Training Provider:');
    setTemplateValue('C8:I8', tData.TrainingProvider || tData.Provider || tData.Trainer || '');
    sheet.getRange('A10').setValue('Reasons for Training:');
    setTemplateValue('A11:I12', tData.Objectives || tData.Reason || '');

    const partList = Array.isArray(tData.ParticipantList) ? tData.ParticipantList : (Array.isArray(tData.participants) ? tData.participants : []);
    if (partList.length > 0) {
      try {
        sheet.getRange('A15:I38').clearContent();
        partList.slice(0, 24).forEach((p, index) => {
          const r = 15 + index;
          setTemplateValue(`A${r}:B${r}`, p.EmployeeID || p.ID || p.EmployeeNo || '');
          setTemplateValue(`C${r}`, p.EmployeeName || p.Name || '');
          setTemplateValue(`D${r}:G${r}`, p.Department || p.CostCentre || '');
          setTemplateValue(`H${r}:I${r}`, p.Position || p.JobTitle || '');
        });
      } catch(pErr) {
        Logger.log('Error populating participants in form creation: ' + pErr.message);
      }
    }

    // Preserve original template header rows 39 and 40 (columns A to I) completely untouched

    if (requesterSigData) {
      const empNo    = requesterSigData.employeeNo || requesterSigData.EmployeeNo || requesterSigData.EmployeeID || requesterSigData.ID || '';
      const empName  = requesterSigData.name || requesterSigData.EmployeeName || requesterSigData.Name || '';
      const position = requesterSigData.position || requesterSigData.JobPosition || requesterSigData.Position || requesterSigData.JobTitle || 'Requester';
      const sigDate  = getFormattedCurrentDate(requesterSigData.date || requesterSigData.Date);

      const formatPrefixed = (prefix, val) => val ? (String(val).toUpperCase().startsWith(prefix.toUpperCase()) ? String(val) : `${prefix} ${String(val).trim()}`) : prefix;

      sheet.getRange('A41').setValue(formatPrefixed('EMPLOYEE NO:', empNo));
      sheet.getRange('A42').setValue(formatPrefixed('NAME:', empName));
      sheet.getRange('A44').setValue(formatPrefixed('JOB POSITION:', position));
      sheet.getRange('A46').setValue(formatPrefixed('DATE:', sigDate));

      sheet.getRange('B41').setValue(empNo);
      sheet.getRange('B42').setValue(empName);
      sheet.getRange('B43').setValue(empName);
      sheet.getRange('B44').setValue(position);
      sheet.getRange('B46').setValue(sigDate);
    }

    SpreadsheetApp.flush();
    return { fileId: copiedFile.getId(), fileUrl: copiedFile.getUrl(), fileName: fileName };
  } catch(e) {
    Logger.log('createTrainingRequisitionForm error: ' + e.message);
    return { fileId: copiedFile ? copiedFile.getId() : '', fileUrl: copiedFile ? copiedFile.getUrl() : '', fileName: fileName };
  }
}

function updateTrainingRequisitionSignatures(trainingId, step, sigData, targetFormId) {
  try {
    SpreadsheetApp.flush();
    const trainingSheet = getSheet('Trainings');
    if (!trainingSheet) return;
    const headers = ensureTrainingSheetColumns(trainingSheet);
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

    // Do not modify or delete template rows 39-40 (columns A to I)

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

    // Auto-check: if a later approval arrives first, populate the requester value cells too.
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

function syncTrainingRequisitionParticipants(trainingId) {
  try {
    const trainingSheet = getSheet('Trainings');
    if (!trainingSheet) return;
    const headers = trainingSheet.getDataRange().getValues()[0].map(h => String(h).trim());
    let row = -1;
    const data = trainingSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(trainingId).trim()) { row = i + 1; break; }
    }
    if (row === -1) return;

    const formCol = headers.indexOf('RequisitionFormFileID') + 1;
    const formId = formCol ? trainingSheet.getRange(row, formCol).getValue() : '';
    if (!formId) return;

    // The per-training Training Data sheet is the sole participant source.
    const trainingData = getTrainingDataSpreadsheet(trainingId);
    const dataParticipantSheet = trainingData ? trainingData.getSheetByName('TrainingParticipants') : null;
    const tpRows = dataParticipantSheet ? sheetToJson(dataParticipantSheet) : [];
    const resolved = canonicalizeTrainingParticipants(tpRows);
    if (resolved.rejected.length > 0) {
      Logger.log('Skipped participant rows not found in Employees for ' + trainingId + ': ' + resolved.rejected.join(', '));
    }

    const formSpreadsheet = SpreadsheetApp.openById(formId);
    const sheet = formSpreadsheet.getSheetByName('Training Form') || formSpreadsheet.getSheets()[0];
    const rowCount = 24;
    sheet.getRange('A15:I38').clearContent();

    resolved.participants.slice(0, rowCount).forEach((p, index) => {
      const r = 15 + index;
      sheet.getRange(`A${r}`).setValue(p.EmployeeID || p.ID || '');
      sheet.getRange(`C${r}`).setValue(p.EmployeeName || p.Name || '');
      sheet.getRange(`D${r}`).setValue(p.Department || p.CostCentre || '');
      sheet.getRange(`H${r}`).setValue(p.Position || p.JobTitle || '');
    });

    const partCountCol = headers.indexOf('Participants') + 1;
    if (partCountCol > 0) trainingSheet.getRange(row, partCountCol).setValue(resolved.participants.length);
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log('syncTrainingRequisitionParticipants error: ' + e.message);
  }
}

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

function resetRequisitionFormApprovals(formId) {
  if (!formId) return;
  try {
    const ss = SpreadsheetApp.openById(formId);
    const sheet = ss.getSheetByName('Training Form') || ss.getSheets()[0];

    // Preserve original template header rows 39 and 40 (columns A to I) completely untouched

    // Reset HOD (Col C), C-Suite (Col D/E), HOHR (Col F/G) signature cells
    sheet.getRange('C41:C46').clearContent();
    sheet.getRange('D41:E46').clearContent();
    sheet.getRange('F41:G46').clearContent();

    sheet.getRange('C41').setValue('STATUS:');
    sheet.getRange('C42').setValue('EMPLOYEE NO:');
    sheet.getRange('C43').setValue('NAME:');
    sheet.getRange('C44').setValue('JOB POSITION:');
    sheet.getRange('C46').setValue('DATE:');

    sheet.getRange('D41').setValue('STATUS:');
    sheet.getRange('D42').setValue('EMPLOYEE NO:');
    sheet.getRange('D43').setValue('NAME:');
    sheet.getRange('D44').setValue('JOB POSITION:');
    sheet.getRange('D46').setValue('DATE:');

    sheet.getRange('F41').setValue('STATUS:');
    sheet.getRange('F42').setValue('EMPLOYEE NO:');
    sheet.getRange('F43').setValue('NAME:');
    sheet.getRange('F44').setValue('JOB POSITION:');
    sheet.getRange('F46').setValue('DATE:');
  } catch(e) {
    Logger.log('resetRequisitionFormApprovals error: ' + e.message);
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
  const rootFolderId = extractCleanDriveId(getConfigProperty('ROOT_FOLDER_ID', ''));
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
