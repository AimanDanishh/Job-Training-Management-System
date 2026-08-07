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
    const rawId = String(participant.ID || participant.EmployeeID || participant.EmployeeNo || participant.EmpID || participant['Employee ID'] || participant['Employee No'] || '').trim();
    const rawName = String(participant.Name || participant.EmployeeName || participant['Employee Name'] || '').trim();
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
  const rootFolderId = extractCleanDriveId(getConfigProperty('TRAINING_FOLDER_ID', '') || getConfigProperty('ROOT_FOLDER_ID', '') || getConfigProperty('DRIVE_ROOT_FOLDER_ID', ''));
  if (rootFolderId) {
    try { return DriveApp.getFolderById(rootFolderId); } catch(e) {
      Logger.log('Could not open configured folder (' + rootFolderId + '): ' + e.message);
    }
  }
  let sysFolderIter = DriveApp.getRootFolder().getFoldersByName('Job Training System');
  return sysFolderIter.hasNext() ? sysFolderIter.next() : DriveApp.createFolder('Job Training System');
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

/** Opens the one Training Data spreadsheet for a training requisition. */
function getTrainingDataSpreadsheet(trainingId) {
  if (!trainingId) return null;
  const trainingSheet = getSheet('Trainings');
  if (!trainingSheet) return null;
  const rows = sheetToJson(trainingSheet);
  const training = rows.find(r => String(r.ID || r.TrainingID || r.Code || '').trim() === String(trainingId).trim());
  if (!training) return null;

  const storedSheetId = training.ParticipantsSheetID || training.AttendanceSheetID || training.SessionsSheetID || training.EvaluationSheetID || training.PostSheetID;
  if (storedSheetId) {
    try { return SpreadsheetApp.openById(storedSheetId); } catch (e) {
      Logger.log('Could not open Training Data sheet by ID: ' + e.message);
    }
  }
  if (!training.FolderID) return null;
  try {
    const file = getOrCreateSingleTrainingSheet(DriveApp.getFolderById(training.FolderID), training.Code || training.ID);
    return SpreadsheetApp.openById(file.getId());
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
    } else {
      const tpSheet = getSheet('TrainingParticipants');
      if (tpSheet && tpSheet.getLastRow() > 1) {
        const tpValues = tpSheet.getDataRange().getValues();
        const tpHeaders = tpValues[0].map(h => String(h).trim());
        const tIdIdx = tpHeaders.indexOf('TrainingID');
        const empIdIdx = tpHeaders.indexOf('EmployeeID');
        const empNameIdx = tpHeaders.indexOf('EmployeeName');
        const deptIdx = tpHeaders.indexOf('Department');
        const posIdx = tpHeaders.indexOf('Position');
        const addedIdx = tpHeaders.indexOf('AddedAt');
        const idIdx = tpHeaders.indexOf('ID');

        for (let i = 1; i < tpValues.length; i++) {
          const r = tpValues[i];
          const rTrainId = String(tIdIdx >= 0 ? r[tIdIdx] : r[1]).trim();
          if (rTrainId === String(trainingId).trim()) {
            participantsToSync.push({
              ID: idIdx >= 0 ? r[idIdx] : r[0],
              TrainingID: trainingId,
              EmployeeID: empIdIdx >= 0 ? r[empIdIdx] : r[2],
              EmployeeName: empNameIdx >= 0 ? r[empNameIdx] : r[3],
              Department: deptIdx >= 0 ? r[deptIdx] : r[4],
              Position: posIdx >= 0 ? r[posIdx] : r[5],
              AddedAt: addedIdx >= 0 ? r[addedIdx] : r[6]
            });
          }
        }
      }
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
  let rawFolderId = extractCleanDriveId(targetFolderId || getConfigProperty('TRAINING_FOLDER_ID', '') || getConfigProperty('ROOT_FOLDER_ID', ''));
  let targetFolder = null;

  if (rawFolderId) {
    try { targetFolder = DriveApp.getFolderById(rawFolderId); } catch(fErr) {
      Logger.log('Error opening targetFolderId (' + rawFolderId + '): ' + fErr.message);
    }
  }
  if (!targetFolder) {
    targetFolder = getSystemRootFolder();
  }

  const rawTemplateId = extractCleanDriveId(getConfigProperty('TRAINING_REQUISITION_TEMPLATE_ID', ''));
  const fileName = `${code} Training Requisition Form`;
  let copiedFile = null;

  if (rawTemplateId) {
    try {
      const templateFile = DriveApp.getFileById(rawTemplateId);
      try {
        copiedFile = templateFile.makeCopy(fileName);
        if (targetFolder) {
          try {
            copiedFile.moveTo(targetFolder);
          } catch(moveErr) {
            try {
              targetFolder.addFile(copiedFile);
              DriveApp.getRootFolder().removeFile(copiedFile);
            } catch(fallbackMoveErr) {}
          }
        }
      } catch(copyErr1) {
        copiedFile = templateFile.makeCopy(fileName, targetFolder);
      }
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
      if (targetFolder) {
        try {
          copiedFile.moveTo(targetFolder);
        } catch(moveErr2) {
          try {
            targetFolder.addFile(copiedFile);
            DriveApp.getRootFolder().removeFile(copiedFile);
          } catch(fallbackMoveErr2) {}
        }
      }
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
    try { sheet.getRange('A1:I1').merge().setValue('APOLLO FOOD INDUSTRIES (M) SDN BHD').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center'); } catch(m1) {
      sheet.getRange('A1').setValue('APOLLO FOOD INDUSTRIES (M) SDN BHD');
    }
    try { sheet.getRange('A2:I2').merge().setValue('TRAINING REQUISITION FORM (AP-HRD-F01-01)').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center'); } catch(m2) {
      sheet.getRange('A2').setValue('TRAINING REQUISITION FORM (AP-HRD-F01-01)');
    }

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
        sheet.getRange('A15:I39').clearContent();
        partList.slice(0, 25).forEach((p, index) => {
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

    sheet.getRange('A42').setValue('REQUEST BY:').setFontWeight('bold');
    sheet.getRange('C42').setValue('VERIFIED BY (HOD):').setFontWeight('bold');
    sheet.getRange('D42').setValue('APPROVED BY (C-Suite):').setFontWeight('bold');
    sheet.getRange('F42').setValue('APPROVED BY (HOHR):').setFontWeight('bold');
    sheet.getRange('H42').setValue('HR DEPT (Arina):').setFontWeight('bold');

    if (requesterSigData) {
      const empNo    = requesterSigData.employeeNo || requesterSigData.EmployeeNo || requesterSigData.EmployeeID || requesterSigData.ID || '';
      const empName  = requesterSigData.name || requesterSigData.EmployeeName || requesterSigData.Name || '';
      const position = requesterSigData.position || requesterSigData.JobPosition || requesterSigData.Position || requesterSigData.JobTitle || 'Requester';
      const sigDate  = requesterSigData.date || requesterSigData.Date || formatDate(new Date());

      const reqSigText = `${empNo}, ${empName}, ${position}, ${sigDate}`;
      sheet.getRange('A43').setValue(reqSigText);
      sheet.getRange('A44').setValue(`Status: Submitted`);
      sheet.getRange('A45').setValue(`Emp No: ${empNo} | Name: ${empName} | Position: ${position} | Date: ${sigDate}`);
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

    const empNo    = sigData.employeeNo || sigData.EmployeeNo || sigData.EmployeeID || sigData.ID || '';
    const empName  = sigData.name || sigData.EmployeeName || sigData.Name || '';
    const position = sigData.position || sigData.JobPosition || sigData.Position || sigData.JobTitle || '';
    const sigDate  = sigData.date || sigData.Date || sigData.Timestamp || formatDate(new Date());
    const status   = sigData.status || sigData.RequestStatus || sigData.ApprovalStatus || '';

    const reqSigText = `${empNo}, ${empName}, ${position}, ${sigDate}`;
    const appSigText = `${status}, ${empNo}, ${empName}, ${position}, ${sigDate}`;

    if (step === 'request') {
      try { sheet.getRange('A43:B43').merge(); } catch(mErr) {}
      try { sheet.getRange('A44:B44').merge(); } catch(mErr) {}
      try { sheet.getRange('A45:B45').merge(); } catch(mErr) {}
      sheet.getRange('A43').setValue(reqSigText);
      sheet.getRange('A44').setValue(`Status: Submitted`);
      sheet.getRange('A45').setValue(`Emp No: ${empNo} | Name: ${empName} | Position: ${position} | Date: ${sigDate}`);
    } else if (step === 'HOD' || step === 'Head of Department') {
      sheet.getRange('C43').setValue(appSigText);
      sheet.getRange('C44').setValue(`Status: ${status}`);
      sheet.getRange('C45').setValue(`HOD: ${empName} (${empNo}) | ${position} | Date: ${sigDate}`);
    } else if (step === 'Csuite' || step === 'C-Suite') {
      try { sheet.getRange('D43:E43').merge(); } catch(mErr) {}
      try { sheet.getRange('D44:E44').merge(); } catch(mErr) {}
      try { sheet.getRange('D45:E45').merge(); } catch(mErr) {}
      sheet.getRange('D43').setValue(appSigText);
      sheet.getRange('D44').setValue(`Status: ${status}`);
      sheet.getRange('D45').setValue(`C-Suite: ${empName} (${empNo}) | ${position} | Date: ${sigDate}`);
    } else if (step === 'HOHR' || step === 'Head of HR') {
      try { sheet.getRange('F43:G43').merge(); } catch(mErr) {}
      try { sheet.getRange('F44:G44').merge(); } catch(mErr) {}
      try { sheet.getRange('F45:G45').merge(); } catch(mErr) {}
      sheet.getRange('F43').setValue(appSigText);
      sheet.getRange('F44').setValue(`Status: ${status}`);
      sheet.getRange('F45').setValue(`HOHR: ${empName} (${empNo}) | ${position} | Date: ${sigDate}`);
    } else if (step === 'HR' || step === 'Arina' || step === 'HR Department') {
      try { sheet.getRange('H43:I43').merge(); } catch(mErr) {}
      try { sheet.getRange('H44:I44').merge(); } catch(mErr) {}
      try { sheet.getRange('H45:I45').merge(); } catch(mErr) {}
      sheet.getRange('H43').setValue(appSigText);
      sheet.getRange('H44').setValue(`Status: ${status}`);
      sheet.getRange('H45').setValue(`HR: ${empName} (${empNo}) | ${position} | Date: ${sigDate}`);
    }

    // Auto-check: If step is NOT 'request' but cell A43 is empty, ensure requester signature is populated too
    if (step !== 'request') {
      const currentReqVal = sheet.getRange('A43').getValue();
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
          const rText = `${reqId}, ${reqName}, ${reqPos}, ${reqDate}`;
          try { sheet.getRange('A43:B43').merge(); } catch(mErr) {}
          try { sheet.getRange('A44:B44').merge(); } catch(mErr) {}
          try { sheet.getRange('A45:B45').merge(); } catch(mErr) {}
          sheet.getRange('A43').setValue(rText);
          sheet.getRange('A44').setValue(`Status: Submitted`);
          sheet.getRange('A45').setValue(`Emp No: ${reqId} | Name: ${reqName} | Position: ${reqPos} | Date: ${reqDate}`);
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

    // Prefer the per-training Training Data sheet. Fall back to the legacy
    // master participant tab only for older requests.
    const trainingData = getTrainingDataSpreadsheet(trainingId);
    const dataParticipantSheet = trainingData ? trainingData.getSheetByName('TrainingParticipants') : null;
    let tpRows = dataParticipantSheet ? sheetToJson(dataParticipantSheet) : [];
    if (tpRows.length === 0) {
      const tpSheet = getSheet('TrainingParticipants');
      tpRows = tpSheet ? sheetToJson(tpSheet).filter(r => String(r.TrainingID || r.TrainingId || '').trim() === String(trainingId).trim()) : [];
    }
    const resolved = canonicalizeTrainingParticipants(tpRows);
    if (resolved.rejected.length > 0) {
      Logger.log('Skipped participant rows not found in Employees for ' + trainingId + ': ' + resolved.rejected.join(', '));
    }

    const formSpreadsheet = SpreadsheetApp.openById(formId);
    const sheet = formSpreadsheet.getSheetByName('Training Form') || formSpreadsheet.getSheets()[0];
    const rowCount = 25;
    sheet.getRange('A15:I39').clearContent();

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
    'ApprovalStatus', 'RequestedBy', 'RequestedByName', 'RequestedByEmail', 'RequestedDate', 'ApprovedBy', 'ApprovedCostCentre', 'ApprovedAt', 'ApprovalRemarks', 'RescheduledDate'
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

function testDrivePermissions() {
  const root = getSystemRootFolder();
  Logger.log('Root folder ID: ' + (root ? root.getId() : 'None'));
  const templateId = extractCleanDriveId(getConfigProperty('TRAINING_REQUISITION_TEMPLATE_ID', ''));
  if (templateId) {
    try {
      const file = DriveApp.getFileById(templateId);
      Logger.log('Template file accessible: ' + file.getName());
    } catch(e) {
      Logger.log('Template test error: ' + e.message);
    }
  }
}
