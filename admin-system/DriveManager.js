/**
 * DriveManager.gs — Google Drive Workspace Automation & Form Template Population
 * Automatically creates and manages training folders, subfolder hierarchies, and template files.
 */

const SUBFOLDER_NAMES = [];

/**
 * Gets or creates the root directory for training workspaces.
 * Uses ROOT_FOLDER_ID from Script Properties, or defaults to "Job Training System/Training" in Google Drive.
 */
function getOrCreateRootFolder() {
  const rootFolderId = getConfigProperty('TRAINING_FOLDER_ID', '') || getConfigProperty('ROOT_FOLDER_ID', '') || getConfigProperty('DRIVE_ROOT_FOLDER_ID', '');

  if (rootFolderId) {
    try {
      return DriveApp.getFolderById(rootFolderId);
    } catch (e) {
      Logger.log('Could not open configured folder (' + rootFolderId + '): ' + e.message);
    }
  }

  // Fallback: Create or retrieve "Job Training System/Training" hierarchy in Google Drive
  let sysFolderIter = DriveApp.getRootFolder().getFoldersByName('Job Training System');
  let sysFolder = sysFolderIter.hasNext() ? sysFolderIter.next() : DriveApp.createFolder('Job Training System');

  let trainFolderIter = sysFolder.getFoldersByName('Training');
  let trainFolder = trainFolderIter.hasNext() ? trainFolderIter.next() : sysFolder.createFolder('Training');

  return trainFolder;
}

/**
 * Automatically creates a dedicated Google Drive workspace for a training programme.
 * Creates per-training sheets directly in the training folder (1 folder per training).
 * 
 * @param {string} code - Training Code (e.g. TR-2026-001)
 * @param {string} name - Training Name (e.g. Safety Induction)
 * @return {Object} Workspace folder IDs and URL dictionary
 */
function createTrainingWorkspace(code, name) {
  try {
    const parentFolder = getOrCreateRootFolder();
    const folderName = `${code} ${name}`.trim();

    // Check if training workspace folder already exists
    let existingIter = parentFolder.getFoldersByName(folderName);
    let mainFolder = existingIter.hasNext() ? existingIter.next() : parentFolder.createFolder(folderName);

    // Create 1 single Google Sheet containing all tabs for this training
    let singleSheetFile = getOrCreateSingleTrainingSheet(mainFolder, code);
    const singleSheetId = singleSheetFile.getId();

    return {
      folderId:            mainFolder.getId(),
      folderUrl:           mainFolder.getUrl(),
      partSheetId:         singleSheetId,
      sessionSheetId:      singleSheetId,
      attendanceSheetId:   singleSheetId,
      evaluationSheetId:   singleSheetId,
      postSheetId:         singleSheetId,
      attendanceFolderId:  mainFolder.getId(),
      evaluationFolderId:  mainFolder.getId(),
      certificateFolderId: mainFolder.getId(),
      materialsFolderId:   mainFolder.getId(),
      photosFolderId:      mainFolder.getId(),
      reportsFolderId:     mainFolder.getId(),
      trainerNotesFolderId:mainFolder.getId()
    };
  } catch (e) {
    Logger.log('DriveManager.createTrainingWorkspace error: ' + e.message);
    return {
      folderId: '', folderUrl: '', partSheetId: '', sessionSheetId: '', attendanceSheetId: '', evaluationSheetId: '', postSheetId: '',
      attendanceFolderId: '', evaluationFolderId: '', certificateFolderId: '',
      materialsFolderId: '', photosFolderId: '', reportsFolderId: '', trainerNotesFolderId: ''
    };
  }
}

/**
 * Gets or creates the single Google Sheet file for a training workspace,
 * ensuring tabs exist for TrainingParticipants, TrainingSessions, Attendance, TrainingEval, PostEval, and Summary.
 */
function getOrCreateSingleTrainingSheet(folder, code) {
  const fileName = `${code} Training Data`;
  let fileIter = folder.getFilesByName(fileName);
  let file;
  let ss;

  if (fileIter.hasNext()) {
    file = fileIter.next();
    ss = SpreadsheetApp.openById(file.getId());
  } else {
    // Also check for legacy sheet names in case of existing folders
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

  // Remove default sheet if present and empty
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Data');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() <= 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  // Ensure Summary tab
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

/**
 * Copies the AP-HRD-F01-01 Google Sheets master into a training workspace and
 * fills its programme-detail fields without changing the template formatting.
 */
function createTrainingRequisitionForm(code, training, targetFolderId, requesterSigData) {
  if (!targetFolderId) return { message: 'Requisition form skipped: training workspace is unavailable.' };
  const templateId = getConfigProperty('TRAINING_REQUISITION_TEMPLATE_ID', '');

  try {
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const fileName = `${code} Training Requisition Form`;
    let copiedFile = null;

    if (templateId) {
      try {
        const templateFile = DriveApp.getFileById(templateId);
        copiedFile = templateFile.makeCopy(fileName, targetFolder);
      } catch (tmplErr) {
        Logger.log('Template copy error, creating clean form: ' + tmplErr.message);
      }
    }

    if (!copiedFile) {
      const newSs = SpreadsheetApp.create(fileName);
      copiedFile = DriveApp.getFileById(newSs.getId());
      targetFolder.addFile(copiedFile);
      try { DriveApp.getRootFolder().removeFile(copiedFile); } catch(rErr) {}
    }

    const spreadsheet = SpreadsheetApp.openById(copiedFile.getId());
    let sheet = spreadsheet.getSheetByName('Training Form');
    if (!sheet) {
      sheet = spreadsheet.getActiveSheet();
      sheet.setName('Training Form');
    }

    const startDate = formatDate(training.StartDate);
    const endDate = formatDate(training.EndDate);
    const duration = Number(training.Duration) || 1;
    const hours = training.TotalHours ? ` (${training.TotalHours} hours)` : '';

    sheet.getRange('A1:I1').merge().setValue('APOLLO FOOD INDUSTRIES (M) SDN BHD').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    sheet.getRange('A2:I2').merge().setValue('TRAINING REQUISITION FORM (AP-HRD-F01-01)').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

    const setTemplateValue = (area, value) => sheet.getRange(area.split(':')[0]).setValue(value == null ? '' : value);
    sheet.getRange('A5').setValue('Training Title:');
    setTemplateValue('C5:I5', training.Name || training.TrainingName || '');
    sheet.getRange('A6').setValue('Course Fee (RM):');
    setTemplateValue('C6:E6', training.CourseFee !== undefined ? training.CourseFee : '0.00');
    sheet.getRange('F6').setValue('Date:');
    setTemplateValue('G6:I6', endDate && endDate !== startDate ? `${startDate} - ${endDate}` : (startDate || ''));
    sheet.getRange('A7').setValue('Duration:');
    setTemplateValue('C7:E7', `${duration} day(s)${hours}`);
    sheet.getRange('F7').setValue('Venue:');
    setTemplateValue('G7:I7', training.Venue || '');
    sheet.getRange('A8').setValue('Training Provider:');
    setTemplateValue('C8:I8', training.TrainingProvider || training.Provider || training.Trainer || '');
    sheet.getRange('A10').setValue('Reasons for Training:');
    setTemplateValue('A11:I12', training.Objectives || training.Reason || '');

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
  } catch (e) {
    Logger.log('createTrainingRequisitionForm error: ' + e.message);
    return { message: 'Requisition form creation failed: ' + e.message };
  }
}

/**
 * Updates signature blocks on AP-HRD-F01-01 form.
 */
function updateTrainingRequisitionSignatures(trainingId, step, sigData, targetFormId) {
  try {
    SpreadsheetApp.flush();
    const trainingSheet = getSheet(SHEET_NAMES.trainings);
    if (!trainingSheet) return;
    const headers = ensureTrainingSheetColumns(trainingSheet);
    const row = findRowById(trainingSheet, trainingId);
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
        if (reqId && getSheet(SHEET_NAMES.employees)) {
          const emps = sheetToJson(getSheet(SHEET_NAMES.employees));
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
  } catch (e) {
    Logger.log('updateTrainingRequisitionSignatures error: ' + e.message);
  }
}

/** Keeps participant grid on requisition form in sync. */
function syncTrainingRequisitionParticipants(trainingId) {
  try {
    const trainingSheet = getSheet(SHEET_NAMES.trainings);
    const trainingHeaders = ensureTrainingSheetColumns(trainingSheet);
    const trainingRow = findRowById(trainingSheet, trainingId);
    if (trainingRow === -1) return;
    const formColumn = trainingHeaders.indexOf('RequisitionFormFileID') + 1;
    const formId = formColumn ? trainingSheet.getRange(trainingRow, formColumn).getValue() : '';
    if (!formId) return;

    const ss = getTrainingDataSpreadsheet(trainingId);
    const tpSheet = ss ? ss.getSheetByName('TrainingParticipants') : null;
    const tpRows  = tpSheet ? sheetToJson(tpSheet) : [];

    const attSheet = ss ? ss.getSheetByName('Attendance') : null;
    const attRows  = attSheet ? sheetToJson(attSheet) : [];

    const uniqueParticipants = [];
    const seen = {};

    tpRows.forEach(row => {
      const empId = row.EmployeeID || row.EmployeeNo || row.ID;
      if (empId && !seen[empId]) {
        seen[empId] = true;
        uniqueParticipants.push({
          EmployeeID: empId,
          EmployeeName: row.EmployeeName || row.Name || '',
          Department: row.Department || '',
          Position: row.Position || row.JobTitle || ''
        });
      }
    });

    attRows.forEach(row => {
      const empId = row.EmployeeNo || row.EmployeeID;
      if (empId && !seen[empId]) {
        seen[empId] = true;
        uniqueParticipants.push({
          EmployeeID: empId,
          EmployeeName: row.EmployeeName || '',
          Department: row.Department || '',
          Position: row.Position || ''
        });
      }
    });

    const resolved = canonicalizeTrainingParticipants(uniqueParticipants);
    if (resolved.rejected.length > 0) {
      Logger.log('Skipped participant rows not found in Employees for ' + trainingId + ': ' + resolved.rejected.join(', '));
    }

    const formSpreadsheet = SpreadsheetApp.openById(formId);
    const sheet = formSpreadsheet.getSheetByName('Training Form') || formSpreadsheet.getSheets()[0];
    const rowCount = 25;
    sheet.getRange('A15:I39').clearContent();

    resolved.participants.slice(0, rowCount).forEach((participant, index) => {
      const r = 15 + index;
      sheet.getRange(`A${r}`).setValue(participant.ID);
      sheet.getRange(`C${r}`).setValue(participant.Name);
      sheet.getRange(`D${r}`).setValue(participant.Department);
      sheet.getRange(`H${r}`).setValue(participant.Position);
    });

    trainingSheet.getRange(trainingRow, 15).setValue(resolved.participants.length);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncTrainingRequisitionParticipants error: ' + e.message);
  }
}

/** Appends attendance log. */
function syncAttendanceToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15];
    const code     = tData[1] || trainingId;

    if (!folderId) return;

    const folder = DriveApp.getFolderById(folderId);
    const file = getOrCreateSingleTrainingSheet(folder, code);

    const ss = SpreadsheetApp.openById(file.getId());
    let sheet = ss.getSheetByName('Attendance') || ss.getActiveSheet();
    sheet.appendRow(recordRow);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncAttendanceToTrainingDriveSheet error: ' + e.message);
  }
}

/** Appends evaluation log. */
function syncEvaluationToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15];
    const code     = tData[1] || trainingId;

    if (!folderId) return;

    const folder = DriveApp.getFolderById(folderId);
    const file = getOrCreateSingleTrainingSheet(folder, code);

    const ss = SpreadsheetApp.openById(file.getId());
    let sheet = ss.getSheetByName('TrainingEval') || ss.getActiveSheet();
    sheet.appendRow(recordRow);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncEvaluationToTrainingDriveSheet error: ' + e.message);
  }
}

/** Appends post eval log. */
function syncPostEvalToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15];
    const code     = tData[1] || trainingId;

    if (!folderId) return;

    const folder = DriveApp.getFolderById(folderId);
    const file = getOrCreateSingleTrainingSheet(folder, code);

    const ss = SpreadsheetApp.openById(file.getId());
    let sheet = ss.getSheetByName('PostEval') || ss.getActiveSheet();
    sheet.appendRow(recordRow);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncPostEvalToTrainingDriveSheet error: ' + e.message);
  }
}

/** Appends session log. */
function syncSessionToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15];
    const code     = tData[1] || trainingId;

    if (!folderId) return;

    const folder = DriveApp.getFolderById(folderId);
    const file = getOrCreateSingleTrainingSheet(folder, code);

    const ss = SpreadsheetApp.openById(file.getId());
    let sheet = ss.getSheetByName('TrainingSessions') || ss.getActiveSheet();
    sheet.appendRow(recordRow);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncSessionToTrainingDriveSheet error: ' + e.message);
  }
}

/** Syncs enrolled training participants to per-training sheet. */
function syncParticipantsToTrainingDriveSheet(trainingId, directParticipantsList) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15];
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
      const tpSheet = getSheet(SHEET_NAMES.trainingParticipants);
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

    // Training Data must contain canonical Employee-sheet details even when
    // the source is an older copied participant record.
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

    const trainingHeaders = ensureTrainingSheetColumns(tSheet);
    const participantSheetIdColumn = trainingHeaders.indexOf('ParticipantsSheetID') + 1;
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

    if (typeof _sheetDataCache !== 'undefined') {
      _sheetDataCache = {};
    }
  } catch (e) {
    Logger.log('syncParticipantsToTrainingDriveSheet error: ' + e.message);
  }
}

function testDrivePermissions() {
  const root = getOrCreateRootFolder();
  Logger.log('Root folder ID: ' + (root ? root.getId() : 'None'));
}
