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
  const rootFolderId = getConfigProperty('ROOT_FOLDER_ID', '');

  if (rootFolderId) {
    try {
      return DriveApp.getFolderById(rootFolderId);
    } catch (e) {
      Logger.log('Could not open configured ROOT_FOLDER_ID (' + rootFolderId + '): ' + e.message);
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
 * Workspace Structure:
 * Job Training System/
 * └── Training/
 *     └── TR-2026-001 Safety Induction/
 *         ├── TR-2026-001 Attendance Sheet
 *         ├── TR-2026-001 Training Evaluation Sheet
 *         ├── TR-2026-001 Post Evaluation Sheet
 *         └── TR-2026-001 Training Requisition Form
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
      DriveApp.getRootFolder().removeFile(file);
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
 * Copies the AP-HRD-F01-00 Google Sheets master into a training workspace and
 * fills its programme-detail fields without changing the template formatting.
 * The configured template must be a native Google Sheet, not an Excel file.
 */
function createTrainingRequisitionForm(code, training, targetFolderId) {
  const templateId = getConfigProperty('TRAINING_REQUISITION_TEMPLATE_ID', '');
  if (!templateId) {
    return { message: 'Requisition form skipped: configure TRAINING_REQUISITION_TEMPLATE_ID.' };
  }
  if (!targetFolderId) return { message: 'Requisition form skipped: training workspace is unavailable.' };

  try {
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const templateFile = DriveApp.getFileById(templateId);
    const fileName = `${code} Training Requisition Form`;
    const copiedFile = templateFile.makeCopy(fileName, targetFolder);
    const spreadsheet = SpreadsheetApp.openById(copiedFile.getId());
    const sheet = spreadsheet.getSheetByName('Training Form') || spreadsheet.getSheets()[0];

    const startDate = formatDate(training.StartDate);
    const endDate = formatDate(training.EndDate);
    const duration = Number(training.Duration) || 1;
    const hours = training.TotalHours ? ` (${training.TotalHours} hours)` : '';

    // Equip programme detail fields with Google Sheets formulas for date range & duration
    sheet.getRange('C5').setValue(training.Name || '');
    sheet.getRange('C6').setValue(training.CourseFee !== undefined ? training.CourseFee : '');

    if (endDate && endDate !== startDate) {
      sheet.getRange('F6').setFormula(`="${startDate}" & " - " & "${endDate}"`);
    } else {
      sheet.getRange('F6').setValue(startDate || '');
    }

    sheet.getRange('C7').setFormula(`="${duration} day" & IF(${duration}>1, "s", "") & "${hours}"`);
    sheet.getRange('F7').setValue(training.Venue || '');
    sheet.getRange('C8').setValue(training.Trainer || '');
    sheet.getRange('A11').setValue(training.Objectives || '');
    SpreadsheetApp.flush();

    return { fileId: copiedFile.getId(), fileUrl: copiedFile.getUrl(), fileName: fileName };
  } catch (e) {
    Logger.log('createTrainingRequisitionForm error: ' + e.message);
    return { message: 'Requisition form skipped: ' + e.message };
  }
}

/** Keeps the participant grid on the requisition form in sync with enrolled training participants and attendance. */
function syncTrainingRequisitionParticipants(trainingId) {
  try {
    const trainingSheet = getSheet(SHEET_NAMES.trainings);
    const trainingHeaders = ensureTrainingSheetColumns(trainingSheet);
    const trainingRow = findRowById(trainingSheet, trainingId);
    if (trainingRow === -1) return;
    const formColumn = trainingHeaders.indexOf('RequisitionFormFileID') + 1;
    const formId = formColumn ? trainingSheet.getRange(trainingRow, formColumn).getValue() : '';
    if (!formId) return;

    // Combine participants from TrainingParticipants sheet & Attendance sheet
    const tpSheet = getSheet(SHEET_NAMES.trainingParticipants);
    const tpRows  = tpSheet ? sheetToJson(tpSheet).filter(row => String(row.TrainingID) === String(trainingId)) : [];

    const attSheet = getSheet(SHEET_NAMES.attendance);
    const attRows  = attSheet ? sheetToJson(attSheet).filter(row => String(row.TrainingID) === String(trainingId)) : [];

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

    const employeeRows = sheetToJson(getSheet(SHEET_NAMES.employees));
    const employees = {};
    employeeRows.forEach(row => { employees[row.ID] = row; });

    const formSpreadsheet = SpreadsheetApp.openById(formId);
    const sheet = formSpreadsheet.getSheetByName('Training Form') || formSpreadsheet.getSheets()[0];
    const rowCount = 25; // Rows 15-39 in the supplied AP-HRD-F01-00 template.
    sheet.getRangeList(['A15:A39', 'C15:C39', 'D15:D39', 'E15:E39', 'G15:G39']).clearContent();

    uniqueParticipants.slice(0, rowCount).forEach((participant, index) => {
      const employee = employees[participant.EmployeeID] || {};
      const row = 15 + index;
      sheet.getRange(`A${row}`).setValue(participant.EmployeeID || '');
      sheet.getRange(`C${row}`).setValue(participant.EmployeeName || employee.Name || '');
      sheet.getRange(`D${row}`).setValue(participant.Department || employee.Department || '');
      sheet.getRange(`E${row}`).setValue(employee.NRIC || '');
      sheet.getRange(`G${row}`).setValue(participant.Position || employee.Position || employee.JobTitle || '');
    });

    trainingSheet.getRange(trainingRow, 15).setValue(uniqueParticipants.length);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncTrainingRequisitionParticipants error: ' + e.message);
  }
}

// ─── Template Data Population Utilities ───────────────────────────────────────

/**
 * Copies a Google Sheet template file and populates placeholder tags (e.g., {{TRAINING_NAME}}) 
 * or writes cell mappings and data rows according to your Job Training Form layout.
 *
 * @param {string} templateFileId - The Drive File ID of your master form template
 * @param {Folder} targetFolder - Drive Folder to place the populated file
 * @param {string} newFileName - Desired name for the generated file
 * @param {Object} replacements - Key-value dictionary of placeholders to replace, e.g. { "{{TRAINING_NAME}}": "Safety Induction" }
 * @param {Array<Array>} tableData - Optional 2D array of rows to populate into the template's data table starting at startRow
 * @param {number} startRow - 1-based row index to start appending table data (default 10)
 * @return {Object} Information about the generated file (fileId, fileUrl)
 */


/**
 * Appends a new attendance record row to the training's dedicated Drive Attendance Sheet.
 */
/**
 * Appends a new attendance record row to the training's dedicated single Drive spreadsheet (Attendance tab).
 */
function syncAttendanceToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15]; // Column P = FolderID
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

/**
 * Appends a new training evaluation record row to the training's dedicated single Drive spreadsheet (TrainingEval tab).
 */
function syncEvaluationToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15]; // Column P = FolderID
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

/**
 * Appends a new 6-month post evaluation record row to the training's dedicated single Drive spreadsheet (PostEval tab).
 */
function syncPostEvalToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15]; // Column P = FolderID
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

/**
 * Appends a new training session row to the training's dedicated single Drive spreadsheet (TrainingSessions tab).
 */
function syncSessionToTrainingDriveSheet(trainingId, recordRow) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15]; // Column P = FolderID
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

/**
 * Syncs the full list of enrolled training participants to the training's dedicated single Drive spreadsheet (TrainingParticipants tab).
 */
function syncParticipantsToTrainingDriveSheet(trainingId) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row === -1) return;

    const tData = tSheet.getRange(row, 1, 1, 27).getValues()[0];
    const folderId = tData[15]; // Column P = FolderID
    const code     = tData[1] || trainingId;

    if (!folderId) return;

    const tpSheet = getSheet(SHEET_NAMES.trainingParticipants);
    const tpRows = tpSheet ? sheetToJson(tpSheet).filter(r => String(r.TrainingID) === String(trainingId)) : [];

    const folder = DriveApp.getFolderById(folderId);
    const headers = ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt'];
    const file = getOrCreateSingleTrainingSheet(folder, code);

    const ss = SpreadsheetApp.openById(file.getId());
    let sheet = ss.getSheetByName('TrainingParticipants') || ss.getActiveSheet();

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
    }

    if (tpRows.length > 0) {
      const dataToAppend = tpRows.map(p => [
        p.ID || generateId('TP'),
        trainingId,
        p.EmployeeID || '',
        p.EmployeeName || p.Name || '',
        p.Department || '',
        p.Position || p.JobTitle || '',
        p.AddedAt || now()
      ]);
      sheet.getRange(2, 1, dataToAppend.length, headers.length).setValues(dataToAppend);
    }
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('syncParticipantsToTrainingDriveSheet error: ' + e.message);
  }
}
