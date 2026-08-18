/**
 * DriveManager.gs - Google Drive Workspace Automation & Form Template Population
 * Automatically creates and manages training folders, subfolder hierarchies, and template files.
 */

const SUBFOLDER_NAMES = [];

/**
 * Gets or creates the Training directory under the configured ROOT_FOLDER_ID.
 *
 * ROOT_FOLDER_ID is always the system root.  Training workspaces are always
 * stored in ROOT_FOLDER_ID/Training/<training code + name>; never in My Drive
 * or directly in the system root.
 */
function getOrCreateSystemRootFolder() {
  const rootFolderId = getConfigProperty('ROOT_FOLDER_ID', '');
  if (!rootFolderId) throw new Error('ROOT_FOLDER_ID is required.');
  return DriveApp.getFolderById(rootFolderId);
}

function getOrCreateRootFolder() {
  return getOrCreateTrainingRootFolder();
}

function getOrCreateTrainingRootFolder() {
  const configuredFolderId = getConfigProperty('TRAINING_FOLDER', '') || getConfigProperty('TRAINING_FOLDER_ID', '');
  if (configuredFolderId) {
    try {
      return DriveApp.getFolderById(configuredFolderId);
    } catch (e) {
      Logger.log('Could not open configured TRAINING_FOLDER (' + configuredFolderId + '): ' + e.message);
    }
  }

  const systemRoot = getOrCreateSystemRootFolder();
  let iter = systemRoot.getFoldersByName('Training Folder');
  if (iter.hasNext()) return iter.next();
  return systemRoot.createFolder('Training Folder');
}

function getOrCreateReportsFolder() {
  const configuredReportsId = getConfigProperty('REPORTS_FOLDER_ID', '');
  if (configuredReportsId) {
    try {
      return DriveApp.getFolderById(configuredReportsId);
    } catch (e) {
      Logger.log('Could not open REPORTS_FOLDER_ID (' + configuredReportsId + '): ' + e.message);
    }
  }

  const systemRoot = getOrCreateSystemRootFolder();
  let repFolderIter = systemRoot.getFoldersByName('Reports');
  if (repFolderIter.hasNext()) return repFolderIter.next();
  throw new Error("Required folder 'Reports' was not found under ROOT_FOLDER_ID.");
}

function getOrCreateNamelistFolder() {
  const systemRoot = getOrCreateSystemRootFolder();
  let nameFolderIter = systemRoot.getFoldersByName('Namelist');
  if (nameFolderIter.hasNext()) return nameFolderIter.next();
  throw new Error("Required folder 'Namelist' was not found under ROOT_FOLDER_ID.");
}

function createTrainingWorkspace(code, name) {
  try {
    const parentFolder = getOrCreateTrainingRootFolder();
    const folderName = `${code} ${name}`.trim();

    let existingIter = parentFolder.getFoldersByName(folderName);
    let mainFolder = existingIter.hasNext() ? existingIter.next() : parentFolder.createFolder(folderName);

    let singleSheetFile = getOrCreateSingleTrainingSheet(mainFolder, code);
    const singleSheetId = singleSheetFile.getId();

    let brochureFolder = mainFolder.getFoldersByName('Brochure');
    if (!brochureFolder.hasNext()) {
      mainFolder.createFolder('Brochure');
    }

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
 * One-time repair for legacy workspaces created outside ROOT_FOLDER_ID/Training.
 * Only folders referenced by the master Trainings sheet are moved, so unrelated
 * Drive folders are never touched. Run this manually from the Apps Script editor
 * after ROOT_FOLDER_ID has been configured.
 */
function migrateTrainingWorkspacesToConfiguredRoot() {
  const result = { moved: 0, alreadyCorrect: 0, skipped: 0, errors: [] };
  try {
    const trainingRoot = getOrCreateRootFolder();
    const trainingRootId = trainingRoot.getId();
    const sheet = getSheet(SHEET_NAMES.trainings);
    const trainings = sheet ? sheetToJson(sheet) : [];

    trainings.forEach(training => {
      const folderId = String(training.FolderID || '').trim();
      if (!folderId) {
        result.skipped++;
        return;
      }
      try {
        const folder = DriveApp.getFolderById(folderId);
        const parents = folder.getParents();
        let isAlreadyCorrect = false;
        while (parents.hasNext()) {
          if (parents.next().getId() === trainingRootId) {
            isAlreadyCorrect = true;
            break;
          }
        }
        if (isAlreadyCorrect) {
          result.alreadyCorrect++;
        } else {
          folder.moveTo(trainingRoot);
          result.moved++;
        }
      } catch (e) {
        result.errors.push(`${training.ID || folderId}: ${e.message}`);
      }
    });
    return ok(result);
  } catch (e) {
    return err('Workspace migration failed: ' + e.message);
  }
}

/**
 * Gets or creates the single Google Sheet file for a training workspace,
 * ensuring tabs exist for TrainingParticipants, TrainingSessions, Attendance, TrainingEval, PostEval, and Summary.
 */
function getOrCreateSingleTrainingSheet(folder, code) {
  const fileName = 'Training Data';
  let fileIter = folder.getFilesByName(fileName);
  if (!fileIter.hasNext()) fileIter = folder.getFilesByName(`${code} Training Data`);
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
      // SpreadsheetApp creates a file in My Drive first. Move it immediately
      // so the final (and only) location is the training workspace.
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
    const setTemplateValue = (area, value) => sheet.getRange(String(area || '').split(':')[0]).setValue(value == null ? '' : value);
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

    const partList = Array.isArray(training.ParticipantList) ? training.ParticipantList : (Array.isArray(training.participants) ? training.participants : []);
    populateRequisitionFormParticipantsAndTabs(spreadsheet, sheet, partList);

    SpreadsheetApp.flush();
    return { fileId: copiedFile.getId(), fileUrl: copiedFile.getUrl(), fileName: fileName };
  } catch (e) {
    Logger.log('createTrainingRequisitionForm error: ' + e.message);
    return { message: 'Requisition form creation failed: ' + e.message };
  }
}

/**
 * Populates participant list in Training Requisition Form.
 * Automatically paginates > 24 participants across multiple tabs (Training Form, Training Form - Page 2, etc.)
 * by duplicating the primary template sheet so all header info and signature boxes are preserved.
 */
function populateRequisitionFormParticipantsAndTabs(spreadsheet, primarySheet, partList) {
  if (!spreadsheet || !primarySheet) return;
  const list = Array.isArray(partList) ? partList : [];
  const setTemplateValue = (s, area, value) => s.getRange(String(area || '').split(':')[0]).setValue(value == null ? '' : value);

  const PAGE_SIZE = 24;
  const total = list.length;
  const numPages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  const baseSheetName = 'Training Form';

  try {
    if (primarySheet.getName() !== baseSheetName && !primarySheet.getName().includes('Page')) {
      primarySheet.setName(baseSheetName);
    }
  } catch(nameErr) {}

  // 1. Populate Page 1 (indices 0 to 23)
  primarySheet.getRange('A15:I38').clearContent();
  const page1Chunk = list.slice(0, PAGE_SIZE);
  page1Chunk.forEach((p, index) => {
    const r = 15 + index;
    setTemplateValue(primarySheet, `A${r}:B${r}`, p.EmployeeID || p.ID || p.EmployeeNo || '');
    setTemplateValue(primarySheet, `C${r}`, p.EmployeeName || p.Name || '');
    setTemplateValue(primarySheet, `D${r}:G${r}`, p.Department || p.CostCentre || '');
    setTemplateValue(primarySheet, `H${r}:I${r}`, p.Position || p.JobTitle || '');
  });

  // 2. Handle subsequent pages (Page 2, Page 3, etc.)
  for (let p = 2; p <= numPages; p++) {
    const pageName = `${baseSheetName} - Page ${p}`;
    let pageSheet = spreadsheet.getSheetByName(pageName);
    if (!pageSheet) {
      pageSheet = primarySheet.copyTo(spreadsheet);
      pageSheet.setName(pageName);
    }

    pageSheet.getRange('A15:I38').clearContent();
    const pageChunk = list.slice((p - 1) * PAGE_SIZE, p * PAGE_SIZE);
    pageChunk.forEach((item, index) => {
      const r = 15 + index;
      setTemplateValue(pageSheet, `A${r}:B${r}`, item.EmployeeID || item.ID || item.EmployeeNo || '');
      setTemplateValue(pageSheet, `C${r}`, item.EmployeeName || item.Name || '');
      setTemplateValue(pageSheet, `D${r}:G${r}`, item.Department || item.CostCentre || '');
      setTemplateValue(pageSheet, `H${r}:I${r}`, item.Position || item.JobTitle || '');
    });
  }

  // 3. Remove obsolete excess page sheets if participant count was reduced
  const allSheets = spreadsheet.getSheets();
  allSheets.forEach(s => {
    const sName = s.getName();
    const match = sName.match(/^Training Form - Page (\d+)$/i);
    if (match) {
      const pageNum = parseInt(match[1], 10);
      if (pageNum > numPages) {
        try {
          spreadsheet.deleteSheet(s);
        } catch(delErr) {
          Logger.log('Could not delete excess page sheet ' + sName + ': ' + delErr.message);
        }
      }
    }
  });
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
    const allSheets = ss.getSheets();
    const targetSheets = allSheets.filter(s => {
      const sName = s.getName();
      return sName.startsWith('Training Form') || sName.includes('Form') || allSheets.length === 1;
    });

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
          if (reqId && getSheet(SHEET_NAMES.employees)) {
            const emps = sheetToJson(getSheet(SHEET_NAMES.employees));
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
  } catch (e) {
    Logger.log('updateTrainingRequisitionSignatures error: ' + e.message);
  }
}

/** Keeps participant grid on requisition form in sync. */
function syncTrainingRequisitionParticipants(trainingId, directParticipantsList) {
  try {
    const trainingSheet = getSheet(SHEET_NAMES.trainings);
    if (!trainingSheet) return;
    const trainingHeaders = ensureTrainingSheetColumns(trainingSheet);
    const trainingRow = findRowById(trainingSheet, trainingId);
    if (trainingRow === -1) return;
    const formColumn = trainingHeaders.indexOf('RequisitionFormFileID') + 1;
    const formId = formColumn ? trainingSheet.getRange(trainingRow, formColumn).getValue() : '';
    if (!formId) return;

    let participantsToSync = [];
    if (Array.isArray(directParticipantsList) && directParticipantsList.length > 0) {
      participantsToSync = directParticipantsList;
    } else {
      const ss = getTrainingDataSpreadsheet(trainingId);
      const tpSheet = ss ? (ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants')) : null;
      if (tpSheet) {
        participantsToSync = sheetToJson(tpSheet);
      }
    }

    if (!participantsToSync || participantsToSync.length === 0) {
      Logger.log('No participants found to sync for training ' + trainingId);
      return;
    }

    const resolved = canonicalizeTrainingParticipants(participantsToSync);
    const participants = resolved.participants;
    if (resolved.rejected.length > 0) {
      Logger.log('Skipped participant rows not found in Employees for ' + trainingId + ': ' + resolved.rejected.join(', '));
    }

    const formSpreadsheet = SpreadsheetApp.openById(formId);
    const primarySheet = formSpreadsheet.getSheetByName('Training Form') || formSpreadsheet.getSheets()[0];

    populateRequisitionFormParticipantsAndTabs(formSpreadsheet, primarySheet, participants);

    const partCol = trainingHeaders.indexOf('Participants') + 1;
    if (partCol > 0) trainingSheet.getRange(trainingRow, partCol).setValue(participants.length);
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
    let sheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants') || ss.getActiveSheet();

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
