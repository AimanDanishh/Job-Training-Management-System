/**
 * DriveManager.gs — Google Drive Workspace Automation & Form Template Population
 * Automatically creates and manages training folders, subfolder hierarchies, and template files.
 */

const SUBFOLDER_NAMES = [
  'Attendance',
  'Evaluation',
  'Certificates',
  'Materials',
  'Photos',
  'Reports',
  'Trainer Notes'
];

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
 * 
 * Workspace Structure:
 * Job Training System/
 * └── Training/
 *     └── TR-2026-001 Safety Induction/
 *         ├── Attendance/
 *         ├── Evaluation/
 *         ├── Certificates/
 *         ├── Materials/
 *         ├── Photos/
 *         ├── Reports/
 *         └── Trainer Notes/
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

    const subfolders = {};
    SUBFOLDER_NAMES.forEach(subName => {
      let subIter = mainFolder.getFoldersByName(subName);
      let subFolder = subIter.hasNext() ? subIter.next() : mainFolder.createFolder(subName);
      subfolders[subName] = subFolder;
    });

    // Optionally copy configured template files into subfolders
    copyTemplateIfConfigured('ATTENDANCE_TEMPLATE_ID', subfolders['Attendance'], `${code} Attendance Record`);
    copyTemplateIfConfigured('EVALUATION_TEMPLATE_ID', subfolders['Evaluation'], `${code} Evaluation Form`);
    copyTemplateIfConfigured('CERTIFICATE_TEMPLATE_ID', subfolders['Certificates'], `${code} Certificate Template`);
    copyTemplateIfConfigured('REPORT_TEMPLATE_ID', subfolders['Reports'], `${code} Training Summary Report`);

    return {
      folderId:            mainFolder.getId(),
      folderUrl:           mainFolder.getUrl(),
      attendanceFolderId:  subfolders['Attendance'].getId(),
      evaluationFolderId:  subfolders['Evaluation'].getId(),
      certificateFolderId: subfolders['Certificates'].getId(),
      materialsFolderId:   subfolders['Materials'].getId(),
      photosFolderId:      subfolders['Photos'].getId(),
      reportsFolderId:     subfolders['Reports'].getId(),
      trainerNotesFolderId:subfolders['Trainer Notes'].getId()
    };
  } catch (e) {
    Logger.log('DriveManager.createTrainingWorkspace error: ' + e.message);
    return {
      folderId: '', folderUrl: '', attendanceFolderId: '', evaluationFolderId: '',
      certificateFolderId: '', materialsFolderId: '', photosFolderId: '',
      reportsFolderId: '', trainerNotesFolderId: ''
    };
  }
}

/**
 * Copies a template file into a target folder if the template property is defined.
 */
function copyTemplateIfConfigured(propertyKey, targetFolder, newFileName) {
  const templateId = getConfigProperty(propertyKey, '');
  if (!templateId) return null;

  try {
    const templateFile = DriveApp.getFileById(templateId);
    return templateFile.makeCopy(newFileName, targetFolder);
  } catch (e) {
    Logger.log(`Failed to copy template [${propertyKey}] (${templateId}): ` + e.message);
    return null;
  }
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
    const dateText = endDate && endDate !== startDate ? `${startDate} - ${endDate}` : startDate;
    const duration = training.Duration || 1;
    const hours = training.TotalHours ? ` (${training.TotalHours} hours)` : '';

    // These top-left cells are intentionally used because the supplied form uses merged ranges.
    sheet.getRange('C5').setValue(training.Name || '');
    sheet.getRange('C6').setValue(training.CourseFee !== undefined ? training.CourseFee : '');
    sheet.getRange('F6').setValue(dateText);
    sheet.getRange('C7').setValue(`${duration} day${Number(duration) === 1 ? '' : 's'}${hours}`);
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
function populateSheetTemplate(templateFileId, targetFolder, newFileName, replacements = {}, tableData = [], startRow = 10) {
  try {
    const templateFile = DriveApp.getFileById(templateFileId);
    const copiedFile = templateFile.makeCopy(newFileName, targetFolder);
    const ss = SpreadsheetApp.openById(copiedFile.getId());
    const sheet = ss.getActiveSheet();

    // 1. Replace placeholder text tags across the sheet (e.g., {{TRAINING_NAME}}, {{START_DATE}})
    if (replacements && Object.keys(replacements).length > 0) {
      const range = sheet.getDataRange();
      let values = range.getValues();

      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          if (typeof values[r][c] === 'string') {
            let str = values[r][c];
            Object.entries(replacements).forEach(([key, val]) => {
              if (str.includes(key)) {
                str = str.replace(new RegExp(escapeRegex(key), 'g'), val || '');
              }
            });
            values[r][c] = str;
          }
        }
      }
      range.setValues(values);
    }

    // 2. Populate tabular records (e.g. attendance list / participants)
    if (tableData && tableData.length > 0) {
      const numRows = tableData.length;
      const numCols = tableData[0].length;
      sheet.getRange(startRow, 1, numRows, numCols).setValues(tableData);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      fileId: copiedFile.getId(),
      fileUrl: copiedFile.getUrl(),
      fileName: newFileName
    };
  } catch (e) {
    Logger.log('populateSheetTemplate error: ' + e.message);
    throw new Error('Failed to populate Sheet template: ' + e.message);
  }
}

/**
 * Copies a Google Doc template file and replaces placeholder tags (e.g. {{TRAINING_NAME}}, {{TRAINER_NAME}}).
 */
function populateDocTemplate(templateFileId, targetFolder, newFileName, replacements = {}) {
  try {
    const templateFile = DriveApp.getFileById(templateFileId);
    const copiedFile = templateFile.makeCopy(newFileName, targetFolder);
    const doc = DocumentApp.openById(copiedFile.getId());
    const body = doc.getBody();

    Object.entries(replacements).forEach(([key, val]) => {
      body.replaceText(escapeRegex(key), val || '');
    });

    doc.saveAndClose();

    return {
      success: true,
      fileId: copiedFile.getId(),
      fileUrl: copiedFile.getUrl(),
      fileName: newFileName
    };
  } catch (e) {
    Logger.log('populateDocTemplate error: ' + e.message);
    throw new Error('Failed to populate Doc template: ' + e.message);
  }
}

function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}
