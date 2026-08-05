/**
 * Training.gs — Training Programme management
 */

const LIFECYCLE_STAGES = [
  'Created',
  'Participants Imported',
  'Attendance In Progress',
  'Training Completed',
  'Evaluation Completed',
  'Waiting for 6-Month Review',
  'Programme Closed'
];

// ─── Read ───────────────────────────────────────────────────────────────────────
function getTrainings() {
  try {
    const rows = autoUpdateTrainingLifecycleStages();
    return ok(rows);
  } catch (e) {
    return err('Failed to load trainings: ' + e.message);
  }
}

function getTrainingById(id) {
  try {
    const rows = autoUpdateTrainingLifecycleStages();
    const t = rows.find(r => r.ID === id);
    if (!t) return err('Training not found.');
    return ok(t);
  } catch (e) {
    return err(e.message);
  }
}

/**
 * Automatically checks and updates lifecycle stages & status for trainings based on training start and end dates.
 */
function autoUpdateTrainingLifecycleStages() {
  try {
    const sheet = getSheet(SHEET_NAMES.trainings);
    if (!sheet) return [];
    ensureTrainingSheetColumns(sheet);
    const rows = sheetToJson(sheet);
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let sheetModified = false;

    rows.forEach(t => {
      let isUpdated = false;
      const startDateStr = t.StartDate;
      const endDateStr   = t.EndDate || t.StartDate;

      if (startDateStr) {
        const startDate = new Date(startDateStr);
        const endDate   = new Date(endDateStr);

        if (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
          startDate.setHours(0, 0, 0, 0);
          endDate.setHours(23, 59, 59, 999);

          // 1. If today has reached or is within training date range (StartDate <= today <= EndDate)
          if (today >= startDate && today <= endDate) {
            if (['Upcoming', 'Draft'].includes(t.Status)) {
              t.Status = 'In Progress';
              isUpdated = true;
            }
            if (['Created', 'Participants Imported'].includes(t.Stage)) {
              t.Stage = 'Attendance In Progress';
              isUpdated = true;
            }
          }
          // 2. If training end date has passed (today > EndDate)
          else if (today > endDate) {
            if (['Created', 'Participants Imported', 'Attendance In Progress'].includes(t.Stage)) {
              t.Stage = 'Training Completed';
              isUpdated = true;
            }
            if (['Upcoming', 'In Progress', 'Draft'].includes(t.Status)) {
              t.Status = 'Completed';
              isUpdated = true;
            }
          }

          // 3. Check 3-month milestone (approx 90 days after endDate)
          const diffMs = today.getTime() - endDate.getTime();
          const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
          t.daysSinceEnd = diffDays;
          t.isThreeMonthsReached = diffDays >= 90;
          t.isSixMonthsReached = diffDays >= 90; // Backward compatibility alias

          if (t.isThreeMonthsReached) {
            if (['Training Completed', 'Evaluation Completed'].includes(t.Stage)) {
              t.Stage = 'Waiting for 3-Month Review';
              isUpdated = true;
            }
          }

          if (isUpdated && t._row) {
            sheet.getRange(t._row, 13).setValue(t.Status); // Col M = Status
            sheet.getRange(t._row, 14).setValue(t.Stage);  // Col N = Stage
            sheet.getRange(t._row, 25).setValue(now());    // Col Y = UpdatedDate
            sheetModified = true;
          }
        }
      } else {
        t.isThreeMonthsReached = false;
        t.isSixMonthsReached = false;
      }
    });

    if (sheetModified) SpreadsheetApp.flush();
    return rows;
  } catch (e) {
    Logger.log('autoUpdateTrainingLifecycleStages error: ' + e.message);
    const sheet = getSheet(SHEET_NAMES.trainings);
    return sheet ? sheetToJson(sheet) : [];
  }
}

/**
 * Automated 3-Month Post Evaluation Email Trigger Engine
 * Sends email notifications to corresponding HODs after 3 months (90 days) post course completion.
 */
function send3MonthPostEvalNotifications() {
  try {
    const trainings = autoUpdateTrainingLifecycleStages();
    const hodPortalUrl = getHodPortalUrl();
    let sentCount = 0;

    trainings.forEach(t => {
      if (t.isThreeMonthsReached && ['Training Completed', 'Evaluation Completed', 'Waiting for 3-Month Review'].includes(t.Stage)) {
        const hodEmail = getConfigProperty('ADMIN_EMAILS', '');
        const postEvalUrl = hodPortalUrl ? `${hodPortalUrl}?page=posteval&id=${t.ID}` : getAppUrl();

        if (hodEmail) {
          const subject = `[TrainHub] 3-Month Post-Training Evaluation Due - ${t.Name}`;
          const body = `Dear HOD / Manager,\n\n` +
            `The 3-month milestone after course completion has elapsed for training programme:\n` +
            `Training Name: ${t.Name} (${t.Code || t.ID})\n` +
            `Completed Date: ${t.EndDate || t.StartDate}\n\n` +
            `Please click the link below to answer the post evaluation form for participants under your Cost Centre:\n` +
            `${postEvalUrl}\n\n` +
            `Thank you,\nTrainHub Training Management System`;

          MailApp.sendEmail(hodEmail, subject, body);
          sentCount++;
        }
      }
    });

    return ok(`Successfully sent 3-month post evaluation email notifications for ${sentCount} training programmes.`);
  } catch (e) {
    Logger.log('send3MonthPostEvalNotifications error: ' + e.message);
    return err('Failed to send 3-month post evaluation emails: ' + e.message);
  }
}

// ─── Create ─────────────────────────────────────────────────────────────────────
function addTraining(data) {
  try {
    if (!data.Name || !data.Trainer || !data.StartDate)
      return err('Name, Trainer, and Start Date are required.');

    const sheet = getSheet(SHEET_NAMES.trainings);
    ensureTrainingSheetColumns(sheet);
    const id = generateId('TRN');
    const code = generateTrainingCode(data.Category);

    // Automatically create dedicated Google Drive workspace for this training
    const workspace = createTrainingWorkspace(code, data.Name);
    const requisitionForm = createTrainingRequisitionForm(code, data, workspace.folderId);

    const timeNow = now();

    sheet.appendRow([
      id,
      code,
      data.Name,
      data.Category    || 'General',
      data.Trainer,
      data.Venue       || '',
      data.StartDate,
      data.EndDate     || '',
      data.Duration    || 1,
      data.TotalHours  || 8,
      data.Department  || '',
      data.Objectives  || '',
      data.Status      || 'Draft',
      data.Stage       || 'Created',
      data.Participants|| 0,
      workspace.folderId          || '',
      workspace.partSheetId       || '',
      workspace.sessionSheetId    || '',
      workspace.attendanceSheetId || '',
      workspace.evaluationSheetId || '',
      workspace.postSheetId       || '',
      requisitionForm.fileId      || '',
      timeNow,
      timeNow,
      data.CourseFee !== undefined ? data.CourseFee : ''
    ]);

    return ok({
      id: id,
      code: code,
      folderId: workspace.folderId,
      folderUrl: workspace.folderUrl,
      requisitionFormUrl: requisitionForm.fileUrl || '',
      requisitionFormMessage: requisitionForm.message || '',
      message: requisitionForm.fileId
        ? 'Training programme, Drive workspace, and requisition form created successfully.'
        : 'Training programme and Drive workspace created successfully. ' + (requisitionForm.message || '')
    });
  } catch (e) {
    return err('Failed to add training: ' + e.message);
  }
}

// ─── Update ─────────────────────────────────────────────────────────────────────
function updateTraining(data) {
  try {
    if (!data.ID) return err('Training ID is required.');

    const sheet = getSheet(SHEET_NAMES.trainings);
    const headers = ensureTrainingSheetColumns(sheet);
    const row = findRowById(sheet, data.ID);
    if (row === -1) return err('Training not found.');

    const dataRange = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

    sheet.getRange(row, 1, 1, 25).setValues([[
      data.ID,
      data.Code        || dataRange[1]  || '',
      data.Name        || dataRange[2]  || '',
      data.Category    || dataRange[3]  || '',
      data.Trainer     || dataRange[4]  || '',
      data.Venue       || dataRange[5]  || '',
      data.StartDate   || dataRange[6]  || '',
      data.EndDate     || dataRange[7]  || '',
      data.Duration    || dataRange[8]  || 1,
      data.TotalHours  || dataRange[9]  || 8,
      data.Department  || dataRange[10] || '',
      data.Objectives  || dataRange[11] || '',
      data.Status      || dataRange[12] || 'Draft',
      data.Stage       || dataRange[13] || 'Created',
      data.Participants|| dataRange[14] || 0,
      dataRange[15]    || '', // Preserve FolderID
      dataRange[16]    || '', // Preserve ParticipantsSheetID
      dataRange[17]    || '', // Preserve SessionsSheetID
      dataRange[18]    || '', // Preserve AttendanceSheetID
      dataRange[19]    || '', // Preserve EvaluationSheetID
      dataRange[20]    || '', // Preserve PostSheetID
      dataRange[21]    || '', // Preserve RequisitionFormFileID
      dataRange[22]    || now(), // CreatedDate
      now(),                   // UpdatedDate
      data.CourseFee !== undefined ? data.CourseFee : (dataRange[24] || '')
    ]]);
    const feeColumn = headers.indexOf('CourseFee') + 1;
    if (feeColumn) {
      sheet.getRange(row, feeColumn).setValue(
        data.CourseFee !== undefined ? data.CourseFee : (dataRange[feeColumn - 1] || '')
      );
    }
    return ok({ message: 'Training updated successfully.' });
  } catch (e) {
    return err('Failed to update training: ' + e.message);
  }
}

// ─── Update Stage Only ──────────────────────────────────────────────────────────
function updateTrainingStage(trainingId, newStage) {
  try {
    if (!LIFECYCLE_STAGES.includes(newStage))
      return err('Invalid lifecycle stage.');

    const sheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(sheet, trainingId);
    if (row === -1) return err('Training not found.');

    sheet.getRange(row, 14).setValue(newStage); // Column N = Stage
    sheet.getRange(row, 25).setValue(now());    // Column Y = UpdatedDate
    return ok({ message: 'Stage updated to: ' + newStage });
  } catch (e) {
    return err(e.message);
  }
}

// ─── Delete ─────────────────────────────────────────────────────────────────────
function deleteTraining(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(sheet, id);
    if (row === -1) return err('Training not found.');
    sheet.deleteRow(row);
    return ok({ message: 'Training deleted.' });
  } catch (e) {
    return err('Failed to delete training: ' + e.message);
  }
}

// ─── Dashboard Summary ──────────────────────────────────────────────────────────
function getTrainingSummary() {
  try {
    const sheet = getSheet(SHEET_NAMES.trainings);
    const rows = sheetToJson(sheet);

    return ok({
      total:      rows.length,
      upcoming:   rows.filter(r => r.Status === 'Upcoming').length,
      ongoing:    rows.filter(r => r.Status === 'In Progress').length,
      completed:  rows.filter(r => r.Status === 'Completed').length,
      draft:      rows.filter(r => r.Status === 'Draft').length,
      pendingEval: rows.filter(r => r.Stage === 'Training Completed').length,
      pendingPost: rows.filter(r => r.Stage === 'Waiting for 6-Month Review').length,
    });
  } catch (e) {
    return err(e.message);
  }
}

// ─── Helpers ────────────────────────────────────────────────────────────────────
function generateTrainingCode(category) {
  const prefix = {
    'Leadership & Management': 'LM',
    'Compliance & Regulatory': 'CR',
    'Technology & Innovation': 'TI',
    'Project Management':      'PM',
    'Data & Analytics':        'DA',
    'Customer Experience':     'CX'
  }[category] || 'TR';
  return prefix + '-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-4);
}

// ─── Training Participants Management ──────────────────────────────────────────
function getTrainingParticipants(trainingId) {
  try {
    const sheet = getSheet(SHEET_NAMES.trainingParticipants);
    const rows = sheetToJson(sheet);
    const filtered = rows.filter(r => String(r.TrainingID) === String(trainingId));
    return ok(filtered);
  } catch (e) {
    return err('Failed to get training participants: ' + e.message);
  }
}

function addTrainingParticipants(trainingId, participants) {
  try {
    if (!trainingId) return err('Training ID is required.');
    if (!participants || !Array.isArray(participants) || participants.length === 0) {
      return err('No participants provided.');
    }

    const empSheet = getSheet(SHEET_NAMES.employees);
    const empMap = {};
    if (empSheet) {
      const empRows = sheetToJson(empSheet);
      empRows.forEach(e => {
        const idKey = String(e.ID || e.EmployeeID || '').trim();
        if (idKey) empMap[idKey] = e;
      });
    }

    const sheet = getSheet(SHEET_NAMES.trainingParticipants);
    const existingRows = sheetToJson(sheet).filter(r => String(r.TrainingID) === String(trainingId));
    const existingEmpIds = new Set(existingRows.map(r => r.EmployeeID));

    let addedCount = 0;
    const addedAt = now();

    participants.forEach(p => {
      const empId = String(p.ID || p.EmployeeID || p.EmployeeNo || '').trim();
      if (empId && !existingEmpIds.has(empId)) {
        const dbEmp = empMap[empId] || {};
        const empName = dbEmp.Name || dbEmp.EmployeeName || p.Name || p.EmployeeName || empId;
        const empDept = dbEmp.Department || p.Department || '';
        const empPos  = dbEmp.Position || dbEmp.JobTitle || p.Position || '';

        sheet.appendRow([
          generateId('TP'),
          trainingId,
          empId,
          empName,
          empDept,
          empPos,
          addedAt
        ]);
        existingEmpIds.add(empId);
        addedCount++;
      }
    });

    const totalCount = existingEmpIds.size;
    updateTrainingParticipantCount(trainingId, totalCount);

    try { syncTrainingRequisitionParticipants(trainingId); } catch(e) {}
    try { syncParticipantsToTrainingDriveSheet(trainingId); } catch(e) {}

    return ok({ message: `Added ${addedCount} participants successfully.`, count: totalCount });
  } catch (e) {
    return err('Failed to add participants: ' + e.message);
  }
}

function removeTrainingParticipant(trainingId, employeeId) {
  try {
    const sheet = getSheet(SHEET_NAMES.trainingParticipants);
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(trainingId) && String(data[i][2]) === String(employeeId)) {
        sheet.deleteRow(i + 1);
        foundRow = i;
        break;
      }
    }

    if (foundRow === -1) return err('Participant record not found.');

    const remainingRows = sheetToJson(sheet).filter(r => String(r.TrainingID) === String(trainingId));
    const totalCount = remainingRows.length;
    updateTrainingParticipantCount(trainingId, totalCount);

    try { syncTrainingRequisitionParticipants(trainingId); } catch(e) {}
    try { syncParticipantsToTrainingDriveSheet(trainingId); } catch(e) {}

    return ok({ message: 'Participant removed successfully.', count: totalCount });
  } catch (e) {
    return err('Failed to remove participant: ' + e.message);
  }
}

function updateTrainingParticipantCount(trainingId, count) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(tSheet, trainingId);
    if (row !== -1) {
      tSheet.getRange(row, 15).setValue(count);
      const currentStage = tSheet.getRange(row, 14).getValue();
      if (currentStage === 'Created' && count > 0) {
        tSheet.getRange(row, 14).setValue('Participants Imported');
      }
      tSheet.getRange(row, 25).setValue(now());
    }
  } catch (e) {
    Logger.log('updateTrainingParticipantCount error: ' + e.message);
  }
}
