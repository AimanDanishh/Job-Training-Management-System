/**
 * Training.gs - Training Programme management
 */

const LIFECYCLE_STAGES = [
  'Created',
  'Participants Imported',
  'Attendance In Progress',
  'Training Completed',
  'Evaluation Completed',
  'Waiting for 3-Month Review',
  'Programme Closed'
];

// --- Read -----------------------------------------------------------------------
function getTrainings() {
  try {
    const rows = autoUpdateTrainingLifecycleStages();
    return ok(rows);
  } catch (e) {
    Logger.log('getTrainings error: ' + e.message + '\n' + (e.stack || ''));
    return err('Failed to load trainings: ' + e.message);
  }
}

function getTrainingById(id) {
  try {
    const rows = autoUpdateTrainingLifecycleStages();
    const t = rows.find(r => r.ID === id);
    if (!t) return err('Training not found.');

    // Auto-sign HR Department acknowledgment in Training Requisition Form using opening admin's info
    try {
      let activeEmail = '';
      try {
        activeEmail = Session.getActiveUser().getEmail();
      } catch (e) {}
      
      const hrProfile = getHrProfileByEmail(activeEmail);
      acknowledgeHRRequisition(id, {
        employeeNo: hrProfile.employeeNo,
        name: hrProfile.name,
        position: hrProfile.position,
        status: 'Acknowledged'
      });
    } catch(hrErr) {
      Logger.log('Auto HR sign error in getTrainingById: ' + hrErr.message);
    }

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
    if (!sheet) {
      throw new Error(`Training database sheet "${SHEET_NAMES.trainings}" could not be found in the configured TrainHub Spreadsheet.`);
    }
    const headers = ensureTrainingSheetColumns(sheet);
    const rows = sheetToJson(sheet);
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const updatedDateCol = headers.indexOf('UpdatedDate') + 1;
    const courseFeeCol = headers.indexOf('CourseFee') + 1;
    const statusCol = headers.indexOf('Status') + 1;
    const stageCol = headers.indexOf('Stage') + 1;

    let sheetModified = false;

    rows.forEach(t => {
      let isUpdated = false;

      // Database Auto-Repair: Fix any CourseFee values corrupted by previous date timestamp overwrites
      const rawFee = String(t.CourseFee || '').trim();
      if (rawFee && (rawFee.includes(':') || rawFee.includes('2026') || rawFee.includes('Aug') || rawFee.includes('/'))) {
        const defaultFee = '0.00';
        t.CourseFee = defaultFee;
        if (t._row && courseFeeCol > 0) {
          sheet.getRange(t._row, courseFeeCol).setValue(defaultFee);
          sheetModified = true;
        }
      }

      let partCount = Number(t.Participants || 0);

      // Auto-resolve participant count if 0, NaN, or empty
      if (isNaN(partCount) || partCount === 0) {
        // Tier 1: Check ParticipantList JSON string
        if (t.ParticipantList) {
          try {
            const pList = typeof t.ParticipantList === 'string' ? JSON.parse(t.ParticipantList) : t.ParticipantList;
            if (Array.isArray(pList) && pList.length > 0) {
              partCount = pList.length;
            }
          } catch(e) {}
        }
        // Tier 2: Check aliases (TotalPax, Total Participant, Pax, etc.)
        if (partCount === 0) {
          const aliasPax = Number(t.TotalPax || t['Total Pax'] || t['Total Participant'] || t.Pax || 0);
          if (!isNaN(aliasPax) && aliasPax > 0) {
            partCount = aliasPax;
          }
        }
        // Tier 3: Check training participants sheet in Drive or database
        if (partCount === 0 && (t.ID || t.Code)) {
          try {
            const pList = getTrainingParticipantsList(t.ID || t.Code);
            if (Array.isArray(pList) && pList.length > 0) {
              partCount = pList.length;
            }
          } catch(e) {}
        }

        if (partCount > 0) {
          t.Participants = partCount;
          const partCol = headers.indexOf('Participants') + 1;
          if (t._row && partCol > 0) {
            sheet.getRange(t._row, partCol).setValue(partCount);
            sheetModified = true;
          }
        }
      }

      // Auto-advance Created -> Participants Imported if participants are attached/requested
      if (partCount > 0 && (!t.Stage || t.Stage === 'Created')) {
        t.Stage = 'Participants Imported';
        isUpdated = true;
      }

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

          // 3. Check 3-month milestone (90 days after endDate) & countdown calculation
          const targetPostEvalDate = new Date(endDate.getTime() + 90 * 24 * 60 * 60 * 1000);
          const remainingMs = targetPostEvalDate.getTime() - new Date().getTime();
          const diffMs = today.getTime() - endDate.getTime();
          const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

          t.daysSinceEnd = diffDays;
          t.isThreeMonthsReached = diffDays >= 90;
          t.isSixMonthsReached = diffDays >= 90;

          if (remainingMs > 0) {
            const remDays = Math.floor(remainingMs / (1000 * 60 * 60 * 24));
            const remHours = Math.floor(remainingMs / (1000 * 60 * 60));
            if (remDays >= 1) {
              t.countdownText = `${remDays} day${remDays > 1 ? 's' : ''} remaining`;
            } else {
              t.countdownText = `${remHours} hour${remHours > 1 ? 's' : ''} remaining`;
            }
          } else {
            t.countdownText = '3-Month Evaluation Due Now';
          }

          if (t.isThreeMonthsReached) {
            if (['Created', 'Participants Imported', 'Attendance In Progress', 'Training Completed', 'Evaluation Completed'].includes(t.Stage)) {
              t.Stage = 'Waiting for 3-Month Review';
              isUpdated = true;
            }
          }
        }

        if (isUpdated) {
          t._isLifecycleUpdated = true;
          if (statusCol > 0) sheet.getRange(t._row, statusCol).setValue(t.Status);
          if (stageCol > 0) sheet.getRange(t._row, stageCol).setValue(t.Stage);
          if (updatedDateCol > 0) sheet.getRange(t._row, updatedDateCol).setValue(now());
          sheetModified = true;
        }
      } else {
        t.isThreeMonthsReached = false;
        t.isSixMonthsReached = false;
        t.countdownText = 'No date set';
      }
    });

    if (sheetModified) {
      try { SpreadsheetApp.flush(); } catch(fErr) {}
      rows.forEach(t => {
        if (t._isLifecycleUpdated && t.ID) {
          try { syncTrainingById(t.ID, 'System Lifecycle Auto-Advance', 'STATUS_CHANGE'); } catch(sErr) {}
        }
      });
    }
    rows.forEach(t => enrichTrainingWithUrls(t));
    return rows;
  } catch (e) {
    Logger.log('autoUpdateTrainingLifecycleStages error: ' + e.message);
    const sheet = getSheet(SHEET_NAMES.trainings);
    if (!sheet) {
      throw new Error(`Training database sheet "${SHEET_NAMES.trainings}" could not be found: ${e.message}`);
    }
    const fallbackRows = sheetToJson(sheet);
    fallbackRows.forEach(t => enrichTrainingWithUrls(t));
    return fallbackRows;
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

// --- Create ---------------------------------------------------------------------
function addTraining(data) {
  try {
    if (!data.Name || !data.Trainer || !data.StartDate)
      return err('Name, Trainer, and Start Date are required.');

    const sheet = getSheet(SHEET_NAMES.trainings);
    ensureTrainingSheetColumns(sheet);
    const id = generateId('TRN');
    const code = generateTrainingCode(data.Category);

    const timeNow = now();
    let requesterSigData = {
      employeeNo: data.RequestedBy || data.EmployeeID || 'ADMIN',
      name: data.RequestedByName || data.EmployeeName || 'Administrator',
      position: 'Administrator',
      date: timeNow
    };

    if (data.RequestedBy && getSheet(SHEET_NAMES.employees)) {
      try {
        const emps = sheetToJson(getSheet(SHEET_NAMES.employees));
        const m = emps.find(e => String(e.ID || e.EmployeeID).toLowerCase() === String(data.RequestedBy).toLowerCase());
        if (m) {
          requesterSigData.name = m.Name || m.EmployeeName || requesterSigData.name;
          requesterSigData.position = m.Position || m.JobTitle || m.PositionTitle || requesterSigData.position;
        }
      } catch(e) {}
    }

    // Automatically create dedicated Google Drive workspace for this training
    const workspace = createTrainingWorkspace(code, data.Name);
    const requisitionForm = createTrainingRequisitionForm(code, data, workspace.folderId, requesterSigData);

    const participantsList = Array.isArray(data.ParticipantList) ? data.ParticipantList : (Array.isArray(data.participants) ? data.participants : []);

    const headers = ensureTrainingSheetColumns(sheet);
    const rowObj = {};
    headers.forEach(h => rowObj[h] = '');

    rowObj['ID'] = id;
    rowObj['Code'] = code;
    rowObj['Name'] = data.Name || '';
    rowObj['Category'] = data.Category || 'General';
    rowObj['Trainer'] = data.Trainer || '';
    rowObj['Venue'] = data.Venue || '';
    rowObj['StartDate'] = data.StartDate || '';
    rowObj['EndDate'] = data.EndDate || '';
    rowObj['Duration'] = data.Duration || 1;
    rowObj['TotalHours'] = data.TotalHours || 8;
    rowObj['Department'] = data.Department || '';
    rowObj['Objectives'] = data.Objectives || '';
    rowObj['Status'] = data.Status || 'Draft';
    rowObj['Stage'] = data.Stage || (participantsList.length > 0 ? 'Participants Imported' : 'Created');
    rowObj['Participants'] = data.Participants || participantsList.length;
    rowObj['FolderID'] = workspace.folderId || '';
    rowObj['ParticipantsSheetID'] = workspace.partSheetId || '';
    rowObj['SessionsSheetID'] = workspace.sessionSheetId || '';
    rowObj['AttendanceSheetID'] = workspace.attendanceSheetId || '';
    rowObj['EvaluationSheetID'] = workspace.evaluationSheetId || '';
    rowObj['PostSheetID'] = workspace.postSheetId || '';
    rowObj['RequisitionFormFileID'] = requisitionForm.fileId || '';
    rowObj['CreatedDate'] = timeNow;
    rowObj['UpdatedDate'] = timeNow;
    rowObj['CourseFee'] = (data.CourseFee !== undefined && data.CourseFee !== null && data.CourseFee !== '') ? String(data.CourseFee) : '0.00';
    rowObj['ApprovalStatus'] = data.ApprovalStatus || 'Approved';
    rowObj['RequestedBy'] = requesterSigData.employeeNo || '';
    rowObj['RequestedByName'] = requesterSigData.name || '';
    rowObj['RequestedByEmail'] = requesterSigData.email || '';
    rowObj['RequestedDate'] = timeNow;

    const newRowValues = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
    sheet.appendRow(newRowValues);

    SpreadsheetApp.flush();

    if (requisitionForm.fileId) {
      try {
        updateTrainingRequisitionSignatures(id, 'request', requesterSigData, requisitionForm.fileId);
        updateTrainingRequisitionSignatures(id, 'HOD', { status: 'Approved', employeeNo: requesterSigData.employeeNo, name: requesterSigData.name, position: requesterSigData.position, date: timeNow }, requisitionForm.fileId);
      } catch(e) {}
    }

    if (participantsList.length > 0) {
      try {
        addTrainingParticipants(id, participantsList);
      } catch (pErr) {
        Logger.log('Error adding participants in addTraining: ' + pErr.message);
      }
    }

    try { syncTrainingById(id, 'Admin Programme Creation', 'CREATE'); } catch(syncErr) { Logger.log('syncTrainingById error in addTraining: ' + syncErr.message); }

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

// --- Update ---------------------------------------------------------------------
function updateTraining(data) {
  try {
    if (!data.ID) return err('Training ID is required.');

    const sheet = getSheet(SHEET_NAMES.trainings);
    const headers = ensureTrainingSheetColumns(sheet);
    const row = findRowById(sheet, data.ID);
    if (row === -1) return err('Training not found.');

    const dataRange = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
    const rowObj = {};
    headers.forEach((h, idx) => {
      rowObj[h] = dataRange[idx] !== undefined ? dataRange[idx] : '';
    });

    if (data.Code !== undefined) rowObj['Code'] = data.Code;
    if (data.Name !== undefined) rowObj['Name'] = data.Name;
    if (data.Category !== undefined) rowObj['Category'] = data.Category;
    if (data.TrainingMode !== undefined) rowObj['TrainingMode'] = data.TrainingMode;
    if (data.TnaSource !== undefined) rowObj['TnaSource'] = data.TnaSource;
    if (data.Trainer !== undefined) rowObj['Trainer'] = data.Trainer;
    if (data.Venue !== undefined) rowObj['Venue'] = data.Venue;
    if (data.TrainingProvider !== undefined) rowObj['TrainingProvider'] = data.TrainingProvider;
    if (data.StartDate !== undefined) rowObj['StartDate'] = data.StartDate;
    if (data.EndDate !== undefined) rowObj['EndDate'] = data.EndDate;
    if (data.Duration !== undefined) rowObj['Duration'] = data.Duration;
    if (data.TotalHours !== undefined) rowObj['TotalHours'] = data.TotalHours;
    if (data.Department !== undefined) rowObj['Department'] = data.Department;
    if (data.ExpiryDate !== undefined || data.CertExpiryDate !== undefined) {
      const expVal = data.ExpiryDate || data.CertExpiryDate || '';
      rowObj['ExpiryDate'] = expVal;
      rowObj['CertExpiryDate'] = expVal;
    }
    if (data.RequisitionUrl !== undefined) rowObj['RequisitionUrl'] = data.RequisitionUrl;
    if (data.Reason !== undefined) rowObj['Reason'] = data.Reason;
    if (data.Objectives !== undefined) rowObj['Objectives'] = data.Objectives;
    if (data.Status !== undefined) rowObj['Status'] = data.Status;
    if (data.Stage !== undefined) rowObj['Stage'] = data.Stage;
    if (data.Participants !== undefined) rowObj['Participants'] = data.Participants;
    if (data.CourseFee !== undefined && data.CourseFee !== null && data.CourseFee !== '') {
      rowObj['CourseFee'] = String(data.CourseFee);
    }
    if (data.ApprovalStatus !== undefined) rowObj['ApprovalStatus'] = data.ApprovalStatus;

    rowObj['UpdatedDate'] = now();

    const updatedRowValues = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
    sheet.getRange(row, 1, 1, headers.length).setValues([updatedRowValues]);

    const appStatusCol = headers.indexOf('ApprovalStatus') + 1;
    if (appStatusCol && data.ApprovalStatus !== undefined) {
      sheet.getRange(row, appStatusCol).setValue(data.ApprovalStatus);
      if (data.RequisitionStep || data.HREmployeeNo || data.HRName) {
        try {
          updateTrainingRequisitionSignatures(data.ID, data.RequisitionStep || 'HR', {
            status: data.ApprovalStatus,
            employeeNo: data.EmployeeNo || data.HREmployeeNo || '',
            name: data.HRName || data.EmployeeName || 'HR Department',
            position: data.HRPosition || 'HR Department',
            date: now()
          });
        } catch(sigErr) {
          Logger.log('Signature update error in updateTraining: ' + sigErr.message);
        }
      }
    }

    try { syncTrainingById(data.ID, 'Admin Programme Update', 'UPDATE'); } catch(syncErr) { Logger.log('syncTrainingById error in updateTraining: ' + syncErr.message); }

    return ok({ message: 'Training updated successfully.' });
  } catch (e) {
    return err('Failed to update training: ' + e.message);
  }
}

/**
 * Explicit HR Department Acknowledgment for a Training Requisition Form
 */
function acknowledgeHRRequisition(trainingId, hrData) {
  try {
    if (!trainingId) return err('Training ID is required.');
    const sheet = getSheet(SHEET_NAMES.trainings);
    const headers = ensureTrainingSheetColumns(sheet);
    const row = findRowById(sheet, trainingId);
    if (row === -1) return err('Training not found.');

    let data = hrData || {};

    // If missing explicit name/employeeNo or using hardcoded defaults, resolve from active user profile
    if (!data.name || !data.employeeNo || data.name === 'Arina Ismail' || data.employeeNo === '1012') {
      let activeEmail = '';
      try {
        activeEmail = Session.getActiveUser().getEmail();
      } catch (e) {}
      const resolvedProfile = getHrProfileByEmail(activeEmail);
      data = {
        employeeNo: data.employeeNo && data.employeeNo !== '1012' ? data.employeeNo : resolvedProfile.employeeNo,
        name: data.name && data.name !== 'Arina Ismail' ? data.name : resolvedProfile.name,
        position: data.position || resolvedProfile.position,
        status: data.status || 'Acknowledged',
        date: data.date || new Date()
      };
    }

    const empNo = data.employeeNo || data.EmployeeNo || data.ID || 'HR-001';
    const empName = data.name || data.Name || data.EmployeeName || 'HR Department';
    const position = data.position || data.Position || data.JobPosition || 'HR Department';
    const status = data.status || data.ApprovalStatus || 'Acknowledged';
    const timeStamp = getFormattedCurrentDate(data.date || new Date());

    updateTrainingRequisitionSignatures(trainingId, 'HR Department', {
      status: status,
      employeeNo: empNo,
      name: empName,
      position: position,
      date: timeStamp
    });

    return ok({ message: `Requisition form acknowledged by HR Department (${empName}).` });
  } catch (e) {
    return err('Failed to record HR acknowledgment: ' + e.message);
  }
}

// --- Update Stage Only ----------------------------------------------------------
function updateTrainingStage(trainingId, newStage) {
  try {
    if (!LIFECYCLE_STAGES.includes(newStage))
      return err('Invalid lifecycle stage.');

    const sheet = getSheet(SHEET_NAMES.trainings);
    if (!sheet) return err(`Sheet "${SHEET_NAMES.trainings}" not found.`);
    const row = findRowById(sheet, trainingId);
    if (row === -1) return err('Training not found.');

    const headers = ensureTrainingSheetColumns(sheet);
    const stageCol = headers.indexOf('Stage') + 1;
    const updatedDateCol = headers.indexOf('UpdatedDate') + 1;

    if (stageCol > 0) sheet.getRange(row, stageCol).setValue(newStage);
    if (updatedDateCol > 0) sheet.getRange(row, updatedDateCol).setValue(now());
    try { syncTrainingById(trainingId, 'Stage Change (' + newStage + ')', 'STATUS_CHANGE'); } catch(sErr) {}
    return ok({ message: 'Stage updated to: ' + newStage });
  } catch (e) {
    return err(e.message);
  }
}

// --- Delete ---------------------------------------------------------------------
function deleteTraining(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.trainings);
    const row = findRowById(sheet, id);
    if (row === -1) return err('Training not found.');

    const headers = ensureTrainingSheetColumns(sheet);
    const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
    const folderId = values[headers.indexOf('FolderID')];
    const trainingCode = values[headers.indexOf('Code')] || id;
    const trainingName = values[headers.indexOf('Name')] || '';
    const startDate = values[headers.indexOf('StartDate')];
    let year = '2026';
    if (startDate) {
      const d = parseDateObj(startDate);
      if (d) year = String(d.getFullYear());
    }

    // 1. Move the training's dedicated Drive folder/workspace to trash
    if (folderId) {
      try {
        DriveApp.getFolderById(String(folderId).trim()).setTrashed(true);
      } catch (driveErr) {
        Logger.log('Drive folder trash error (continuing deletion): ' + driveErr.message);
      }
    }

    // 2. Remove orphaned rows from central database sheets if any exist
    cleanCentralDatabaseForTraining(id, trainingCode);

    // 3. Delete the row from the master Trainings sheet
    sheet.deleteRow(row);

    // 4. Update and remove from Master Annual Training Report spreadsheets
    try {
      if (typeof removeTrainingFromAllReports === 'function') {
        removeTrainingFromAllReports(id, trainingName, trainingCode, year);
      }
    } catch(syncErr) {
      Logger.log('Master report sync on delete warning: ' + syncErr.message);
    }

    return ok({ message: 'Training, database records, and reports were successfully updated.' });
  } catch (e) {
    return err('Failed to delete training: ' + e.message);
  }
}

/**
 * Cleans up any rows associated with the deleted training across central database sheets.
 */
function cleanCentralDatabaseForTraining(trainingId, trainingCode) {
  const cleanTid = String(trainingId || '').trim().toLowerCase();
  const cleanCode = String(trainingCode || '').trim().toLowerCase();
  if (!cleanTid && !cleanCode) return;

  const targetSheetNames = [
    SHEET_NAMES.attendance,
    SHEET_NAMES.trainingSessions,
    SHEET_NAMES.trainingEval,
    SHEET_NAMES.postEval,
    SHEET_NAMES.trainingParticipants
  ];

  targetSheetNames.forEach(sheetName => {
    try {
      const sh = getSheet(sheetName);
      if (!sh) return;
      const data = sh.getDataRange().getValues();
      if (data.length <= 1) return;

      const headers = data[0].map(h => String(h || '').trim().toLowerCase());
      const tidCol = headers.findIndex(h => h === 'trainingid' || h === 'training_id' || h === 'training id' || h === 'training code' || h === 'trainingcode');

      if (tidCol === -1) return;

      for (let r = data.length - 1; r >= 1; r--) {
        const val = String(data[r][tidCol] || '').trim().toLowerCase();
        if ((cleanTid && val === cleanTid) || (cleanCode && val === cleanCode)) {
          sh.deleteRow(r + 1);
        }
      }
    } catch (e) {
      Logger.log(`Error cleaning sheet ${sheetName} for training ${trainingId}: ${e.message}`);
    }
  });
}

// --- Dashboard Summary ----------------------------------------------------------
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
      pendingPost: rows.filter(r => r.Stage === 'Waiting for 3-Month Review' || r.Stage === 'Waiting for 6-Month Review').length,
    });
  } catch (e) {
    return err(e.message);
  }
}

// --- Helpers --------------------------------------------------------------------
function generateTrainingCode(category) {
  const prefix = {
    'Behavioral Skills':       'BS',
    'Technical Skills':        'TS',
    'Compliance Training':     'CT',
    'Business Skills':         'BUS',
    'Onboarding':              'ONB',
    'Leadership & Management': 'LM',
    'Compliance & Regulatory': 'CR',
    'Technology & Innovation': 'TI',
    'Project Management':      'PM',
    'Data & Analytics':        'DA',
    'Customer Experience':     'CX'
  }[category] || 'TR';
  return prefix + '-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-4);
}

/**
 * Diagnostic function to verify central database sheet connectivity, headers, and record counts.
 */
function debugGetTrainings() {
  try {
    const ss = getSpreadsheet();
    const sheetName = SHEET_NAMES.trainings;
    const sheet = getSheet(sheetName);
    if (!sheet) {
      return {
        success: false,
        error: `Sheet "${sheetName}" not found.`,
        spreadsheetName: ss ? ss.getName() : 'No spreadsheet found',
        spreadsheetId: ss ? ss.getId() : 'None'
      };
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = lastRow > 0 && lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const rows = sheetToJson(sheet);
    return {
      success: true,
      sheetNameConfigured: sheetName,
      actualSheetName: sheet.getName(),
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      lastRow: lastRow,
      lastColumn: lastCol,
      headers: headers,
      recordsCount: rows.length,
      firstRecord: rows.length > 0 ? rows[0] : null
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      stack: e.stack
    };
  }
}

// --- Training Participants Management ------------------------------------------
function getTrainingParticipants(trainingId) {
  try {
    if (!trainingId) return ok([]);
    const list = getTrainingParticipantsList(trainingId);
    return ok(list);
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

    const resolution = canonicalizeTrainingParticipants(participants);
    if (resolution.rejected.length > 0) {
      return err('The following participant(s) do not match an Employee-sheet record: ' + resolution.rejected.join(', '));
    }

    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return err('Could not access training data spreadsheet for ID: ' + trainingId);

    let sheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
    if (!sheet) {
      sheet = ss.insertSheet('Participants');
      sheet.appendRow(['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt', 'SupervisorID', 'SupervisorEmail', 'SupervisorName']);
      sheet.getRange('A1:J1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const existingRows = sheetToJson(sheet);
    const existingEmpIds = new Set(existingRows.map(r => String(r.EmployeeID || r.ID || '').trim().toLowerCase()));

    let addedCount = 0;
    const addedAt = now();

    resolution.participants.forEach(dbEmp => {
      const empIdLower = dbEmp.ID.toLowerCase();
      if (!existingEmpIds.has(empIdLower)) {
        const empId   = dbEmp.ID;
        const empName = dbEmp.Name;
        const empDept = dbEmp.Department;
        const empPos  = dbEmp.Position;

        sheet.appendRow([
          generateId('TP'),
          trainingId,
          empId,
          empName,
          empDept,
          empPos,
          addedAt
        ]);
        existingEmpIds.add(empIdLower);
        addedCount++;
      }
    });

    const totalCount = existingEmpIds.size;
    updateTrainingParticipantCount(trainingId, totalCount);

    try { syncTrainingRequisitionParticipants(trainingId); } catch(e) {}

    try {
      const tSheet = getSheet(SHEET_NAMES.trainings);
      if (tSheet) {
        const tRows = sheetToJson(tSheet);
        const parentT = tRows.find(r => r.ID === trainingId);
        if (parentT && (!parentT.Stage || parentT.Stage === 'Created')) {
          updateTrainingStage(trainingId, 'Participants Imported');
        }
      }
    } catch(sErr) {
      Logger.log('Auto update stage error in addTrainingParticipants: ' + sErr.message);
    }

    return ok({ message: `Added ${addedCount} participants successfully.`, count: totalCount });
  } catch (e) {
    return err('Failed to add participants: ' + e.message);
  }
}

function removeTrainingParticipant(trainingId, employeeId) {
  try {
    if (!trainingId || !employeeId) return err('Training ID and Employee ID are required.');
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return err('Could not access training data spreadsheet.');
    const sheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
    if (!sheet) return err('Participant record not found.');

    const data = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      const rowEmpId = String(data[i][2] || data[i][0] || '').trim().toLowerCase();
      if (rowEmpId === String(employeeId).trim().toLowerCase()) {
        sheet.deleteRow(i + 1);
        foundRow = i;
        break;
      }
    }

    if (foundRow === -1) return err('Participant record not found.');

    const remainingRows = sheetToJson(sheet);
    const totalCount = remainingRows.length;
    updateTrainingParticipantCount(trainingId, totalCount);

    try { syncTrainingRequisitionParticipants(trainingId); } catch(e) {}

    return ok({ message: 'Participant removed successfully.', count: totalCount });
  } catch (e) {
    return err('Failed to remove participant: ' + e.message);
  }
}

function updateTrainingParticipantCount(trainingId, count) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return;
    const row = findRowById(tSheet, trainingId);
    if (row !== -1) {
      const headers = ensureTrainingSheetColumns(tSheet);
      const partCol = headers.indexOf('Participants') + 1;
      const stageCol = headers.indexOf('Stage') + 1;
      const updatedDateCol = headers.indexOf('UpdatedDate') + 1;

      if (partCol > 0) tSheet.getRange(row, partCol).setValue(count);
      if (stageCol > 0) {
        const currentStage = tSheet.getRange(row, stageCol).getValue();
        if ((!currentStage || currentStage === 'Created') && count > 0) {
          tSheet.getRange(row, stageCol).setValue('Participants Imported');
        }
      }
      if (updatedDateCol > 0) tSheet.getRange(row, updatedDateCol).setValue(now());
    }
  } catch (e) {
    Logger.log('updateTrainingParticipantCount error: ' + e.message);
  }
}

/**
 * Returns actionable reminders and notifications for upcoming training evaluations,
 * supervisor assignments, and 3-month post-evaluations.
 * Purely read-only calculation without mutating or adding any database sheets.
 */
function getTrainingActionNotifications() {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return ok({ count: 0, notifications: [] });

    const tRows = sheetToJson(tSheet);
    if (!tRows || tRows.length === 0) return ok({ count: 0, notifications: [] });

    const now = new Date();
    const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    let publicBaseUrl = '';
    try {
      const pUrl = getPublicPortalUrl();
      if (pUrl) publicBaseUrl = String(pUrl).split('?')[0];
    } catch(e) {}

    const notifications = [];

    tRows.forEach(t => {
      const trainingId = String(t.ID || t.Code || '').trim();
      if (!trainingId) return;

      const trainingTitle = String(t.Name || 'Untitled Programme').trim();
      const trainingCode  = String(t.Code || trainingId).trim();
      const status        = String(t.Status || '').trim();
      const stage         = String(t.Stage || '').trim();
      const approvalStatus = String(t.ApprovalStatus || 'Approved').trim();

      const startDateStr = t.StartDate;
      const endDateStr   = t.EndDate || t.StartDate;
      if (!startDateStr && !endDateStr) return;

      const compDateRaw = endDateStr || startDateStr;
      let compDate = null;
      if (compDateRaw instanceof Date) {
        compDate = !isNaN(compDateRaw.getTime()) ? new Date(compDateRaw.getFullYear(), compDateRaw.getMonth(), compDateRaw.getDate()) : null;
      } else {
        const str = String(compDateRaw).trim();
        const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
        if (ymd) {
          compDate = new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10));
        } else {
          const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
          if (dmy) {
            compDate = new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
          } else {
            const d = new Date(str);
            if (!isNaN(d.getTime())) compDate = new Date(d.getFullYear(), d.getMonth(), d.getDate());
          }
        }
      }
      if (!compDate) return;

      // Determine if training has completed
      const isCompleted = ['Training Completed', 'Evaluation Completed', 'Waiting for 3-Month Review', 'Programme Closed', 'Completed'].includes(stage) ||
                          status === 'Completed' ||
                          (todayMidnight.getTime() > compDate.getTime());

      if (!isCompleted) {
        return;
      }

      const evalPublicUrl = publicBaseUrl ? `${publicBaseUrl}?page=evaluation&id=${encodeURIComponent(trainingId)}` : '';
      const postPublicUrl = publicBaseUrl ? `${publicBaseUrl}?page=post&id=${encodeURIComponent(trainingId)}` : '';

      // -------------------------------------------------------------------------
      // 1. Participant Training Evaluation (2-Week Window)
      // -------------------------------------------------------------------------
      const isEvalCompleted = ['Evaluation Completed', 'Waiting for 3-Month Review', 'Programme Closed'].includes(stage);

      if (!isEvalCompleted) {
        // Calculate 2-Week evaluation deadline (14 days after completion date)
        const evalDeadline = new Date(compDate.getFullYear(), compDate.getMonth(), compDate.getDate() + 14, 23, 59, 59, 999);
        const evalDeadlineMidnight = new Date(compDate.getFullYear(), compDate.getMonth(), compDate.getDate() + 14);
        const daysRemaining = Math.round((evalDeadlineMidnight.getTime() - todayMidnight.getTime()) / (1000 * 60 * 60 * 24));
        const deadlineFormatted = formatMinimalistDate(evalDeadline);

        if (daysRemaining < 0) {
          // Overdue
          const overdueDays = Math.abs(daysRemaining);
          notifications.push({
            id: `${trainingId}_EVALUATION`,
            category: 'EVALUATION',
            type: 'EVALUATION_OVERDUE',
            priority: 'OVERDUE',
            priorityRank: 1,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: `Training Evaluation overdue by ${overdueDays} day${overdueDays > 1 ? 's' : ''}`,
            message: `Participant evaluation deadline passed on ${deadlineFormatted}. Review completed submissions.`,
            countdownText: `Overdue by ${overdueDays} day${overdueDays > 1 ? 's' : ''}`,
            daysRemaining: daysRemaining,
            deadlineDate: evalDeadline.toISOString(),
            deadlineFormatted: deadlineFormatted,
            action: 'VIEW_EVALUATION',
            actionLabel: 'View Evaluation',
            evalUrl: evalPublicUrl
          });
        } else if (daysRemaining === 0) {
          // Due Today
          notifications.push({
            id: `${trainingId}_EVALUATION`,
            category: 'EVALUATION',
            type: 'EVALUATION_DUE_TODAY',
            priority: 'CRITICAL',
            priorityRank: 2,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: 'Training Evaluation due today',
            message: 'Participants have until the end of today to complete the 2-week training evaluation.',
            countdownText: 'Due today',
            daysRemaining: 0,
            deadlineDate: evalDeadline.toISOString(),
            deadlineFormatted: deadlineFormatted,
            action: 'SHARE_EVALUATION_QR',
            actionLabel: 'Share Evaluation QR',
            secondaryAction: 'VIEW_EVALUATION',
            secondaryActionLabel: 'View Evaluation',
            evalUrl: evalPublicUrl
          });
        } else if (daysRemaining === 1) {
          // 1 Day Remaining (Tomorrow)
          notifications.push({
            id: `${trainingId}_EVALUATION`,
            category: 'EVALUATION',
            type: 'EVALUATION_1_DAY',
            priority: 'CRITICAL',
            priorityRank: 2,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: 'Tomorrow: Training Evaluation deadline',
            message: 'Participant evaluation deadline is tomorrow. Ensure attendees have received the evaluation QR.',
            countdownText: 'Tomorrow',
            daysRemaining: 1,
            deadlineDate: evalDeadline.toISOString(),
            deadlineFormatted: deadlineFormatted,
            action: 'SHARE_EVALUATION_QR',
            actionLabel: 'Share Evaluation QR',
            evalUrl: evalPublicUrl
          });
        } else if (daysRemaining <= 3) {
          // 3 Days Remaining
          notifications.push({
            id: `${trainingId}_EVALUATION`,
            category: 'EVALUATION',
            type: 'EVALUATION_3_DAYS',
            priority: 'CRITICAL',
            priorityRank: 2,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: `${daysRemaining} days left for Training Evaluation`,
            message: `Participant evaluation deadline is in ${daysRemaining} days. Consider sharing the Evaluation QR.`,
            countdownText: `${daysRemaining} days remaining`,
            daysRemaining: daysRemaining,
            deadlineDate: evalDeadline.toISOString(),
            deadlineFormatted: deadlineFormatted,
            action: 'SHARE_EVALUATION_QR',
            actionLabel: 'Share Evaluation QR',
            evalUrl: evalPublicUrl
          });
        } else if (daysRemaining <= 7) {
          // 7 Days Remaining (1 week)
          notifications.push({
            id: `${trainingId}_EVALUATION`,
            category: 'EVALUATION',
            type: 'EVALUATION_1_WEEK',
            priority: 'HIGH',
            priorityRank: 3,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: '1 week left for Training Evaluation',
            message: `Participants have ${daysRemaining} days remaining to complete their evaluation. Consider sharing the Evaluation QR.`,
            countdownText: `${daysRemaining} days remaining`,
            daysRemaining: daysRemaining,
            deadlineDate: evalDeadline.toISOString(),
            deadlineFormatted: deadlineFormatted,
            action: 'SHARE_EVALUATION_QR',
            actionLabel: 'Share Evaluation QR',
            evalUrl: evalPublicUrl
          });
        }
      }

      // -------------------------------------------------------------------------
      // 2. Supervisor Assignment (Approaching 3-Month Milestone)
      // -------------------------------------------------------------------------
      const targetY = compDate.getFullYear() + Math.floor((compDate.getMonth() + 3) / 12);
      const targetM = (compDate.getMonth() + 3) % 12;
      const daysInTargetMonth = new Date(targetY, targetM + 1, 0).getDate();
      const targetDay = Math.min(compDate.getDate(), daysInTargetMonth);
      const threeMonthDate = new Date(targetY, targetM, targetDay, 23, 59, 59, 999);
      const threeMonthMidnight = new Date(targetY, targetM, targetDay);
      const daysUntil3Mo = Math.round((threeMonthMidnight.getTime() - todayMidnight.getTime()) / (1000 * 60 * 60 * 24));
      const milestoneFormatted = formatMinimalistDate(threeMonthDate);

      const isClosed = stage === 'Programme Closed';

      if (!isClosed && daysUntil3Mo <= 14) {
        let participants = [];
        try {
          participants = getTrainingParticipantsList(trainingId);
        } catch(pErr) {}

        const unassignedParticipants = participants.filter(p => {
          const supId = String(p.SupervisorID || '').trim();
          const supEmail = String(p.SupervisorEmail || '').trim();
          const supName = String(p.SupervisorName || '').trim();
          return !supId && !supEmail && !supName;
        });

        const unassignedCount = unassignedParticipants.length;

        if (participants.length > 0 && unassignedCount > 0) {
          if (daysUntil3Mo < 0) {
            // Overdue
            const overdueDays = Math.abs(daysUntil3Mo);
            notifications.push({
              id: `${trainingId}_SUPERVISOR_ASSIGNMENT`,
              category: 'SUPERVISOR',
              type: 'SUPERVISOR_ASSIGNMENT_OVERDUE',
              priority: 'OVERDUE',
              priorityRank: 1,
              trainingId: trainingId,
              trainingTitle: trainingTitle,
              trainingCode: trainingCode,
              status: status,
              stage: stage,
              title: `Supervisor assignment overdue by ${overdueDays} day${overdueDays > 1 ? 's' : ''}`,
              message: `${unassignedCount} participant${unassignedCount > 1 ? 's' : ''} still require supervisor assignment. 3-Month post evaluation is now active.`,
              countdownText: `Overdue by ${overdueDays} day${overdueDays > 1 ? 's' : ''}`,
              daysRemaining: daysUntil3Mo,
              deadlineDate: threeMonthDate.toISOString(),
              deadlineFormatted: milestoneFormatted,
              unassignedCount: unassignedCount,
              action: 'ASSIGN_SUPERVISOR',
              actionLabel: 'Assign Supervisor'
            });
          } else if (daysUntil3Mo === 0) {
            // Due Today
            notifications.push({
              id: `${trainingId}_SUPERVISOR_ASSIGNMENT`,
              category: 'SUPERVISOR',
              type: 'SUPERVISOR_ASSIGNMENT_DUE_TODAY',
              priority: 'CRITICAL',
              priorityRank: 2,
              trainingId: trainingId,
              trainingTitle: trainingTitle,
              trainingCode: trainingCode,
              status: status,
              stage: stage,
              title: 'Supervisor assignment due today',
              message: `${unassignedCount} participant${unassignedCount > 1 ? 's' : ''} require supervisor assignment before 3-Month Post Evaluation unlocks today.`,
              countdownText: 'Due today',
              daysRemaining: 0,
              deadlineDate: threeMonthDate.toISOString(),
              deadlineFormatted: milestoneFormatted,
              unassignedCount: unassignedCount,
              action: 'ASSIGN_SUPERVISOR',
              actionLabel: 'Assign Supervisor'
            });
          } else if (daysUntil3Mo <= 3) {
            // 3 Days Remaining
            notifications.push({
              id: `${trainingId}_SUPERVISOR_ASSIGNMENT`,
              category: 'SUPERVISOR',
              type: 'SUPERVISOR_ASSIGNMENT_3_DAYS',
              priority: 'CRITICAL',
              priorityRank: 2,
              trainingId: trainingId,
              trainingTitle: trainingTitle,
              trainingCode: trainingCode,
              status: status,
              stage: stage,
              title: '3 days left to assign Supervisor',
              message: `3-Month Post Evaluation milestone is in ${daysUntil3Mo} days. ${unassignedCount} participant${unassignedCount > 1 ? 's' : ''} still require a supervisor.`,
              countdownText: `${daysUntil3Mo} days remaining`,
              daysRemaining: daysUntil3Mo,
              deadlineDate: threeMonthDate.toISOString(),
              deadlineFormatted: milestoneFormatted,
              unassignedCount: unassignedCount,
              action: 'ASSIGN_SUPERVISOR',
              actionLabel: 'Assign Supervisor'
            });
          } else if (daysUntil3Mo <= 7) {
            // 1 Week Remaining
            notifications.push({
              id: `${trainingId}_SUPERVISOR_ASSIGNMENT`,
              category: 'SUPERVISOR',
              type: 'SUPERVISOR_ASSIGNMENT_1_WEEK',
              priority: 'HIGH',
              priorityRank: 3,
              trainingId: trainingId,
              trainingTitle: trainingTitle,
              trainingCode: trainingCode,
              status: status,
              stage: stage,
              title: '1 week left to assign Supervisor',
              message: `Post Evaluation will begin on ${milestoneFormatted}. Please assign supervisors for ${unassignedCount} participant${unassignedCount > 1 ? 's' : ''}.`,
              countdownText: `${daysUntil3Mo} days remaining`,
              daysRemaining: daysUntil3Mo,
              deadlineDate: threeMonthDate.toISOString(),
              deadlineFormatted: milestoneFormatted,
              unassignedCount: unassignedCount,
              action: 'ASSIGN_SUPERVISOR',
              actionLabel: 'Assign Supervisor'
            });
          } else if (daysUntil3Mo <= 14) {
            // 2 Weeks Remaining
            notifications.push({
              id: `${trainingId}_SUPERVISOR_ASSIGNMENT`,
              category: 'SUPERVISOR',
              type: 'SUPERVISOR_ASSIGNMENT_2_WEEKS',
              priority: 'MEDIUM',
              priorityRank: 4,
              trainingId: trainingId,
              trainingTitle: trainingTitle,
              trainingCode: trainingCode,
              status: status,
              stage: stage,
              title: '2 weeks left to assign Supervisor',
              message: `Post Evaluation begins on ${milestoneFormatted}. ${unassignedCount} participant${unassignedCount > 1 ? 's' : ''} currently have no assigned supervisor.`,
              countdownText: `${daysUntil3Mo} days remaining`,
              daysRemaining: daysUntil3Mo,
              deadlineDate: threeMonthDate.toISOString(),
              deadlineFormatted: milestoneFormatted,
              unassignedCount: unassignedCount,
              action: 'ASSIGN_SUPERVISOR',
              actionLabel: 'Assign Supervisor'
            });
          }
        }
      }

      // -------------------------------------------------------------------------
      // 3. Three-Month Post Evaluation (Active / Due Milestone & Email Status)
      // -------------------------------------------------------------------------
      if (!isClosed && (daysUntil3Mo <= 0 || stage === 'Waiting for 3-Month Review')) {
        let isPostDone = false;
        try {
          const ss = getTrainingDataSpreadsheet(trainingId);
          if (ss) {
            const postSheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
            const tpSheet   = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
            if (postSheet && tpSheet) {
              const postRows = sheetToJson(postSheet);
              const tpRows   = sheetToJson(tpSheet);
              if (tpRows.length > 0 && postRows.length >= tpRows.length) {
                isPostDone = true;
              }
            }
          }
        } catch(postErr) {}

        let participants = [];
        try {
          participants = getTrainingParticipantsList(trainingId);
        } catch(pErr) {}

        const unassignedParticipants = participants.filter(p => {
          const supId = String(p.SupervisorID || '').trim();
          const supEmail = String(p.SupervisorEmail || '').trim();
          const supName = String(p.SupervisorName || '').trim();
          return !supId && !supEmail && !supName;
        });

        const unassignedCount = unassignedParticipants.length;
        const emailStatus = String(t.PostEvalEmailStatus || '').trim().toUpperCase();

        if (unassignedCount > 0) {
          // Critical: Supervisor Required Before Post Evaluation Email
          notifications.push({
            id: `${trainingId}_POST_EVAL_EMAIL_SUPERVISOR_MISSING`,
            category: 'POST_EVALUATION',
            type: 'SUPERVISOR_REQUIRED_BEFORE_EMAIL',
            priority: 'CRITICAL',
            priorityRank: 1,
            isActionRequired: true,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: 'Supervisor Required Before Post Evaluation Email',
            message: `Post Evaluation email cannot be sent for ${unassignedCount} participant${unassignedCount > 1 ? 's' : ''} because supervisors have not been assigned.`,
            countdownText: '3-Month Milestone Reached',
            daysRemaining: 0,
            deadlineDate: threeMonthDate.toISOString(),
            deadlineFormatted: milestoneFormatted,
            unassignedCount: unassignedCount,
            action: 'ASSIGN_SUPERVISOR',
            actionLabel: 'Assign Supervisor'
          });
        } else if (emailStatus === 'FAILED') {
          // Critical: Post Evaluation Email Failed
          notifications.push({
            id: `${trainingId}_POST_EVAL_EMAIL_FAILED`,
            category: 'POST_EVALUATION',
            type: 'POST_EVALUATION_EMAIL_FAILED',
            priority: 'CRITICAL',
            priorityRank: 1,
            isActionRequired: true,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: 'Post Evaluation Email Failed',
            message: `The supervisor Post Evaluation email could not be sent${t.PostEvalEmailError ? ': ' + t.PostEvalEmailError : '.'}`,
            countdownText: 'Email Failed',
            daysRemaining: 0,
            deadlineDate: threeMonthDate.toISOString(),
            deadlineFormatted: milestoneFormatted,
            action: 'RETRY_POST_EVAL_EMAIL',
            actionLabel: 'Retry Email',
            errorDetails: t.PostEvalEmailError || ''
          });
        } else if (emailStatus === 'SENT') {
          // Informational: Post Evaluation Email Sent
          notifications.push({
            id: `${trainingId}_POST_EVAL_EMAIL_SENT`,
            category: 'POST_EVALUATION',
            type: 'POST_EVALUATION_EMAIL_SENT',
            priority: 'INFORMATIONAL',
            priorityRank: 5,
            isActionRequired: false,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: 'Post Evaluation Email Sent',
            message: `The 3-month Post Evaluation email was successfully sent to assigned supervisor(s)${t.PostEvalEmailSentAt ? ' on ' + t.PostEvalEmailSentAt : ''}.`,
            countdownText: 'Email Sent',
            daysRemaining: 0,
            deadlineDate: threeMonthDate.toISOString(),
            deadlineFormatted: milestoneFormatted,
            emailSentAt: t.PostEvalEmailSentAt || '',
            action: 'VIEW_POST_EVALUATION',
            actionLabel: 'View Post Evaluation',
            postUrl: postPublicUrl
          });
        } else {
          // High: Post Evaluation Email Pending
          notifications.push({
            id: `${trainingId}_POST_EVAL_EMAIL_PENDING`,
            category: 'POST_EVALUATION',
            type: 'POST_EVALUATION_EMAIL_PENDING',
            priority: 'HIGH',
            priorityRank: 2,
            isActionRequired: true,
            trainingId: trainingId,
            trainingTitle: trainingTitle,
            trainingCode: trainingCode,
            status: status,
            stage: stage,
            title: 'Post Evaluation Email Pending',
            message: 'The 3-month milestone has been reached. System is ready to send the supervisor evaluation email.',
            countdownText: '3-Month Milestone Reached',
            daysRemaining: 0,
            deadlineDate: threeMonthDate.toISOString(),
            deadlineFormatted: milestoneFormatted,
            action: 'SEND_POST_EVAL_EMAIL',
            actionLabel: 'Send Email Now',
            postUrl: postPublicUrl
          });
        }
      }
    });

    // Sort notifications: Priority rank (1 to 5), then earliest deadline, then training title
    notifications.sort((a, b) => {
      if (a.priorityRank !== b.priorityRank) return a.priorityRank - b.priorityRank;
      const dateA = a.deadlineDate ? new Date(a.deadlineDate).getTime() : 0;
      const dateB = b.deadlineDate ? new Date(b.deadlineDate).getTime() : 0;
      if (dateA !== dateB) return dateA - dateB;
      return String(a.trainingTitle).localeCompare(String(b.trainingTitle));
    });

    const actionCount = notifications.filter(n => n.isActionRequired !== false).length;

    return ok({
      count: actionCount,
      totalCount: notifications.length,
      notifications: notifications
    });

  } catch (e) {
    Logger.log('getTrainingActionNotifications error: ' + e.message);
    return err('Failed to calculate training notifications: ' + e.message);
  }
}
