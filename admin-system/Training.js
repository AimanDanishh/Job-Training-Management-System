/**
 * Training.gs — Training Programme management
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
    if (!sheet) return [];
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
        const defaultFee = (t.ID === 'TRN-101' || t.Code === 'TRN-101') ? '1500.00'
                         : (t.ID === 'TRN-102' || t.Code === 'TRN-102') ? '2800.00'
                         : (t.ID === 'TRN-103' || t.Code === 'TRN-103') ? '1200.00'
                         : (t.ID === 'TRN-104' || t.Code === 'TRN-104') ? '2200.00'
                         : '1000.00';
        t.CourseFee = defaultFee;
        if (t._row && courseFeeCol > 0) {
          sheet.getRange(t._row, courseFeeCol).setValue(defaultFee);
          sheetModified = true;
        }
      }

      const partCount = Number(t.Participants || 0);

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
      SpreadsheetApp.flush();
      rows.forEach(t => {
        if (t._isLifecycleUpdated && t.ID) {
          try { syncTrainingById(t.ID, 'System Lifecycle Auto-Advance', 'STATUS_CHANGE'); } catch(sErr) {}
        }
      });
    }
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

// ─── Update ─────────────────────────────────────────────────────────────────────
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
    if (data.Trainer !== undefined) rowObj['Trainer'] = data.Trainer;
    if (data.Venue !== undefined) rowObj['Venue'] = data.Venue;
    if (data.StartDate !== undefined) rowObj['StartDate'] = data.StartDate;
    if (data.EndDate !== undefined) rowObj['EndDate'] = data.EndDate;
    if (data.Duration !== undefined) rowObj['Duration'] = data.Duration;
    if (data.TotalHours !== undefined) rowObj['TotalHours'] = data.TotalHours;
    if (data.Department !== undefined) rowObj['Department'] = data.Department;
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
    try { syncTrainingById(trainingId, 'Stage Change (' + newStage + ')', 'STATUS_CHANGE'); } catch(sErr) {}
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
      pendingPost: rows.filter(r => r.Stage === 'Waiting for 3-Month Review' || r.Stage === 'Waiting for 6-Month Review').length,
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
    if (!trainingId) return ok([]);
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return ok([]);
    const sheet = ss.getSheetByName('TrainingParticipants');
    if (!sheet) return ok([]);
    const rows = sheetToJson(sheet);
    const resolution = canonicalizeTrainingParticipants(rows);
    return ok(resolution.participants);
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

    let sheet = ss.getSheetByName('TrainingParticipants');
    if (!sheet) {
      sheet = ss.insertSheet('TrainingParticipants');
      sheet.appendRow(['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Department', 'Position', 'AddedAt']);
      sheet.getRange('A1:G1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
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

    try {
      const tpMasterSheet = getSheet(SHEET_NAMES.trainingParticipants);
      if (tpMasterSheet) {
        const tpMasterRows = resolution.participants.map(p => [
          generateId('TP'),
          trainingId,
          p.ID,
          p.Name,
          p.Department,
          p.Position,
          addedAt
        ]);
        tpMasterSheet.getRange(tpMasterSheet.getLastRow() + 1, 1, tpMasterRows.length, 7).setValues(tpMasterRows);
      }
    } catch(mErr) {
      Logger.log('Master TrainingParticipants update error: ' + mErr.message);
    }

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
    const sheet = ss.getSheetByName('TrainingParticipants');
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

/**
 * One-time repair for existing training records created before participant
 * canonicalisation. It rewrites each Training Data participant tab from the
 * Employee sheet and refreshes the linked requisition form.
 */
function repairAllTrainingParticipantData() {
  try {
    const trainingSheet = getSheet(SHEET_NAMES.trainings);
    const trainings = trainingSheet ? sheetToJson(trainingSheet) : [];
    let repaired = 0;
    let skipped = 0;

    trainings.forEach(training => {
      if (!training.ID) return;
      const trainingData = getTrainingDataSpreadsheet(training.ID);
      const participantSheet = trainingData ? trainingData.getSheetByName('TrainingParticipants') : null;
      const existingParticipants = participantSheet ? sheetToJson(participantSheet) : [];
      try {
        syncParticipantsToTrainingDriveSheet(training.ID, existingParticipants);
        syncTrainingRequisitionParticipants(training.ID);
        repaired++;
      } catch (e) {
        skipped++;
        Logger.log('Participant repair failed for ' + training.ID + ': ' + e.message);
      }
    });
    return ok({ message: `Repaired ${repaired} training participant record(s).`, repaired: repaired, skipped: skipped });
  } catch (e) {
    return err('Could not repair training participant data: ' + e.message);
  }
}
