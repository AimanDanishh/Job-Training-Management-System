/**
 * Code.gs — HOD Portal Web App Entry Point & Server Engine
 */

function doGet(e) {
  const pageParam = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase().trim() : 'review';

  const pageMap = {
    'review':   'HodReview',
    'posteval': 'HodPostEvaluation',
    'success':  'Success',
    'error':    'Error'
  };

  const templateName = pageMap[pageParam] || 'HodReview';
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — HOD Management Portal');

  try {
    const template = HtmlService.createTemplateFromFile(templateName);
    template.params = (e && e.parameter) ? e.parameter : {};
    template.page = pageParam;

    return template.evaluate()
      .setTitle(appTitle)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log('doGet routing error: ' + err.message);
    try {
      const errTemplate = HtmlService.createTemplateFromFile('Error');
      errTemplate.params = { message: 'Failed to load page: ' + err.message };
      return errTemplate.evaluate()
        .setTitle('Error — TrainHub HOD Portal')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (fallbackErr) {
      return HtmlService.createHtmlOutput('<h3 style="font-family:sans-serif;color:#ef4444;padding:20px;">System Error: ' + err.message + '</h3>');
    }
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return '';
  }
}


/**
 * API: Fetch complete Training Requisition details for HOD Review
 */
function getRequisitionDetails(trainingId) {
  try {
    const auth = validateHODAccess();
    if (!auth.valid) return err(auth.message);

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const trainings = sheetToJson(tSheet);
    let training = null;

    if (trainingId) {
      const cleanId = String(trainingId).trim().toLowerCase();
      training = trainings.find(t =>
        String(t.ID || '').toLowerCase() === cleanId ||
        String(t.Code || '').toLowerCase() === cleanId
      );
    }

    // Fallback: Default to most recent pending or first training
    if (!training && trainings.length > 0) {
      training = trainings.find(t => String(t.ApprovalStatus || t.Status || '').toLowerCase().includes('pending')) || trainings[0];
    }

    if (!training) return err('Training requisition request not found.');

    // Fetch Participants for this training
    const partSheet = getSheet('TrainingParticipants');
    let participants = [];
    if (partSheet) {
      const allParts = sheetToJson(partSheet);
      participants = allParts.filter(p => String(p.TrainingID || '').trim() === String(training.ID || '').trim());
    }

    // Requester employee info lookup
    let requester = {
      ID: training.RequestedBy || training.EmployeeID || 'N/A',
      Name: training.RequestedByName || training.Trainer || 'Employee Requester',
      Department: training.Department || 'N/A',
      Email: training.RequestedByEmail || ''
    };

    const empSheet = getSheet('Employees');
    if (empSheet && requester.ID !== 'N/A') {
      const emps = sheetToJson(empSheet);
      const matched = emps.find(e => String(e.ID || '').toLowerCase() === String(requester.ID).toLowerCase() || String(e.Email || '').toLowerCase() === String(requester.Email).toLowerCase());
      if (matched) {
        requester.Name = matched.Name || requester.Name;
        requester.Department = matched.Department || requester.Department;
        requester.Email = matched.Email || requester.Email;
      }
    }

    return ok({
      training: training,
      participants: participants,
      requester: requester,
      hod: auth.hod
    });
  } catch (e) {
    return err('Failed to load requisition details: ' + e.message);
  }
}

/**
 * API: Submit HOD Decision (Approve / Reject / Postpone / Reschedule)
 */
function submitHODDecision(data) {
  try {
    if (!data || !data.trainingId || !data.decision) {
      return err('Training ID and decision are required.');
    }

    const auth = validateHODAccess();
    if (!auth.valid) return err(auth.message);

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const cleanId = String(data.trainingId).trim();
    const row = findRowById(tSheet, cleanId);
    if (row === -1) return err(`Training request (${cleanId}) not found in database.`);

    const headers = tSheet.getDataRange().getValues()[0].map(h => String(h).trim());
    const timestamp = formatDateTime(new Date());

    const updateCol = (colName, val) => {
      const idx = headers.indexOf(colName) + 1;
      if (idx > 0) {
        tSheet.getRange(row, idx).setValue(val);
      }
    };

    const validDecision = String(data.decision).trim(); // Approved, Rejected, Postponed, Rescheduled
    const hodId = data.hodId || auth.hod.ID;
    const hodName = data.hodName || auth.hod.Name;
    const hodCostCentre = data.hodCostCentre || auth.hod.CostCentre;
    const remarks = data.remarks || '';
    const rescheduledDate = data.rescheduledDate || '';

    updateCol('ApprovalStatus', validDecision);
    updateCol('ApprovedBy', `${hodName} (${hodId})`);
    updateCol('ApprovedCostCentre', hodCostCentre);
    updateCol('ApprovedAt', timestamp);
    updateCol('ApprovalRemarks', remarks);
    if (rescheduledDate) updateCol('RescheduledDate', rescheduledDate);

    // Update Status & Stage if Approved or Rejected/Postponed
    if (validDecision === 'Approved') {
      updateCol('Status', 'Upcoming');
      updateCol('Stage', 'Created');
    } else if (validDecision === 'Rejected') {
      updateCol('Status', 'Cancelled');
      updateCol('Stage', 'Programme Closed');
    } else if (validDecision === 'Postponed' || validDecision === 'Rescheduled') {
      updateCol('Status', 'Postponed');
      updateCol('Stage', 'Created');
    }

    updateCol('UpdatedDate', timestamp);
    SpreadsheetApp.flush();

    // Send email notification to requester
    try {
      const trainings = sheetToJson(tSheet);
      const currentT = trainings.find(t => String(t.ID) === cleanId);
      let requesterEmail = currentT ? (currentT.RequestedByEmail || currentT.Email || '') : '';
      const requesterId = currentT ? (currentT.RequestedBy || '') : '';

      if (!requesterEmail && requesterId && getSheet('Employees')) {
        const emps = sheetToJson(getSheet('Employees'));
        const empMatch = emps.find(e => String(e.ID).toLowerCase() === String(requesterId).toLowerCase());
        if (empMatch) requesterEmail = empMatch.Email || '';
      }

      if (requesterEmail) {
        const subject = `[TrainHub] Training Requisition Request ${validDecision.toUpperCase()} - ${currentT.Name || cleanId}`;
        const body = `Dear Requester,\n\nYour Training Requisition Request for "${currentT.Name || cleanId}" (${currentT.Code || cleanId}) has been updated by HOD/Manager:\n\n` +
          `Decision Status: ${validDecision.toUpperCase()}\n` +
          `Reviewed By: ${hodName} (${hodId})\n` +
          `Cost Centre: ${hodCostCentre}\n` +
          `Timestamp: ${timestamp}\n` +
          (rescheduledDate ? `Rescheduled Date: ${rescheduledDate}\n` : '') +
          (remarks ? `HOD Remarks: ${remarks}\n` : '') +
          `\nThank you,\nTrainHub Training Management System`;

        MailApp.sendEmail(requesterEmail, subject, body);
      }
    } catch (mailErr) {
      Logger.log('Notification mail error: ' + mailErr.message);
    }

    return ok({
      trainingId: cleanId,
      decision: validDecision,
      approvedBy: hodName,
      timestamp: timestamp,
      message: `Training requisition request successfully marked as ${validDecision.toUpperCase()}.`
    });
  } catch (e) {
    return err('Failed to process HOD decision: ' + e.message);
  }
}

/**
 * API: Fetch participants pending 3-Month Post Evaluation for a training under this HOD
 */
function getPendingPostEvalParticipants(trainingId, hodCostCentre) {
  try {
    const auth = validateHODAccess();
    if (!auth.valid) return err(auth.message);

    const targetHodCostCentre = hodCostCentre || auth.hod.CostCentre;

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const trainings = sheetToJson(tSheet);
    let training = null;

    if (trainingId) {
      const cleanId = String(trainingId).trim().toLowerCase();
      training = trainings.find(t =>
        String(t.ID || '').toLowerCase() === cleanId ||
        String(t.Code || '').toLowerCase() === cleanId
      );
    }

    if (!training && trainings.length > 0) {
      training = trainings.find(t => String(t.Stage || '').includes('3-Month') || String(t.Stage || '').includes('6-Month') || String(t.Status || '').includes('Completed')) || trainings[0];
    }

    if (!training) return err('No training programme found for post evaluation.');

    // Fetch all participants enrolled in this training
    const partSheet = getSheet('TrainingParticipants');
    let allParticipants = [];
    if (partSheet) {
      const pRows = sheetToJson(partSheet);
      allParticipants = pRows.filter(p => String(p.TrainingID || '').trim() === String(training.ID || '').trim());
    }

    // Filter participants under this HOD's Cost Centre / Department
    let hodParticipants = allParticipants.filter(p => {
      const dept = String(p.Department || p.CostCentre || '').toLowerCase();
      const hodCc = String(targetHodCostCentre).toLowerCase();
      return dept.includes(hodCc) || hodCc.includes(dept) || targetHodCostCentre === 'ALL';
    });

    // If filter yields none, fallback to all participants for this training
    if (hodParticipants.length === 0) {
      hodParticipants = allParticipants;
    }

    // Check which participants have already been evaluated in PostEval sheet
    const postSheet = getSheet('PostEval');
    let evaluatedEmpIds = [];
    if (postSheet) {
      const postRows = sheetToJson(postSheet);
      evaluatedEmpIds = postRows
        .filter(pr => String(pr.TrainingID || '').trim() === String(training.ID || '').trim())
        .map(pr => String(pr.EmployeeID || '').trim().toLowerCase());
    }

    const pendingList = hodParticipants.filter(p => !evaluatedEmpIds.includes(String(p.EmployeeID || p.ID || '').trim().toLowerCase()));
    const completedList = hodParticipants.filter(p => evaluatedEmpIds.includes(String(p.EmployeeID || p.ID || '').trim().toLowerCase()));

    return ok({
      training: training,
      hod: auth.hod,
      totalUnderHod: hodParticipants.length,
      pendingCount: pendingList.length,
      completedCount: completedList.length,
      pendingParticipants: pendingList,
      completedParticipants: completedList
    });
  } catch (e) {
    return err('Failed to load post evaluation participants: ' + e.message);
  }
}

/**
 * API: Submit HOD Post Evaluation for a participant
 */
function submitHODPostEval(data) {
  try {
    if (!data || !data.trainingId || !data.employeeId) {
      return err('Training ID and Employee ID are required.');
    }

    const auth = validateHODAccess();
    if (!auth.valid) return err(auth.message);

    const postSheet = getSheet('PostEval');
    if (!postSheet) return err('PostEval sheet unavailable.');

    const cleanTId = String(data.trainingId).trim();
    const cleanEmpId = String(data.employeeId).trim();

    // Check for duplicate submission
    const existingRows = sheetToJson(postSheet);
    const duplicate = existingRows.find(r =>
      String(r.TrainingID || '').trim() === cleanTId &&
      String(r.EmployeeID || '').trim().toLowerCase() === cleanEmpId.toLowerCase()
    );

    if (duplicate) {
      return err(`Post evaluation for employee (${cleanEmpId}) has already been submitted.`);
    }

    const evalId = 'PEVAL-' + Math.floor(100000 + Math.random() * 900000);
    const evaluatorName = auth.hod.Name;
    const evaluatorId = auth.hod.ID;
    const submittedAt = formatDateTime(new Date());

    const rowData = [
      evalId,
      cleanTId,
      cleanEmpId,
      evaluatorName,
      evaluatorId,
      data.competencyBefore || '3 - Proficient',
      data.competencyAfter  || '5 - Expert',
      data.improvement       || 'High',
      data.canApply          || 'Yes',
      data.furtherTraining   || 'None required',
      data.comments          || 'Good progress demonstrated in work output.',
      submittedAt
    ];

    postSheet.appendRow(rowData);
    SpreadsheetApp.flush();

    // Return updated pending list for this HOD
    return getPendingPostEvalParticipants(cleanTId, auth.hod.CostCentre);
  } catch (e) {
    return err('Failed to submit post evaluation: ' + e.message);
  }
}
