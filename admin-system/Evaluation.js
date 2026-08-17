/**
 * Evaluation.gs — Training Evaluation and Post-Training (3-month) Evaluation
 */

// ─── Training Evaluation ────────────────────────────────────────────────────────
function getTrainingEvaluations(trainingId) {
  try {
    if (!trainingId) return ok([]);
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return ok([]);
    const sheet = ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval');
    if (!sheet) return ok([]);
    const rows = sheetToJson(sheet);
    return ok(rows);
  } catch (e) {
    return err(e.message);
  }
}

function saveTrainingEvaluation(data) {
  try {
    if (!data.TrainingID || !data.EmployeeID)
      return err('TrainingID and EmployeeID are required.');

    const ss = getTrainingDataSpreadsheet(data.TrainingID);
    if (!ss) return err('Could not open per-training sheet for ID: ' + data.TrainingID);

    let sheet = ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval');
    if (!sheet) {
      sheet = ss.insertSheet('Evaluation');
      sheet.appendRow(['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt']);
      sheet.getRange('A1:P1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // Prevent duplicate submission
    const rows = sheetToJson(sheet);
    const exists = rows.find(r =>
      String(r.TrainingID || '').trim() === String(data.TrainingID).trim() &&
      String(r.EmployeeID || '').trim().toLowerCase() === String(data.EmployeeID).trim().toLowerCase()
    );
    if (exists) return err('You have already submitted an evaluation for this training.');

    const scores = [data.Q1, data.Q2, data.Q3, data.Q4, data.Q5, data.Q6, data.Q7].map(Number);
    if (scores.some(s => isNaN(s) || s < 1 || s > 5)) {
      return err('All 7 evaluation questions (scale 1-5) are required.');
    }

    const avg = (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2);

    const evalRow = [
      generateId('EVL'),
      data.TrainingID,
      data.EmployeeID,
      data.EmployeeName  || '',
      data.Q1, data.Q2, data.Q3,
      data.Q4, data.Q5, data.Q6, data.Q7,
      data.SectionB1     || '',
      data.SectionB2     || '',
      data.SectionB3     || '',
      avg,
      now()
    ];
    sheet.appendRow(evalRow);

    // Auto-advance lifecycle stage when evaluation is submitted
    tryAdvanceToEvaluationCompleted(data.TrainingID);

    return ok({ message: 'Training evaluation submitted. Average score: ' + avg });
  } catch (e) {
    return err('Failed to save evaluation: ' + e.message);
  }
}

// ─── Post-Training Evaluation (3-month) ────────────────────────────────────────
function getPostEvaluations(trainingId) {
  try {
    if (trainingId) {
      const ss = getTrainingDataSpreadsheet(trainingId);
      if (!ss) return ok([]);
      const sheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
      if (!sheet) return ok([]);
      return ok(sheetToJson(sheet));
    }

    // Iterate across all trainings if trainingId omitted
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return ok([]);
    const trainings = sheetToJson(tSheet);
    let allPosts = [];
    trainings.forEach(t => {
      if (t.ID) {
        const ss = getTrainingDataSpreadsheet(t.ID);
        if (ss) {
          const sheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
          if (sheet) allPosts = allPosts.concat(sheetToJson(sheet));
        }
      }
    });
    return ok(allPosts);
  } catch (e) {
    return err(e.message);
  }
}

function savePostEvaluation(data) {
  try {
    if (!data.TrainingID || !data.EmployeeID || !data.EvaluatorName)
      return err('TrainingID, EmployeeID, and Evaluator Name are required.');

    const cb = Number(data.CompetencyBefore);
    const ca = Number(data.CompetencyAfter);
    if (isNaN(cb) || cb < 1 || cb > 5 || isNaN(ca) || ca < 1 || ca > 5) {
      return err('Competency levels before and after training (scale 1-5) are required.');
    }

    const ss = getTrainingDataSpreadsheet(data.TrainingID);
    if (!ss) return err('Could not open per-training sheet for ID: ' + data.TrainingID);

    let sheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
    if (!sheet) {
      sheet = ss.insertSheet('Post Evaluation');
      sheet.appendRow(['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID', 'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply', 'FurtherTraining', 'Comments', 'SubmittedAt']);
      sheet.getRange('A1:L1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const rows = sheetToJson(sheet);
    const exists = rows.find(r =>
      String(r.TrainingID || '').trim() === String(data.TrainingID).trim() &&
      String(r.EmployeeID || '').trim().toLowerCase() === String(data.EmployeeID).trim().toLowerCase()
    );
    if (exists) return err('A post-training evaluation has already been submitted for this employee.');

    const postRow = [
      generateId('PEV'),
      data.TrainingID,
      data.EmployeeID,
      data.EvaluatorName      || '',
      data.EvaluatorID        || '',
      cb,
      ca,
      data.Improvement        || '',
      data.CanApply           || '',
      data.FurtherTraining    || '',
      data.Comments           || '',
      now()
    ];
    sheet.appendRow(postRow);

    // Auto-advance training stage to Programme Closed when post-eval is submitted
    try {
      updateTrainingStage(data.TrainingID, 'Programme Closed');
    } catch (e) {
      Logger.log('Post-eval stage update error: ' + e.message);
    }

    return ok({ message: 'Post-training evaluation submitted successfully.' });
  } catch (e) {
    return err('Failed to save post-evaluation: ' + e.message);
  }
}

// ─── Evaluation Summary ─────────────────────────────────────────────────────────
function getEvaluationSummary(trainingId) {
  try {
    if (!trainingId) return ok({ evalCompleted: 0, avgScore: null, postCompleted: 0 });
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return ok({ evalCompleted: 0, avgScore: null, postCompleted: 0 });

    const evalSheet = ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval');
    const evalRows  = evalSheet ? sheetToJson(evalSheet) : [];

    const postSheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
    const postRows  = postSheet ? sheetToJson(postSheet) : [];

    const scores = evalRows.map(r => Number(r.AvgScore)).filter(n => n > 0);
    const avgScore = scores.length > 0
      ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2)
      : null;

    return ok({
      evalCompleted:  evalRows.length,
      avgScore:       avgScore,
      postCompleted:  postRows.length,
    });
  } catch (e) {
    return err(e.message);
  }
}

/**
 * API: Admin assign a Supervisor or Person in Charge for an employee's post evaluation.
 * Validates supervisor Email or Employee ID against the Employees sheet database.
 * 
 * @param {string} trainingId - Unique training programme ID
 * @param {string} employeeId - Participant employee ID
 * @param {string} supervisorInput - Supervisor Email or Employee ID
 */
function assignPostEvalSupervisor(trainingId, employeeId, supervisorInput) {
  try {
    if (!trainingId || !employeeId || !supervisorInput) {
      return err('Training ID, Employee ID, and Supervisor Email/ID are required.');
    }

    const cleanTId = String(trainingId).trim();
    const cleanEmpId = String(employeeId).trim();
    const cleanSupInput = String(supervisorInput).trim().toLowerCase();

    // 1. System check supervisor email or employee ID
    let supervisor = null;
    const empSheet = getSheet(SHEET_NAMES.employees);
    if (empSheet) {
      const employees = sheetToJson(empSheet);
      supervisor = employees.find(e => 
        isSameEmployeeId(e.ID || e.EmployeeID || e.EmployeeNo || '', cleanSupInput) ||
        String(e.Email || '').trim().toLowerCase() === cleanSupInput
      );
    }

    // Fallback lookup if not found in Employees sheet
    if (!supervisor) {
      if (cleanSupInput.includes('@')) {
        supervisor = {
          ID: cleanSupInput.split('@')[0],
          Name: cleanSupInput.split('@')[0],
          Email: cleanSupInput
        };
      } else {
        return err(`Supervisor with Email or Employee ID '${supervisorInput}' was not found in employee records.`);
      }
    }

    // 2. Open per-training spreadsheet and update TrainingParticipants
    const ss = getTrainingDataSpreadsheet(cleanTId);
    if (!ss) return err('Could not open training database.');

    let tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
    if (!tpSheet) return err('Participants sheet not found for this training.');

    ensureTrainingParticipantsColumns(tpSheet);
    const tpRows = sheetToJson(tpSheet);
    const headerRow = tpSheet.getRange(1, 1, 1, tpSheet.getLastColumn()).getValues()[0];
    
    const findColIdx = (colName) => headerRow.findIndex(h => String(h || '').trim().toLowerCase() === colName.toLowerCase());
    const supIdIdx = findColIdx('SupervisorID');
    const supEmailIdx = findColIdx('SupervisorEmail');
    const supNameIdx = findColIdx('SupervisorName');

    const targetRowIndex = tpRows.findIndex(r => isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId));
    if (targetRowIndex === -1) {
      return err(`Employee (${cleanEmpId}) is not enrolled in this training programme.`);
    }

    const rowNum = targetRowIndex + 2;
    if (supIdIdx !== -1) tpSheet.getRange(rowNum, supIdIdx + 1).setValue(supervisor.ID || supervisor.EmployeeID || '');
    if (supEmailIdx !== -1) tpSheet.getRange(rowNum, supEmailIdx + 1).setValue(supervisor.Email || '');
    if (supNameIdx !== -1) tpSheet.getRange(rowNum, supNameIdx + 1).setValue(supervisor.Name || supervisor.EmployeeName || '');

    return ok({
      message: `Supervisor ${supervisor.Name || supervisor.Email} successfully assigned to ${cleanEmpId} for post evaluation.`,
      supervisor: supervisor
    });

  } catch (e) {
    Logger.log('assignPostEvalSupervisor error: ' + e.message);
    return err('Failed to assign supervisor: ' + e.message);
  }
}

function ensureTrainingParticipantsColumns(sheet) {
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  const required = ['SupervisorID', 'SupervisorEmail', 'SupervisorName'];
  required.forEach(req => {
    const exists = headers.some(h => String(h || '').trim().toLowerCase() === req.toLowerCase());
    if (!exists) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(req).setFontWeight('bold');
    }
  });
}

function getAttendedParticipantsForPostEval(trainingId) {
  try {
    if (!trainingId) return ok({ training: null, participants: [] });

    // 1. Get official participants using getTrainingParticipants directly
    const tpRes = getTrainingParticipants(trainingId);
    let tpRows = [];
    if (tpRes && tpRes.success && Array.isArray(tpRes.data)) {
      tpRows = tpRes.data;
    } else if (Array.isArray(tpRes)) {
      tpRows = tpRes;
    }

    // 2. Fetch Attendance and Post Evaluation rows
    const ss = getTrainingDataSpreadsheet(trainingId);
    let attRows = [];
    let postRows = [];
    if (ss) {
      const attSheet = ss.getSheetByName('Attendance');
      if (attSheet) attRows = sheetToJson(attSheet);

      const postSheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
      if (postSheet) postRows = sheetToJson(postSheet);
    }

    // 3. Get master training record for 3-month milestone date
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const trainings = tSheet ? sheetToJson(tSheet) : [];
    const cleanTId = String(trainingId).trim();
    const training = trainings.find(t => 
      String(t.ID || '').trim() === cleanTId ||
      String(t.Code || '').trim() === cleanTId ||
      String(t.TrainingID || '').trim() === cleanTId ||
      (t.Name && String(t.Name).trim().toLowerCase() === cleanTId.toLowerCase())
    );

    // 4. Map tpRows to display objects
    const participantsMap = new Map();

    tpRows.forEach(p => {
      const empId = String(p.EmployeeID || p.ID || p.EmployeeNo || p.EmpID || '').trim();
      if (!empId) return;

      const empKey = empId.toLowerCase();
      const empName = String(p.EmployeeName || p.Name || empId).trim();
      const dept = String(p.Department || p.Dept || '—').trim();
      const pos = String(p.Position || p.JobTitle || 'Participant').trim();

      // Look up attendance record by Employee ID OR Employee Name
      const attendanceRecord = attRows.find(a => {
        const aId = String(a.EmployeeID || a.EmployeeNo || a.EmpID || a.ID || '').trim();
        const aName = String(a.EmployeeName || a.Name || '').trim().toLowerCase();
        return isSameEmployeeId(aId, empId) || (empName && aName && aName === empName.toLowerCase());
      });

      let attStatusText = 'Did Not Attend';
      let attColor = 'red';
      let attCode = 'ABSENT';
      let hasAttended = false;

      if (attendanceRecord) {
        const rawStatus = String(attendanceRecord.Status || 'Present').trim();
        const lower = rawStatus.toLowerCase();
        if (lower === 'absent' || lower === 'did not attend' || lower === 'not attend') {
          attStatusText = 'Did Not Attend';
          attColor = 'red';
          attCode = 'ABSENT';
          hasAttended = false;
        } else if (lower === 'late' || lower === 'partial' || lower === 'partial attendance') {
          attStatusText = 'Late / Partial';
          attColor = 'yellow';
          attCode = 'PARTIAL';
          hasAttended = true;
        } else {
          attStatusText = 'Attended';
          attColor = 'green';
          attCode = 'ATTENDED';
          hasAttended = true;
        }
      }

      const postEvalDone = postRows.some(pr => isSameEmployeeId(pr.EmployeeID || pr.EmployeeNo || pr.ID || '', empId));
      const supId = String(p.SupervisorID || '').trim();
      const supEmail = String(p.SupervisorEmail || '').trim();
      const supName = String(p.SupervisorName || '').trim();

      participantsMap.set(empKey, {
        ID: empId,
        EmployeeID: empId,
        EmployeeName: empName,
        Department: dept,
        Position: pos,
        SupervisorID: supId,
        SupervisorName: supName,
        SupervisorEmail: supEmail,
        Attended: hasAttended,
        AttendanceStatus: attStatusText,
        AttendanceColor: attColor,
        AttendanceCode: attCode,
        PostEvalCompleted: postEvalDone,
        Status: postEvalDone ? 'Evaluation Completed' : (supId || supEmail ? 'Supervisor Assigned' : 'Pending Assignment')
      });
    });

    // Fallback: If tpRows is empty, parse from master training.Participants
    if (participantsMap.size === 0 && training && training.Participants) {
      try {
        let rawParts = training.Participants;
        if (typeof rawParts === 'string') {
          const trimmed = rawParts.trim();
          if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
            try { rawParts = JSON.parse(trimmed); } catch(e) { rawParts = trimmed.split(',').map(s => s.trim()).filter(Boolean); }
          } else {
            rawParts = trimmed.split(',').map(s => s.trim()).filter(Boolean);
          }
        }
        if (Array.isArray(rawParts)) {
          const directory = getOfficialEmployeeDirectory();
          rawParts.forEach(pItem => {
            let pEmpId = typeof pItem === 'object' ? (pItem.EmployeeID || pItem.ID || '') : String(pItem || '').trim();
            if (!pEmpId && typeof pItem === 'string') pEmpId = pItem.trim();
            if (!pEmpId) return;

            const empKey = pEmpId.toLowerCase();
            if (participantsMap.has(empKey)) return;

            const empRecord = (directory && directory.byId && directory.byId[empKey]) || (typeof pItem === 'object' ? pItem : {});
            const empName = empRecord.Name || empRecord.EmployeeName || (typeof pItem === 'object' ? pItem.EmployeeName : pEmpId);
            const dept = empRecord.Department || (typeof pItem === 'object' ? pItem.Department : '') || '—';
            const pos = empRecord.Position || (typeof pItem === 'object' ? (pItem.Position || pItem.JobTitle) : '') || 'Participant';

            const attendanceRecord = attRows.find(a => {
              const aId = String(a.EmployeeID || a.EmployeeNo || a.EmpID || a.ID || '').trim();
              const aName = String(a.EmployeeName || a.Name || '').trim().toLowerCase();
              return isSameEmployeeId(aId, pEmpId) || (empName && aName && aName === empName.toLowerCase());
            });

            let attStatusText = 'Did Not Attend';
            let attColor = 'red';
            let attCode = 'ABSENT';
            let hasAttended = false;

            if (attendanceRecord) {
              const rawStatus = String(attendanceRecord.Status || 'Present').trim();
              const lower = rawStatus.toLowerCase();
              if (lower === 'absent' || lower === 'did not attend' || lower === 'not attend') {
                attStatusText = 'Did Not Attend'; attColor = 'red'; attCode = 'ABSENT'; hasAttended = false;
              } else if (lower === 'late' || lower === 'partial' || lower === 'partial attendance') {
                attStatusText = 'Late / Partial'; attColor = 'yellow'; attCode = 'PARTIAL'; hasAttended = true;
              } else {
                attStatusText = 'Attended'; attColor = 'green'; attCode = 'ATTENDED'; hasAttended = true;
              }
            }

            const postEvalDone = postRows.some(pr => isSameEmployeeId(pr.EmployeeID || pr.EmployeeNo || pr.ID || '', pEmpId));

            participantsMap.set(empKey, {
              ID: pEmpId,
              EmployeeID: pEmpId,
              EmployeeName: empName,
              Department: dept,
              Position: pos,
              SupervisorID: empRecord.SupervisorID || (typeof pItem === 'object' ? pItem.SupervisorID : '') || '',
              SupervisorName: empRecord.SupervisorName || (typeof pItem === 'object' ? pItem.SupervisorName : '') || '',
              SupervisorEmail: empRecord.SupervisorEmail || (typeof pItem === 'object' ? pItem.SupervisorEmail : '') || '',
              Attended: hasAttended,
              AttendanceStatus: attStatusText,
              AttendanceColor: attColor,
              AttendanceCode: attCode,
              PostEvalCompleted: postEvalDone,
              Status: postEvalDone ? 'Evaluation Completed' : 'Pending Assignment'
            });
          });
        }
      } catch(err) {
        Logger.log('Fallback parsing error: ' + err.message);
      }
    }

    const participants = Array.from(participantsMap.values());
    const endDate = training ? parseDateObj(training.EndDate || training.StartDate) || new Date() : new Date();
    const target3MonthDate = new Date(endDate.getTime() + 90 * 24 * 60 * 60 * 1000);
    const now = new Date();
    const is3MonthReached = now >= target3MonthDate;

    return ok({
      training: {
        ID: training ? String(training.ID || cleanTId).trim() : cleanTId,
        Name: training ? training.Name : cleanTId,
        Code: training ? (training.Code || training.ID) : cleanTId,
        EndDate: formatMinimalistDate(endDate),
        Target3MonthDate: formatMinimalistDate(target3MonthDate),
        Is3MonthReached: is3MonthReached
      },
      participants: participants
    });

  } catch (e) {
    Logger.log('getAttendedParticipantsForPostEval error: ' + e.message);
    return ok({
      training: null,
      participants: []
    });
  }
}

function assignPostEvalSupervisorsBulk(trainingId, participantIds, supervisorInput, overrideEmail, overrideName) {
  try {
    if (!trainingId || !participantIds || !Array.isArray(participantIds) || participantIds.length === 0 || !supervisorInput) {
      return err('Training ID, participant selection, and Supervisor Email/ID are required.');
    }

    const cleanTId = String(trainingId).trim();
    const cleanSupInput = String(supervisorInput).trim().toLowerCase();

    let supervisor = null;
    const empSheet = getSheet(SHEET_NAMES.employees);
    if (empSheet) {
      const employees = sheetToJson(empSheet);
      supervisor = employees.find(e => 
        isSameEmployeeId(e.ID || e.EmployeeID || e.EmployeeNo || '', cleanSupInput) ||
        String(e.Email || '').trim().toLowerCase() === cleanSupInput ||
        String(e.Name || '').trim().toLowerCase() === cleanSupInput
      );
    }

    if (!supervisor) {
      supervisor = {
        ID: cleanSupInput.split('@')[0],
        Name: overrideName || cleanSupInput.split('@')[0],
        Email: overrideEmail || (cleanSupInput.includes('@') ? cleanSupInput : '')
      };
    } else {
      supervisor = {
        ID: supervisor.ID || supervisor.EmployeeID || cleanSupInput,
        Name: overrideName || supervisor.Name || supervisor.EmployeeName || cleanSupInput,
        Email: overrideEmail || supervisor.Email || (cleanSupInput.includes('@') ? cleanSupInput : '')
      };
    }

    if (overrideEmail) supervisor.Email = String(overrideEmail).trim();
    if (overrideName) supervisor.Name = String(overrideName).trim();

    const ss = getTrainingDataSpreadsheet(cleanTId);
    if (!ss) return err('Could not open training database.');

    let tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
    if (!tpSheet) return err('Participants sheet not found.');

    ensureTrainingParticipantsColumns(tpSheet);
    const tpRows = sheetToJson(tpSheet);
    const headerRow = tpSheet.getRange(1, 1, 1, tpSheet.getLastColumn()).getValues()[0];
    
    const findColIdx = (colName) => headerRow.findIndex(h => String(h || '').trim().toLowerCase() === colName.toLowerCase());
    const supIdIdx = findColIdx('SupervisorID');
    const supEmailIdx = findColIdx('SupervisorEmail');
    const supNameIdx = findColIdx('SupervisorName');

    let updatedCount = 0;
    participantIds.forEach(pEmpId => {
      const cleanEmpId = String(pEmpId).trim();
      const targetRowIndex = tpRows.findIndex(r => isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId));
      if (targetRowIndex !== -1) {
        const rowNum = targetRowIndex + 2;
        if (supIdIdx !== -1) tpSheet.getRange(rowNum, supIdIdx + 1).setValue(supervisor.ID || supervisor.Email || '');
        if (supEmailIdx !== -1) tpSheet.getRange(rowNum, supEmailIdx + 1).setValue(supervisor.Email || '');
        if (supNameIdx !== -1) tpSheet.getRange(rowNum, supNameIdx + 1).setValue(supervisor.Name || supervisor.Email || '');
        updatedCount++;
      }
    });

    // Check if 3-month milestone is reached for this training
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const trainings = tSheet ? sheetToJson(tSheet) : [];
    const tObj = trainings.find(tr => String(tr.ID || tr.TrainingID || '').trim() === cleanTId);

    const endDate = tObj ? parseDateObj(tObj.EndDate || tObj.StartDate) || new Date() : new Date();
    const target3MonthDate = new Date(endDate.getTime() + 90 * 24 * 60 * 60 * 1000);
    const is3MonthReached = (new Date()) >= target3MonthDate;

    let emailMsg = '';
    if (is3MonthReached && supervisor.Email) {
      try {
        sendSupervisorPostEvalEmail(cleanTId, supervisor, tObj);
        emailMsg = ' 3-Month milestone reached: notification email sent to supervisor immediately.';
      } catch(e) {
        emailMsg = ' (Notification email error: ' + e.message + ')';
      }
    } else {
      emailMsg = ' Notification email will automatically be sent when the 3-month milestone is reached.';
    }

    return ok({
      message: `Supervisor ${supervisor.Name || supervisor.Email} assigned to ${updatedCount} participant(s).${emailMsg}`,
      updatedCount: updatedCount,
      supervisor: supervisor
    });

  } catch (e) {
    Logger.log('assignPostEvalSupervisorsBulk error: ' + e.message);
    return err('Failed to assign supervisor: ' + e.message);
  }
}

function editParticipantSupervisor(trainingId, employeeId, newSupervisorInput, overrideEmail, overrideName) {
  try {
    if (!trainingId || !employeeId || (!newSupervisorInput && !overrideEmail)) {
      return err('Training ID, Employee ID, and Supervisor Email/ID are required.');
    }
    return assignPostEvalSupervisorsBulk(trainingId, [employeeId], newSupervisorInput || overrideEmail, overrideEmail, overrideName);
  } catch (e) {
    Logger.log('editParticipantSupervisor error: ' + e.message);
    return err('Failed to edit supervisor: ' + e.message);
  }
}

function sendSupervisorPostEvalNotificationSingle(trainingId, employeeId) {
  try {
    if (!trainingId || !employeeId) {
      return err('Training ID and Employee ID are required.');
    }
    const cleanTId = String(trainingId).trim();
    const cleanEmpId = String(employeeId).trim();

    const ss = getTrainingDataSpreadsheet(cleanTId);
    if (!ss) return err('Training spreadsheet not found.');

    const tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
    const tpRows = tpSheet ? sheetToJson(tpSheet) : [];
    const participant = tpRows.find(r => isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId));

    if (!participant || (!participant.SupervisorID && !participant.SupervisorEmail)) {
      return err(`No supervisor assigned for employee (${cleanEmpId}). Please assign a supervisor first.`);
    }

    const tSheet = getSheet(SHEET_NAMES.trainings);
    const trainings = tSheet ? sheetToJson(tSheet) : [];
    const t = trainings.find(tr => String(tr.ID || tr.TrainingID || '').trim() === cleanTId);

    const supervisor = {
      ID: participant.SupervisorID || participant.SupervisorEmail.split('@')[0],
      Name: participant.SupervisorName || participant.SupervisorID || participant.SupervisorEmail,
      Email: participant.SupervisorEmail
    };

    if (!supervisor.Email) {
      const empSheet = getSheet(SHEET_NAMES.employees);
      if (empSheet) {
        const employees = sheetToJson(empSheet);
        const sEmp = employees.find(e => isSameEmployeeId(e.ID || e.EmployeeID || '', supervisor.ID));
        if (sEmp && sEmp.Email) supervisor.Email = sEmp.Email;
      }
    }

    if (!supervisor.Email) {
      return err(`Supervisor (${supervisor.Name}) does not have a valid email address configured.`);
    }

    sendSupervisorPostEvalEmail(cleanTId, supervisor, t);
    return ok(`Notification email sent to supervisor ${supervisor.Name} (${supervisor.Email}).`);

  } catch (e) {
    Logger.log('sendSupervisorPostEvalNotificationSingle error: ' + e.message);
    return err('Failed to send notification email: ' + e.message);
  }
}

function sendSupervisorPostEvalEmail(trainingId, supervisor, trainingObj) {
  try {
    if (!supervisor || !supervisor.Email) return;
    const publicUrl = getPublicPortalUrl() || getAppUrl();
    const reviewUrl = `${publicUrl}?page=posteval&id=${trainingId}&eval=${encodeURIComponent(supervisor.ID || supervisor.Email)}`;
    const tName = trainingObj ? trainingObj.Name : trainingId;

    const subject = `[TrainHub] 3-Month Post-Training Competency Review Required - ${tName}`;
    const body = `Dear ${supervisor.Name || 'Supervisor / PIC'},\n\n` +
      `The 3-month milestone after course completion has elapsed for training programme:\n` +
      `Training Name: ${tName} (${trainingObj ? (trainingObj.Code || trainingId) : trainingId})\n\n` +
      `You have been assigned as the Supervisor / Person In Charge to evaluate assigned participant(s).\n\n` +
      `Please click the link below to access the 3-Month Competency Review Portal:\n` +
      `${reviewUrl}\n\n` +
      `Thank you,\nTrainHub Training Management System`;

    MailApp.sendEmail(supervisor.Email, subject, body);
  } catch (e) {
    Logger.log('sendSupervisorPostEvalEmail error: ' + e.message);
  }
}

/**
 * Daily Cron Task: Checks all active trainings. If 3-month milestone (90 days post EndDate) is reached,
 * emails all assigned supervisors who haven't completed post-evaluation yet.
 */
function cronCheck3MonthPostEvalNotifications() {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return;
    const trainings = sheetToJson(tSheet);
    const now = new Date();

    trainings.forEach(t => {
      const tId = String(t.ID || t.Code || '').trim();
      if (!tId) return;

      const endDate = parseDateObj(t.EndDate || t.StartDate);
      if (!endDate) return;

      const target3MonthDate = new Date(endDate.getTime() + 90 * 24 * 60 * 60 * 1000);
      if (now < target3MonthDate) return;

      const ss = getTrainingDataSpreadsheet(tId);
      if (!ss) return;

      const tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
      if (!tpSheet) return;

      const tpRows = sheetToJson(tpSheet);
      const postSheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
      const postRows = postSheet ? sheetToJson(postSheet) : [];

      const notifiedSupervisors = new Set();

      tpRows.forEach(p => {
        const empId = String(p.EmployeeID || p.ID || '').trim();
        if (!empId) return;

        const postCompleted = postRows.some(pr => isSameEmployeeId(pr.EmployeeID || pr.ID || '', empId));
        if (postCompleted) return;

        const supEmail = String(p.SupervisorEmail || '').trim();
        const supId = String(p.SupervisorID || '').trim();
        const supName = String(p.SupervisorName || supId || supEmail).trim();

        if (!supEmail || notifiedSupervisors.has(supEmail.toLowerCase())) return;

        const supervisorObj = { ID: supId || supEmail, Name: supName, Email: supEmail };
        sendSupervisorPostEvalEmail(tId, supervisorObj, t);
        notifiedSupervisors.add(supEmail.toLowerCase());
        Logger.log(`Automated 3-Month Post Eval Email sent to ${supEmail} for training ${tId}`);
      });
    });

  } catch(e) {
    Logger.log('cronCheck3MonthPostEvalNotifications error: ' + e.message);
  }
}

// ─── Internal: auto-advance stage ──────────────────────────────────────────────
function tryAdvanceToEvaluationCompleted(trainingId) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const tRows  = sheetToJson(tSheet);
    const t = tRows.find(r => r.ID === trainingId);
    if (!t) return;

    const ss = getTrainingDataSpreadsheet(trainingId);
    const eSheet = ss ? (ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval')) : null;
    const evalCount = eSheet ? sheetToJson(eSheet).length : 0;

    if (evalCount > 0 && ['Created', 'Participants Imported', 'Attendance In Progress', 'Training Completed'].includes(t.Stage)) {
      updateTrainingStage(trainingId, 'Evaluation Completed');
    }
  } catch (e) {
    Logger.log('Stage auto-advance error: ' + e.message);
  }
}
