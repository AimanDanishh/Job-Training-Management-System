/**
 * Evaluation.gs - Training Evaluation and Post-Training (3-month) Evaluation
 */

// --- Training Evaluation --------------------------------------------------------
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

// --- Post-Training Evaluation (3-month) ----------------------------------------
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

// --- Evaluation Summary ---------------------------------------------------------
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
    const cleanTId = String(trainingId).trim();

    // 1. Get official participants using robust multi-tier resolver
    let tpRows = getTrainingParticipantsList(cleanTId) || [];

    // 2. Fetch Attendance and Post Evaluation rows
    const ss = getTrainingDataSpreadsheet(cleanTId);
    let attRows = [];
    let postRows = [];
    if (ss) {
      const allSheets = ss.getSheets();
      const attSheet = allSheets.find(s => s.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === 'attendance');
      if (attSheet) attRows = sheetToJson(attSheet);

      const postSheet = allSheets.find(s => {
        const clean = s.getName().toLowerCase().replace(/[^a-z0-9]/g, '');
        return clean === 'postevaluation' || clean === 'posteval';
      });
      if (postSheet) postRows = sheetToJson(postSheet);
    }

    // 3. Get master training record for 3-month milestone date
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const trainings = tSheet ? sheetToJson(tSheet) : [];
    const training = trainings.find(t => 
      String(t.ID || '').trim().toLowerCase() === cleanTId.toLowerCase() ||
      String(t.Code || '').trim().toLowerCase() === cleanTId.toLowerCase() ||
      String(t.TrainingID || '').trim().toLowerCase() === cleanTId.toLowerCase() ||
      (t.Name && String(t.Name).trim().toLowerCase() === cleanTId.toLowerCase())
    );

    // If tpRows is still empty, check if Attendance sheet has participants
    if (tpRows.length === 0 && attRows.length > 0) {
      const uniqueFromAtt = new Map();
      attRows.forEach(a => {
        const empId = String(a.EmployeeNo || a.EmployeeID || a.EmpID || a.ID || '').trim();
        if (empId && !uniqueFromAtt.has(empId.toLowerCase())) {
          uniqueFromAtt.set(empId.toLowerCase(), {
            ID: empId,
            EmployeeID: empId,
            EmployeeName: a.EmployeeName || a.Name || empId,
            Department: a.Department || '-',
            Position: a.Position || 'Participant',
            SupervisorID: a.SupervisorID || '',
            SupervisorName: a.SupervisorName || '',
            SupervisorEmail: a.SupervisorEmail || ''
          });
        }
      });
      tpRows = Array.from(uniqueFromAtt.values());
    }

    // 4. Map tpRows to display objects
    const participantsMap = new Map();

    tpRows.forEach(p => {
      let empId = String(p.EmployeeID || p.EmployeeNo || p.EmpID || p.ID || '').trim();
      if (!empId && typeof p === 'string' && !/^\d+$/.test(p.trim())) empId = p.trim();
      if (!empId) return;
      // Guard: Ignore if empId is purely a participant count (e.g. "47")
      if (/^\d+$/.test(empId) && (!p.EmployeeName || p.EmployeeName === empId) && !p.Department && tpRows.length === 1) {
        return;
      }

      const empKey = empId.toLowerCase();
      const empName = String(p.EmployeeName || p.Name || empId).trim();
      const dept = String(p.Department || p.Dept || '-').trim();
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

    // Fallback: If participantsMap is still empty, parse from master training.ParticipantList
    if (participantsMap.size === 0 && training) {
      try {
        let rawParts = training.ParticipantList || training.ParticipantsList;
        if (typeof rawParts === 'string') {
          const trimmed = rawParts.trim();
          if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
            try { rawParts = JSON.parse(trimmed); } catch(e) { rawParts = []; }
          } else {
            rawParts = [];
          }
        }
        if (Array.isArray(rawParts) && rawParts.length > 0) {
          const directory = getOfficialEmployeeDirectory();
          rawParts.forEach(pItem => {
            let pEmpId = typeof pItem === 'object' ? (pItem.EmployeeID || pItem.ID || '') : String(pItem || '').trim();
            if (!pEmpId && typeof pItem === 'string') pEmpId = pItem.trim();
            if (!pEmpId) return;

            const empKey = pEmpId.toLowerCase();
            if (participantsMap.has(empKey)) return;

            const empRecord = (directory && directory.byId && directory.byId[empKey]) || (typeof pItem === 'object' ? pItem : {});
            const empName = empRecord.Name || empRecord.EmployeeName || (typeof pItem === 'object' ? pItem.EmployeeName : pEmpId);
            const dept = empRecord.Department || (typeof pItem === 'object' ? pItem.Department : '') || '-';
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

/**
 * Sends a supervisor Post Evaluation email using Apollo's established styling and subject format.
 * Supports consolidated multi-participant formatting.
 * 
 * @param {string} trainingId - Unique training ID
 * @param {Object} supervisor - { ID, Name, Email }
 * @param {Object} trainingObj - Training record
 * @param {Array} participantsList - Optional list of participants assigned to this supervisor
 * @param {string} testRecipientEmail - Optional test recipient email for safe redirection in Test Mode
 */
function sendSupervisorPostEvalEmail(trainingId, supervisor, trainingObj, participantsList, testRecipientEmail) {
  if (!supervisor || !supervisor.Email) {
    throw new Error('Supervisor email address is missing or invalid.');
  }

  const cleanTId = String(trainingId || '').trim();
  const tName = trainingObj ? String(trainingObj.Name || cleanTId) : cleanTId;
  const tCode = trainingObj ? String(trainingObj.Code || cleanTId) : cleanTId;
  const compDateStr = trainingObj ? String(trainingObj.EndDate || trainingObj.StartDate || '') : '';
  const compDateFormatted = formatMinimalistDate(compDateStr) || compDateStr || 'Recently';

  const publicUrl = getPublicPortalUrl() || getAppUrl();
  const reviewUrl = `${publicUrl}?page=post&id=${encodeURIComponent(cleanTId)}&eval=${encodeURIComponent(supervisor.ID || supervisor.Email)}`;

  const isTestMode = Boolean(testRecipientEmail);
  const targetEmail = isTestMode ? String(testRecipientEmail).trim() : supervisor.Email;
  const subjectPrefix = isTestMode ? '[TEST MODE - REDIRECTED] ' : '';
  const subject = `${subjectPrefix}[Apollo] 3-Month Post-Training Competency Review Required - ${tName}`;

  // Build participants table/list
  let participantListText = '';
  let participantListHtml = '';

  if (Array.isArray(participantsList) && participantsList.length > 0) {
    participantListText = '\nAssigned Participant(s):\n' +
      participantsList.map((p, idx) => `  ${idx + 1}. ${p.EmployeeName || p.Name || 'Employee'} (${p.EmployeeID || p.EmployeeNo || p.ID || '-'}) - ${p.Department || 'Department'}`).join('\n') + '\n\n';

    participantListHtml = `
      <div style="margin: 16px 0; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 14px 16px;">
        <div style="font-size: 13px; font-weight: 700; color: #1E293B; margin-bottom: 8px;">Assigned Participant(s) for Review:</div>
        <table style="width: 100%; border-collapse: collapse; font-size: 12.5px; color: #334155;">
          <thead>
            <tr style="border-bottom: 1px solid #CBD5E1; text-align: left;">
              <th style="padding: 6px 8px; font-size: 11.5px; color: #64748B;">No.</th>
              <th style="padding: 6px 8px; font-size: 11.5px; color: #64748B;">Employee Name</th>
              <th style="padding: 6px 8px; font-size: 11.5px; color: #64748B;">Employee ID</th>
              <th style="padding: 6px 8px; font-size: 11.5px; color: #64748B;">Department</th>
            </tr>
          </thead>
          <tbody>
            ${participantsList.map((p, idx) => `
              <tr style="border-bottom: 1px solid #F1F5F9;">
                <td style="padding: 6px 8px; font-weight: 600;">${idx + 1}</td>
                <td style="padding: 6px 8px; font-weight: 700; color: #0F172A;">${p.EmployeeName || p.Name || 'Employee'}</td>
                <td style="padding: 6px 8px; color: #475569;">${p.EmployeeID || p.EmployeeNo || p.ID || '-'}</td>
                <td style="padding: 6px 8px; color: #64748B;">${p.Department || '-'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  const plainText = `Dear ${supervisor.Name || 'Supervisor / PIC'},\n\n` +
    `The 3-month milestone after course completion has elapsed for training programme:\n` +
    `Training Name: ${tName} (${tCode})\n` +
    `Course Completed: ${compDateFormatted}\n\n` +
    `You have been assigned as the Supervisor / Person In Charge to evaluate the performance and competency improvement of the assigned participant(s).\n` +
    participantListText +
    `Please click the link below to access the 3-Month Competency Review Portal:\n` +
    `${reviewUrl}\n\n` +
    `Thank you,\nApollo Job Training Management System`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #1E293B; line-height: 1.5;">
      <div style="background: #2563EB; padding: 18px 24px; border-radius: 8px 8px 0 0; color: #FFFFFF;">
        <h2 style="margin: 0; font-size: 18px; font-weight: 800;">Apollo Job Training Management System</h2>
        <div style="font-size: 12px; opacity: 0.9; margin-top: 4px;">3-Month Post-Training Competency Review Notification</div>
      </div>
      <div style="border: 1px solid #E2E8F0; border-top: none; padding: 24px; border-radius: 0 0 8px 8px; background: #FFFFFF;">
        <p style="font-size: 14px; margin-top: 0;">Dear <strong>${supervisor.Name || 'Supervisor / PIC'}</strong>,</p>
        <p style="font-size: 13px; color: #475569;">
          The <strong>3-month post-training milestone</strong> has been reached for the following training programme.
          As the designated supervisor, please complete the competency review for your assigned participant(s).
        </p>

        <div style="background: #F8FAFC; border-left: 4px solid #2563EB; padding: 12px 16px; border-radius: 4px; margin: 16px 0;">
          <div style="font-size: 14px; font-weight: 800; color: #1E293B;">${tName}</div>
          <div style="font-size: 12px; color: #64748B; margin-top: 3px;">
            Training Code: <strong>${tCode}</strong> &bull; Completed: <strong>${compDateFormatted}</strong>
          </div>
        </div>

        ${participantListHtml}

        <div style="text-align: center; margin: 24px 0;">
          <a href="${reviewUrl}" target="_blank" style="background: #2563EB; color: #FFFFFF; text-decoration: none; padding: 12px 28px; border-radius: 6px; font-size: 14px; font-weight: 700; display: inline-block;">
            Open 3-Month Competency Review Portal &rarr;
          </a>
        </div>

        <p style="font-size: 11.5px; color: #94A3B8; margin-bottom: 0;">
          If the button above does not work, copy and paste this link into your browser:<br/>
          <a href="${reviewUrl}" style="color: #2563EB; word-break: break-all;">${reviewUrl}</a>
        </p>
        <hr style="border: none; border-top: 1px solid #E2E8F0; margin: 20px 0 14px;" />
        <div style="font-size: 11px; color: #94A3B8; text-align: center;">
          This is an automated system email from the Apollo Job Training Management System. Please do not reply directly.
        </div>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: targetEmail,
    subject: subject,
    body: plainText,
    htmlBody: htmlBody
  });

  Logger.log(`Post Eval Email sent to ${targetEmail} (Supervisor: ${supervisor.Name || supervisor.Email}) for training ${cleanTId}`);
  return { success: true, targetEmail: targetEmail };
}

/**
 * Processes and sends 3-Month Post Evaluation emails for a specific training programme.
 * Enforces supervisor validation, per-supervisor grouping, idempotency, and detailed state logging.
 * 
 * @param {string} trainingId - Training programme ID
 * @param {Object} options - { simulatedDate, testEmailRecipient, forceResend }
 */
function sendSupervisorPostEvaluationEmailsForTraining(trainingId, options) {
  options = options || {};
  const cleanTId = String(trainingId || '').trim();
  if (!cleanTId) return err('Training ID is required.');

  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) return err('Trainings database sheet not found.');

  const headers = ensureTrainingSheetColumns(tSheet);
  const trainings = sheetToJson(tSheet);
  const tObj = trainings.find(tr => String(tr.ID || tr.Code || '').trim() === cleanTId);
  if (!tObj) return err(`Training record '${cleanTId}' not found.`);

  const now = options.simulatedDate ? new Date(options.simulatedDate) : new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // Check 3-month milestone using the standard single source of truth
  const compDateRaw = tObj.EndDate || tObj.StartDate;
  let compDate = null;
  if (compDateRaw instanceof Date) {
    compDate = !isNaN(compDateRaw.getTime()) ? new Date(compDateRaw.getFullYear(), compDateRaw.getMonth(), compDateRaw.getDate()) : null;
  } else {
    const str = String(compDateRaw || '').trim();
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

  if (!compDate) {
    return err('Cannot determine training completion date for 3-month milestone check.');
  }

  const targetY = compDate.getFullYear() + Math.floor((compDate.getMonth() + 3) / 12);
  const targetM = (compDate.getMonth() + 3) % 12;
  const daysInTargetMonth = new Date(targetY, targetM + 1, 0).getDate();
  const targetDay = Math.min(compDate.getDate(), daysInTargetMonth);
  const threeMonthMidnight = new Date(targetY, targetM, targetDay);

  const isMilestoneReached = todayMidnight.getTime() >= threeMonthMidnight.getTime();
  if (!isMilestoneReached && !options.forceResend) {
    return err(`3-Month milestone has not yet been reached for this training (Due: ${formatMinimalistDate(threeMonthMidnight)}).`);
  }

  // 1. Fetch participants
  let participants = [];
  try {
    participants = getTrainingParticipantsList(cleanTId);
  } catch(pErr) {
    return err('Failed to retrieve training participants: ' + pErr.message);
  }

  if (!participants || participants.length === 0) {
    return ok({
      trainingId: cleanTId,
      status: 'NO_PARTICIPANTS',
      message: 'No participants enrolled in this training programme.'
    });
  }

  // 2. Detect missing supervisor assignments
  const unassignedParticipants = participants.filter(p => {
    const supId = String(p.SupervisorID || '').trim();
    const supEmail = String(p.SupervisorEmail || '').trim();
    const supName = String(p.SupervisorName || '').trim();
    return !supId && !supEmail && !supName;
  });

  const statusCol = headers.indexOf('PostEvalEmailStatus') + 1;
  const sentAtCol = headers.indexOf('PostEvalEmailSentAt') + 1;
  const errorCol  = headers.indexOf('PostEvalEmailError') + 1;
  const logCol    = headers.indexOf('PostEvalEmailLog') + 1;

  if (unassignedParticipants.length > 0) {
    if (tObj._row && statusCol > 0) {
      tSheet.getRange(tObj._row, statusCol).setValue('SUPERVISOR_MISSING');
      if (errorCol > 0) tSheet.getRange(tObj._row, errorCol).setValue(`${unassignedParticipants.length} participant(s) require supervisor assignment.`);
    }
    return ok({
      trainingId: cleanTId,
      status: 'SUPERVISOR_MISSING',
      unassignedCount: unassignedParticipants.length,
      message: `Post Evaluation email cannot be sent for ${unassignedParticipants.length} participant(s) because supervisors have not been assigned.`
    });
  }

  // 3. Group participants by Supervisor Email
  const supervisorMap = {};
  participants.forEach(p => {
    let email = String(p.SupervisorEmail || '').trim().toLowerCase();
    if (!email && p.SupervisorID) {
      email = String(p.SupervisorID).trim().toLowerCase();
      if (!email.includes('@')) email = `${email}@apollofood.com.my`;
    }
    if (!email) return;

    if (!supervisorMap[email]) {
      supervisorMap[email] = {
        ID: String(p.SupervisorID || email).trim(),
        Name: String(p.SupervisorName || p.SupervisorID || email).trim(),
        Email: email,
        participants: []
      };
    }
    supervisorMap[email].participants.push(p);
  });

  const supervisorEmails = Object.keys(supervisorMap);
  if (supervisorEmails.length === 0) {
    return err('No valid supervisor email addresses found for participants.');
  }

  // 4. Parse existing dispatch log
  let currentLog = {};
  try {
    if (tObj.PostEvalEmailLog) {
      currentLog = typeof tObj.PostEvalEmailLog === 'string' ? JSON.parse(tObj.PostEvalEmailLog) : tObj.PostEvalEmailLog;
    }
  } catch(e) {
    currentLog = {};
  }
  if (!currentLog || typeof currentLog !== 'object') currentLog = {};
  if (!currentLog.supervisors) currentLog.supervisors = {};

  let successCount = 0;
  let failCount = 0;
  let skippedCount = 0;
  const dispatchErrors = [];

  // 5. Send consolidated email per supervisor
  supervisorEmails.forEach(supEmail => {
    const supObj = supervisorMap[supEmail];
    const prevEntry = currentLog.supervisors[supEmail];

    // Idempotency: Skip if already SENT unless forced
    if (prevEntry && prevEntry.status === 'SENT' && !options.forceResend) {
      skippedCount++;
      return;
    }

    try {
      sendSupervisorPostEvalEmail(cleanTId, supObj, tObj, supObj.participants, options.testEmailRecipient);
      currentLog.supervisors[supEmail] = {
        supervisorId: supObj.ID,
        supervisorName: supObj.Name,
        email: supEmail,
        status: 'SENT',
        sentAt: now(),
        participantCount: supObj.participants.length
      };
      successCount++;
    } catch(sendErr) {
      const errMsg = sendErr.message || 'Unknown email service error';
      currentLog.supervisors[supEmail] = {
        supervisorId: supObj.ID,
        supervisorName: supObj.Name,
        email: supEmail,
        status: 'FAILED',
        error: errMsg,
        attemptedAt: now(),
        participantCount: supObj.participants.length
      };
      failCount++;
      dispatchErrors.push(`${supObj.Name || supEmail}: ${errMsg}`);
    }
  });

  // 6. Update overall state in database
  let overallStatus = 'PENDING';
  let overallError = '';
  const totalSupervisors = supervisorEmails.length;
  const allSent = supervisorEmails.every(e => currentLog.supervisors[e] && currentLog.supervisors[e].status === 'SENT');

  if (allSent) {
    overallStatus = 'SENT';
    overallError = '';
  } else if (failCount > 0) {
    overallStatus = 'FAILED';
    overallError = dispatchErrors.join('; ');
  } else {
    overallStatus = tObj.PostEvalEmailStatus || 'PENDING';
  }

  currentLog.lastUpdated = now();
  currentLog.milestoneDate = threeMonthMidnight.toISOString();

  if (tObj._row) {
    if (statusCol > 0) tSheet.getRange(tObj._row, statusCol).setValue(overallStatus);
    if (allSent && sentAtCol > 0) tSheet.getRange(tObj._row, sentAtCol).setValue(now());
    if (errorCol > 0) tSheet.getRange(tObj._row, errorCol).setValue(overallError);
    if (logCol > 0) tSheet.getRange(tObj._row, logCol).setValue(JSON.stringify(currentLog));
  }

  return ok({
    trainingId: cleanTId,
    status: overallStatus,
    totalSupervisors: totalSupervisors,
    sentCount: successCount,
    failedCount: failCount,
    skippedCount: skippedCount,
    error: overallError,
    log: currentLog,
    message: allSent
      ? `3-Month Post Evaluation email successfully dispatched to ${totalSupervisors} supervisor(s).`
      : (failCount > 0 ? `Failed to send email to ${failCount} supervisor(s): ${overallError}` : `Dispatched ${successCount} emails, ${skippedCount} previously sent.`)
  });
}

/**
 * Hourly / Daily Scheduled Task: Evaluates all completed training programmes.
 * Automatically dispatches 3-Month Post Evaluation emails to supervisors upon reaching the milestone.
 * Completely server-side and independent of the Admin Dashboard.
 */
function processAutomated3MonthPostEvaluationEmails(simulatedDate, testEmailRecipient) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return ok({ processed: 0, sent: 0, skipped: 0, errors: [] });

    ensureTrainingSheetColumns(tSheet);
    const trainings = sheetToJson(tSheet);
    if (!trainings || trainings.length === 0) return ok({ processed: 0, sent: 0, skipped: 0, errors: [] });

    const now = simulatedDate ? new Date(simulatedDate) : new Date();
    const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    let processedCount = 0;
    let sentCount = 0;
    let skippedCount = 0;
    let missingSupCount = 0;
    const errors = [];

    trainings.forEach(t => {
      const tId = String(t.ID || t.Code || '').trim();
      if (!tId) return;

      const stage = String(t.Stage || '').trim();
      const status = String(t.Status || '').trim();

      const startDateStr = t.StartDate;
      const endDateStr = t.EndDate || t.StartDate;
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

      const isCompleted = ['Training Completed', 'Evaluation Completed', 'Waiting for 3-Month Review', 'Programme Closed', 'Completed'].includes(stage) ||
                          status === 'Completed' ||
                          (todayMidnight.getTime() > compDate.getTime());
      if (!isCompleted) return;

      // 3-Month Milestone calculation
      const targetY = compDate.getFullYear() + Math.floor((compDate.getMonth() + 3) / 12);
      const targetM = (compDate.getMonth() + 3) % 12;
      const daysInTargetMonth = new Date(targetY, targetM + 1, 0).getDate();
      const targetDay = Math.min(compDate.getDate(), daysInTargetMonth);
      const threeMonthMidnight = new Date(targetY, targetM, targetDay);

      const daysUntil3Mo = Math.round((threeMonthMidnight.getTime() - todayMidnight.getTime()) / (1000 * 60 * 60 * 24));

      // Only process when 3-month milestone is reached or passed (daysUntil3Mo <= 0)
      if (daysUntil3Mo <= 0) {
        processedCount++;

        // Idempotency: Skip if already successfully sent
        const emailStatus = String(t.PostEvalEmailStatus || '').trim().toUpperCase();
        if (emailStatus === 'SENT' && !testEmailRecipient) {
          skippedCount++;
          return;
        }

        try {
          const res = sendSupervisorPostEvaluationEmailsForTraining(tId, {
            simulatedDate: simulatedDate,
            testEmailRecipient: testEmailRecipient
          });
          const resData = (res && typeof res === 'object') ? (res.data || res) : {};
          if (resData.status === 'SENT') {
            sentCount++;
          } else if (resData.status === 'SUPERVISOR_MISSING') {
            missingSupCount++;
          } else if (resData.status === 'FAILED') {
            errors.push(`${tId}: ${resData.error || 'Email dispatch failed'}`);
          }
        } catch(autoErr) {
          errors.push(`${tId}: ${autoErr.message}`);
        }
      }
    });

    Logger.log(`processAutomated3MonthPostEvaluationEmails complete: Processed ${processedCount}, Sent ${sentCount}, Skipped ${skippedCount}, Missing Sup ${missingSupCount}, Errors: ${errors.length}`);
    return ok({
      processedCount: processedCount,
      sentCount: sentCount,
      skippedCount: skippedCount,
      missingSupCount: missingSupCount,
      errors: errors
    });

  } catch(e) {
    Logger.log('processAutomated3MonthPostEvaluationEmails error: ' + e.message);
    return err('Failed to execute automated post evaluation emails: ' + e.message);
  }
}

/**
 * Developer / Test Mode: Simulates 3-Month Post Evaluation automation for any date.
 * Safely redirects emails to a test recipient to prevent accidental live supervisor messaging.
 * 
 * @param {string} trainingId - Specific training ID to test (or empty for all)
 * @param {string} simulatedDateStr - e.g. '2026-11-19'
 * @param {string} testEmailRecipient - Email address to receive redirected test emails
 */
function runPostEvaluationAutomationTest(trainingId, simulatedDateStr, testEmailRecipient) {
  try {
    const simDate = simulatedDateStr ? new Date(simulatedDateStr) : new Date();
    if (isNaN(simDate.getTime())) return err('Invalid simulated date format. Use YYYY-MM-DD.');

    const testEmail = testEmailRecipient || getConfigProperty('ADMIN_EMAILS', '') || 'test@apollofood.com.my';

    if (trainingId) {
      return sendSupervisorPostEvaluationEmailsForTraining(trainingId, {
        simulatedDate: simDate,
        testEmailRecipient: testEmail,
        forceResend: true
      });
    } else {
      return processAutomated3MonthPostEvaluationEmails(simDate, testEmail);
    }
  } catch(e) {
    return err('runPostEvaluationAutomationTest error: ' + e.message);
  }
}

/**
 * API: Manually trigger or retry sending Post Evaluation emails from Admin Dashboard.
 */
function triggerPostEvaluationEmailSend(trainingId) {
  try {
    if (!trainingId) return err('Training ID is required.');
    return sendSupervisorPostEvaluationEmailsForTraining(trainingId, { forceResend: false });
  } catch(e) {
    return err('Failed to send Post Evaluation email: ' + e.message);
  }
}

/**
 * Automations & Trigger Setup: Hourly Trigger for Apps Script
 */
function setupAutomatedLifecycleTrigger() {
  try {
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(tr => {
      if (tr.getHandlerFunction() === 'runTrainingLifecycleAutomation' || tr.getHandlerFunction() === 'cronCheck3MonthPostEvalNotifications') {
        ScriptApp.deleteTrigger(tr);
      }
    });

    ScriptApp.newTrigger('runTrainingLifecycleAutomation')
      .timeBased()
      .everyHours(1)
      .create();

    Logger.log('Created hourly trigger for runTrainingLifecycleAutomation');
    return ok({
      message: 'Hourly trigger configured successfully for runTrainingLifecycleAutomation.',
      frequency: 'HOURLY',
      count: 1
    });
  } catch(e) {
    Logger.log('setupAutomatedLifecycleTrigger error: ' + e.message);
    return err('Failed to configure trigger: ' + e.message);
  }
}

/**
 * Inspects and returns the status of Apps Script time-driven triggers.
 */
function getAutomationTriggerStatus() {
  try {
    const allTriggers = ScriptApp.getProjectTriggers();
    const lifecycleTriggers = allTriggers.filter(tr => tr.getHandlerFunction() === 'runTrainingLifecycleAutomation');

    return ok({
      exists: lifecycleTriggers.length > 0,
      count: lifecycleTriggers.length,
      frequency: 'HOURLY',
      hasDuplicates: lifecycleTriggers.length > 1,
      handler: 'runTrainingLifecycleAutomation'
    });
  } catch(e) {
    return err('Failed to inspect project triggers: ' + e.message);
  }
}

function runTrainingLifecycleAutomation() {
  try {
    autoUpdateTrainingLifecycleStages();
  } catch(e) {
    Logger.log('autoUpdateTrainingLifecycleStages error: ' + e.message);
  }

  try {
    processAutomated3MonthPostEvaluationEmails();
  } catch(e) {
    Logger.log('processAutomated3MonthPostEvaluationEmails error: ' + e.message);
  }
}

// --- Internal: auto-advance stage ----------------------------------------------
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
