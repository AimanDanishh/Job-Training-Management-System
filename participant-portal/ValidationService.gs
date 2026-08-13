/**
 * ValidationService.gs — Server-Side Security & Data Validation Engine for Public Participant Portal
 * 
 * Enforces strict validation rules on the server for all public submissions.
 * Rejects invalid requests before writing to the database.
 */

// ─── 1. Check Training Existence ────────────────────────────────────────────────
function getValidTraining(trainingId) {
  if (!trainingId || String(trainingId).trim() === '') {
    return { valid: false, message: 'Training ID is missing or invalid.' };
  }
  const cleanId = String(trainingId).trim();
  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) {
    if (!getConfigProperty('SPREADSHEET_ID', '')) {
      return { valid: false, message: 'Spreadsheet ID not configured. Please set SPREADSHEET_ID in Apps Script Project Settings.' };
    }
    return { valid: false, message: 'Trainings database sheet unavailable.' };
  }

  const rows = sheetToJson(tSheet);
  const training = rows.find(r => String(r.ID || r.TrainingID || '').trim() === cleanId);

  if (!training) {
    return { valid: false, message: `Training programme (${cleanId}) does not exist.` };
  }
  return { valid: true, training: training };
}

function getValidEmployee(employeeId) {
  if (!employeeId || String(employeeId).trim() === '') {
    return { valid: false, message: 'Employee ID is required.' };
  }
  const cleanEmpId = String(employeeId).trim();
  
  // 1. Primary Lookup: Master Employees Sheet
  const empSheet = getSheet(SHEET_NAMES.employees);
  if (empSheet) {
    const rows = sheetToJson(empSheet);
    const emp = rows.find(r => isSameEmployeeId(r.ID || r.EmployeeID || r.EmployeeNo || '', cleanEmpId));
    if (emp) {
      return { valid: true, employee: emp };
    }
  }

  // 2. Fallback Lookup: TrainingParticipants Sheet
  try {
    const tpSheet = getSheet(SHEET_NAMES.trainingParticipants);
    if (tpSheet) {
      const tpRows = sheetToJson(tpSheet);
      const tpEmp = tpRows.find(r => isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId));
      if (tpEmp) {
        return {
          valid: true,
          employee: {
            ID: tpEmp.EmployeeID || String(employeeId).trim(),
            Name: tpEmp.EmployeeName || String(employeeId).trim(),
            Department: tpEmp.CostCentre || tpEmp.Department || ''
          }
        };
      }
    }
  } catch (e) {
    Logger.log('Fallback employee lookup error: ' + e.message);
  }

  if (!empSheet && !getConfigProperty('SPREADSHEET_ID', '')) {
    return { valid: false, message: 'Spreadsheet ID not configured. Please set SPREADSHEET_ID in Apps Script Project Settings.' };
  }

  return { valid: false, message: `Employee ID (${String(employeeId).trim()}) is not registered in the system.` };
}

// ─── 3. Check Employee Enrollment for Training ──────────────────────────────────
function validateParticipantEnrollment(trainingId, employeeId) {
  const cleanTId   = String(trainingId || '').trim();
  const cleanEmpId = String(employeeId || '').trim();

  const ss = getTrainingDataSpreadsheet(cleanTId);
  const tpSheet = ss ? ss.getSheetByName('TrainingParticipants') : null;
  if (!tpSheet) {
    // Fallback: If TrainingParticipants sheet isn't populated, permit lookup in Employees
    return getValidEmployee(employeeId);
  }

  const tpRows = sheetToJson(tpSheet);
  const enrolled = tpRows.find(r => 
    isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId)
  );

  if (!enrolled) {
    return { 
      valid: false, 
      message: `Employee (${String(employeeId).trim()}) is not registered as an official participant for this training programme.` 
    };
  }

  return { valid: true, participant: enrolled };
}

// ─── 4. Public Attendance Validation ─────────────────────────────────────────────
function validatePublicAttendance(sessionId, employeeId) {
  try {
    if (!sessionId || String(sessionId).trim() === '') {
      return { valid: false, message: 'Session ID is required for check-in.' };
    }
    if (!employeeId || String(employeeId).trim() === '') {
      return { valid: false, message: 'Employee ID is required for check-in.' };
    }

    const cleanSessionId = String(sessionId).trim();
    const cleanEmpId     = String(employeeId).trim();

    // A. Validate Session Existence & Status
    const found = findTrainingBySessionId(cleanSessionId);
    if (!found || !found.session) {
      return { valid: false, message: 'Invalid session ID. Session does not exist.' };
    }

    const session = found.session;
    const sSheet  = found.sessionSheet;
    const qrStatus = String(session.QRStatus || 'Active').trim();

    // Check if SessionDate + EndTime has passed
    if (session.SessionDate) {
      try {
        const currentDate = new Date();
        const endTimeStr = session.EndTime || '17:00';
        const sDate = new Date(session.SessionDate);
        if (!isNaN(sDate.getTime())) {
          const [hours, minutes] = String(endTimeStr).split(':').map(Number);
          sDate.setHours(hours || 17, minutes || 0, 0, 0);

          if (currentDate > sDate) {
            if (qrStatus === 'Active' && sSheet) {
              const row = findRowById(sSheet, cleanSessionId);
              if (row !== -1) sSheet.getRange(row, 9).setValue('Expired');
            }
            return { valid: false, message: `Attendance registration for this session is closed. Session ended at ${endTimeStr} on ${formatDate(sDate)}.` };
          }
        }
      } catch (err) {
        Logger.log('Session time check error: ' + err.message);
      }
    }

    if (qrStatus === 'Expired') {
      return { valid: false, message: 'Attendance registration for this session is closed (Expired).' };
    }
    if (qrStatus === 'Inactive') {
      return { valid: false, message: 'This training session is currently inactive.' };
    }
    if (qrStatus === 'Scheduled' || qrStatus === 'Draft') {
      return { valid: false, message: 'Attendance check-in has not opened yet for this session.' };
    }

    const trainingId = session.TrainingID;

    // B. Validate Training Existence
    const tCheck = getValidTraining(trainingId);
    if (!tCheck.valid) return tCheck;

    // C. Validate Employee Existence
    const empCheck = getValidEmployee(cleanEmpId);
    if (!empCheck.valid) return empCheck;

    // D. Validate Employee Enrollment for Training
    const enrollCheck = validateParticipantEnrollment(trainingId, cleanEmpId);
    if (!enrollCheck.valid) return enrollCheck;

    // E. Prevent Duplicate Attendance Submission
    const ss = found.spreadsheet;
    const attSheet = ss ? ss.getSheetByName('Attendance') : null;
    if (attSheet) {
      const attRows = sheetToJson(attSheet);
      const duplicate = attRows.find(r => {
        const rSessionId = String(r.SessionID || '').trim();
        const rEmpNo     = String(r.EmployeeNo || r.EmployeeID || '').trim();
        return rSessionId === cleanSessionId && isSameEmployeeId(rEmpNo, cleanEmpId);
      });

      if (duplicate) {
        return {
          valid: false,
          duplicate: true,
          scanTime: duplicate.ScanTime || duplicate.EditedAt || 'Earlier',
          message: `Attendance has already been recorded for Employee ID (${cleanEmpId}) for session ${cleanSessionId}.`,
          employee: empCheck.employee,
          session: session
        };
      }
    }

    return {
      valid: true,
      session: session,
      training: tCheck.training,
      employee: empCheck.employee
    };
  } catch (e) {
    Logger.log('validatePublicAttendance error: ' + e.message);
    return { valid: false, message: 'Attendance validation failure: ' + e.message };
  }
}

// ─── 4.1 Attendance Eligibility Check for Evaluation ──────────────────────────────
function checkEmployeeAttendanceEligibility(trainingId, employeeId) {
  if (!trainingId || String(trainingId).trim() === '') {
    return { eligible: false, message: 'Training ID is required.' };
  }
  if (!employeeId || String(employeeId).trim() === '') {
    return { eligible: false, message: 'Employee ID is required.' };
  }

  const cleanTId   = String(trainingId).trim();
  const cleanEmpId = String(employeeId).trim().toLowerCase();

  const ss = getTrainingDataSpreadsheet(cleanTId);
  if (!ss) {
    return { eligible: false, message: 'Training data sheet is unavailable.' };
  }

  const attSheet = ss.getSheetByName('Attendance');
  if (!attSheet) {
    return {
      eligible: false,
      message: 'Our records show that you do not have an attendance record for this training.'
    };
  }

  const attRows = sheetToJson(attSheet);
  const validRecord = attRows.find(r => {
    const rEmpNo = String(r.EmployeeNo || r.EmployeeID || r.ID || '').trim().toLowerCase();
    const status = String(r.Status || 'Present').trim().toLowerCase();
    return rEmpNo === cleanEmpId && status !== 'absent';
  });

  if (!validRecord) {
    return {
      eligible: false,
      message: 'Our records show that you do not have an attendance record for this training.'
    };
  }

  return { eligible: true, attendanceRecord: validRecord };
}

// ─── 4.2 Endpoint Verification Helpers ─────────────────────────────────────────
function verifyEmployeeForAttendance(sessionId, employeeId) {
  try {
    const validation = validatePublicAttendance(sessionId, employeeId);
    if (!validation.valid) {
      if (validation.duplicate) {
        return ok({
          alreadyRecorded: true,
          scanTime: validation.scanTime || 'Earlier',
          employee: validation.employee,
          session: validation.session,
          message: `Attendance has already been recorded for ${validation.employee ? validation.employee.Name : employeeId} (${employeeId}) for this session.`
        });
      }
      return err(validation.message);
    }

    return ok({
      valid: true,
      alreadyRecorded: false,
      employee: validation.employee,
      session: validation.session,
      training: validation.training
    });
  } catch (e) {
    return err('Verification failed: ' + e.message);
  }
}

function verifyEmployeeForEvaluation(trainingId, employeeId) {
  try {
    if (!trainingId || String(trainingId).trim() === '') {
      return err('Training ID parameter missing from URL link.');
    }
    if (!employeeId || String(employeeId).trim() === '') {
      return err('Employee ID is required.');
    }

    const cleanTId   = String(trainingId).trim();
    const cleanEmpId = String(employeeId).trim();

    // A. Validate Training Existence
    const tCheck = getValidTraining(cleanTId);
    if (!tCheck.valid) return err(tCheck.message);

    // B. Validate Employee Existence
    const empCheck = getValidEmployee(cleanEmpId);
    if (!empCheck.valid) return err(empCheck.message);

    // C. Validate Employee Enrollment
    const enrollCheck = validateParticipantEnrollment(cleanTId, cleanEmpId);
    if (!enrollCheck.valid) return err(enrollCheck.message);

    // D. Validate Attendance Eligibility (CRITICAL REQUIREMENT: Must have attended at least 1 session)
    const attCheck = checkEmployeeAttendanceEligibility(cleanTId, cleanEmpId);
    if (!attCheck.eligible) {
      return ok({
        eligible: false,
        noAttendance: true,
        message: attCheck.message,
        training: tCheck.training,
        employee: empCheck.employee
      });
    }

    // E. Check if Evaluation Already Submitted
    const ss = getTrainingDataSpreadsheet(cleanTId);
    const evalSheet = ss ? ss.getSheetByName('TrainingEval') : null;
    if (evalSheet) {
      const evalRows = sheetToJson(evalSheet);
      const duplicate = evalRows.find(r => 
        String(r.EmployeeID || r.EmployeeNo || '').trim().toLowerCase() === cleanEmpId.toLowerCase()
      );

      if (duplicate) {
        return ok({
          eligible: false,
          alreadySubmitted: true,
          submittedAt: duplicate.SubmittedAt || 'Earlier',
          message: 'You have already completed the evaluation for this training.',
          training: tCheck.training,
          employee: empCheck.employee
        });
      }
    }

    return ok({
      eligible: true,
      valid: true,
      training: tCheck.training,
      employee: empCheck.employee
    });
  } catch (e) {
    return err('Evaluation verification failed: ' + e.message);
  }
}

function verifyEvaluatorByEmployeeId(evaluatorEmployeeId, trainingId) {
  try {
    if (!evaluatorEmployeeId || String(evaluatorEmployeeId).trim() === '') {
      return err('Supervisor / PIC Employee ID is required.');
    }
    const cleanEvalEmpId = String(evaluatorEmployeeId).trim();

    // 1. Verify Evaluator Employee Record
    const empCheck = getValidEmployee(cleanEvalEmpId);
    if (!empCheck.valid) {
      return err(`Evaluator Employee ID (${cleanEvalEmpId}) is not registered in the system.`);
    }

    const evaluator = empCheck.employee;
    const evalName  = evaluator.Name || evaluator.EmployeeName || cleanEvalEmpId;
    const evalDept  = evaluator.CostCentre || evaluator.Department || 'Supervisor / Manager';
    const evalEmail = evaluator.Email || evaluator.EmailAddress || '';

    // 2. Fetch pending participants needing 3-Month Post Evaluation
    let trnIdFilter = String(trainingId || '').trim();
    let pendingList = [];
    let completedCount = 0;
    let targetTraining = null;

    if (trnIdFilter) {
      const tCheck = getValidTraining(trnIdFilter);
      if (tCheck.valid) targetTraining = tCheck.training;

      const ss = getTrainingDataSpreadsheet(trnIdFilter);
      if (ss) {
        const tpSheet   = ss.getSheetByName('TrainingParticipants');
        const postSheet = ss.getSheetByName('PostEval');

        const tpRows   = tpSheet ? sheetToJson(tpSheet) : [];
        const postRows = postSheet ? sheetToJson(postSheet) : [];
        const completedEmpIds = postRows.map(r => String(r.EmployeeID || '').trim().toLowerCase());

        tpRows.forEach(p => {
          const empId = String(p.EmployeeID || p.ID || '').trim();
          if (!empId) return;
          if (completedEmpIds.includes(empId.toLowerCase())) {
            completedCount++;
          } else {
            pendingList.push({
              EmployeeID: empId,
              Name: p.EmployeeName || p.Name || empId,
              Department: p.CostCentre || p.Department || '',
              Position: p.Position || p.JobTitle || 'Participant'
            });
          }
        });
      }
    } else {
      const tSheet = getSheet(SHEET_NAMES.trainings);
      if (tSheet) {
        const trainings = sheetToJson(tSheet);
        if (trainings.length > 0) {
          targetTraining = trainings[0];
          trnIdFilter = targetTraining.ID;
          const ss = getTrainingDataSpreadsheet(trnIdFilter);
          if (ss) {
            const tpSheet   = ss.getSheetByName('TrainingParticipants');
            const postSheet = ss.getSheetByName('PostEval');

            const tpRows   = tpSheet ? sheetToJson(tpSheet) : [];
            const postRows = postSheet ? sheetToJson(postSheet) : [];
            const completedEmpIds = postRows.map(r => String(r.EmployeeID || '').trim().toLowerCase());

            tpRows.forEach(p => {
              const empId = String(p.EmployeeID || p.ID || '').trim();
              if (!empId) return;
              if (completedEmpIds.includes(empId.toLowerCase())) {
                completedCount++;
              } else {
                pendingList.push({
                  EmployeeID: empId,
                  Name: p.EmployeeName || p.Name || empId,
                  Department: p.CostCentre || p.Department || '',
                  Position: p.Position || p.JobTitle || 'Participant'
                });
              }
            });
          }
        }
      }
    }

    return ok({
      evaluator: {
        EmployeeID: cleanEvalEmpId,
        Name: evalName,
        Department: evalDept,
        Email: evalEmail
      },
      training: targetTraining,
      pendingParticipants: pendingList,
      completedCount: completedCount
    });

  } catch (e) {
    return err('Evaluator verification error: ' + e.message);
  }
}

// ─── 5. Public Training Evaluation Validation ────────────────────────────────────
function validatePublicEvaluation(trainingId, employeeId) {
  try {
    if (!trainingId || String(trainingId).trim() === '') {
      return { valid: false, message: 'Training ID is required for evaluation.' };
    }
    if (!employeeId || String(employeeId).trim() === '') {
      return { valid: false, message: 'Employee ID is required for evaluation.' };
    }

    const cleanTId   = String(trainingId).trim();
    const cleanEmpId = String(employeeId).trim();

    // A. Validate Training Existence
    const tCheck = getValidTraining(cleanTId);
    if (!tCheck.valid) return tCheck;

    // B. Validate Employee Existence
    const empCheck = getValidEmployee(cleanEmpId);
    if (!empCheck.valid) return empCheck;

    // C. Validate Employee Enrollment
    const enrollCheck = validateParticipantEnrollment(cleanTId, cleanEmpId);
    if (!enrollCheck.valid) return enrollCheck;

    // D. Attendance Verification: Disallow evaluation for participants without attendance
    const attCheck = checkEmployeeAttendanceEligibility(cleanTId, cleanEmpId);
    if (!attCheck.eligible) {
      return {
        valid: false,
        message: attCheck.message
      };
    }

    // E. Prevent Duplicate Evaluation Submission
    const ss = getTrainingDataSpreadsheet(cleanTId);
    const evalSheet = ss ? ss.getSheetByName('TrainingEval') : null;
    if (evalSheet) {
      const evalRows = sheetToJson(evalSheet);
      const duplicate = evalRows.find(r => 
        String(r.EmployeeID || '').trim().toLowerCase() === cleanEmpId.toLowerCase()
      );

      if (duplicate) {
        return {
          valid: false,
          message: `You have already completed the evaluation for this training programme (${tCheck.training.Code || cleanTId}).`
        };
      }
    }

    return {
      valid: true,
      training: tCheck.training,
      employee: empCheck.employee
    };
  } catch (e) {
    Logger.log('validatePublicEvaluation error: ' + e.message);
    return { valid: false, message: 'Evaluation validation failure: ' + e.message };
  }
}

// ─── 6. Public Post-Evaluation Validation (3-Month Supervisor Review) ────────────
function validatePublicPostEvaluation(trainingId, employeeId, token) {
  try {
    const cleanTId   = String(trainingId || '').trim();
    const cleanEmpId = String(employeeId || '').trim();
    const cleanToken = String(token || '').trim();

    if (!cleanTId && !cleanToken) {
      return { valid: false, message: 'Training identifier or review token is required.' };
    }
    if (!cleanEmpId && !cleanToken) {
      return { valid: false, message: 'Employee ID is required.' };
    }

    let effectiveTId   = cleanTId;
    let effectiveEmpId = cleanEmpId;

    if (cleanToken && cleanToken.includes('_')) {
      const parts = cleanToken.split('_');
      if (parts.length >= 3) {
        effectiveTId   = parts[1];
        effectiveEmpId = parts[2];
      }
    }

    // A. Validate Training Existence
    const tCheck = getValidTraining(effectiveTId);
    if (!tCheck.valid) return tCheck;

    // B. Validate Employee Existence
    const empCheck = getValidEmployee(effectiveEmpId);
    if (!empCheck.valid) return empCheck;

    // C. Validate Employee Enrollment
    const enrollCheck = validateParticipantEnrollment(effectiveTId, effectiveEmpId);
    if (!enrollCheck.valid) return enrollCheck;

    // D. Prevent Duplicate Post-Evaluation Submission
    const ss = getTrainingDataSpreadsheet(effectiveTId);
    const postSheet = ss ? ss.getSheetByName('PostEval') : null;
    if (postSheet) {
      const postRows = sheetToJson(postSheet);
      const duplicate = postRows.find(r =>
        String(r.EmployeeID || '').trim().toLowerCase() === effectiveEmpId.toLowerCase()
      );

      if (duplicate) {
        return {
          valid: false,
          message: `A 3-Month Post-Training Evaluation has already been submitted for employee (${effectiveEmpId}).`
        };
      }
    }

    return {
      valid: true,
      trainingId: effectiveTId,
      employeeId: effectiveEmpId,
      training: tCheck.training,
      employee: empCheck.employee
    };
  } catch (e) {
    Logger.log('validatePublicPostEvaluation error: ' + e.message);
    return { valid: false, message: 'Post-evaluation validation error: ' + e.message };
  }
}

/**
 * Public helper: Retrieve employee details by Employee ID for auto-populating Name & Department
 */
function getEmployeeDetails(employeeId) {
  try {
    const res = getValidEmployee(employeeId);
    if (!res.valid) return err(res.message);
    return ok(res.employee);
  } catch (e) {
    return err(e.message);
  }
}
