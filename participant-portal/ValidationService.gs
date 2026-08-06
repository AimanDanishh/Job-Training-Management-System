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
  const cleanEmpId = String(employeeId).trim().toLowerCase();
  
  // 1. Primary Lookup: Master Employees Sheet
  const empSheet = getSheet(SHEET_NAMES.employees);
  if (empSheet) {
    const rows = sheetToJson(empSheet);
    const emp = rows.find(r => String(r.ID || r.EmployeeID || r.EmployeeNo || '').trim().toLowerCase() === cleanEmpId);
    if (emp) {
      return { valid: true, employee: emp };
    }
  }

  // 2. Fallback Lookup: TrainingParticipants Sheet
  try {
    const tpSheet = getSheet(SHEET_NAMES.trainingParticipants);
    if (tpSheet) {
      const tpRows = sheetToJson(tpSheet);
      const tpEmp = tpRows.find(r => String(r.EmployeeID || r.EmployeeNo || r.ID || '').trim().toLowerCase() === cleanEmpId);
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
  const cleanEmpId = String(employeeId || '').trim().toLowerCase();

  const ss = getTrainingDataSpreadsheet(cleanTId);
  const tpSheet = ss ? ss.getSheetByName('TrainingParticipants') : null;
  if (!tpSheet) {
    // Fallback: If TrainingParticipants sheet isn't populated, permit lookup in Employees
    return getValidEmployee(employeeId);
  }

  const tpRows = sheetToJson(tpSheet);
  const enrolled = tpRows.find(r => 
    String(r.EmployeeID || r.EmployeeNo || r.ID || '').trim().toLowerCase() === cleanEmpId
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
        const formattedDate = formatDate(new Date(session.SessionDate));
        const [hours, minutes] = String(endTimeStr).split(':').map(Number);
        const sessionEndDateTime = new Date(formattedDate);
        sessionEndDateTime.setHours(hours || 17, minutes || 0, 0, 0);

        if (currentDate > sessionEndDateTime) {
          if (qrStatus === 'Active' && sSheet) {
            const row = findRowById(sSheet, cleanSessionId);
            if (row !== -1) sSheet.getRange(row, 9).setValue('Expired');
          }
          return { valid: false, message: `Attendance registration for this session is closed. Session ended at ${endTimeStr} on ${formattedDate}.` };
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
        const rEmpNo     = String(r.EmployeeNo || r.EmployeeID || '').trim().toLowerCase();
        return rSessionId === cleanSessionId && rEmpNo === cleanEmpId.toLowerCase();
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

    // D. 2-Week Submission Deadline Enforcement (14 days post completion)
    const endDateStr = tCheck.training.EndDate || tCheck.training.StartDate;
    if (endDateStr) {
      try {
        const endDate = new Date(endDateStr);
        if (!isNaN(endDate.getTime())) {
          endDate.setHours(23, 59, 59, 999);
          const now = new Date();
          const diffMs = now.getTime() - endDate.getTime();
          const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
          if (diffDays > 14) {
            return {
              valid: false,
              message: `Participant evaluation must be submitted within 2 weeks (14 days) after course completion. The deadline for this training passed on ${formatDate(new Date(endDate.getTime() + 14 * 24 * 60 * 60 * 1000))}.`
            };
          }
        }
      } catch (errDate) {
        Logger.log('Date validation check error: ' + errDate.message);
      }
    }

    // E. Attendance Verification: Disallow evaluation for absent participants
    const ss = getTrainingDataSpreadsheet(cleanTId);
    const attSheet = ss ? ss.getSheetByName('Attendance') : null;
    if (attSheet) {
      const attRows = sheetToJson(attSheet);
      const attRecord = attRows.find(r => 
        String(r.EmployeeNo || r.EmployeeID || '').trim().toLowerCase() === cleanEmpId.toLowerCase()
      );

      // If record exists and status is Absent
      if (attRecord && String(attRecord.Status || '').toLowerCase() === 'absent') {
        return {
          valid: false,
          message: `Training evaluation is disallowed for participant (${cleanEmpId}) who was marked ABSENT for this training.`
        };
      }
    }

    // F. Prevent Duplicate Evaluation Submission
    const evalSheet = ss ? ss.getSheetByName('TrainingEval') : null;
    if (evalSheet) {
      const evalRows = sheetToJson(evalSheet);
      const duplicate = evalRows.find(r => 
        String(r.EmployeeID || '').trim().toLowerCase() === cleanEmpId.toLowerCase()
      );

      if (duplicate) {
        return {
          valid: false,
          message: `You have already submitted an evaluation for this training programme (${tCheck.training.Code || cleanTId}).`
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
