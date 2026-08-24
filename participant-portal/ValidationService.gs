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
  const cleanId = String(trainingId).trim().toLowerCase();
  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) {
    if (!getConfigProperty('SPREADSHEET_ID', '')) {
      return { valid: false, message: 'Spreadsheet ID not configured. Please set SPREADSHEET_ID in Apps Script Project Settings.' };
    }
    return { valid: false, message: 'Trainings database sheet unavailable.' };
  }

  const rows = sheetToJson(tSheet);
  const training = rows.find(r => {
    const id = String(r.ID || '').trim().toLowerCase();
    const code = String(r.Code || '').trim().toLowerCase();
    const tId = String(r.TrainingID || '').trim().toLowerCase();
    return id === cleanId || code === cleanId || tId === cleanId;
  });

  if (!training) {
    return { valid: false, message: `Training programme (${cleanId}) does not exist.` };
  }
  return { valid: true, training: training };
}

function getValidEmployee(employeeId, trainingId) {
  if (!employeeId || String(employeeId).trim() === '') {
    return { valid: false, message: 'Employee identifier is required.' };
  }
  const cleanInput = String(employeeId).trim();
  const cleanInputLower = cleanInput.toLowerCase();
  const cleanTId   = trainingId ? String(trainingId).trim() : '';

  const getRowEmpId = (r) => {
    return r.ID || r.EmployeeID || r.EmployeeNo || r.EmpID || r.StaffID || r.StaffNo || r.EmpNo ||
           r['Employee ID'] || r['Employee No'] || r['Staff ID'] || r['Staff No'] || r['Emp ID'] || r['Emp No'] ||
           r['Employee_ID'] || r['Staff_ID'] || r['Badge No'] || r['BadgeNo'] || r['No'] || r['No.'] || r['Employee Number'] || '';
  };

  const getRowEmail = (r) => {
    return r.Email || r.EmailAddress || r['Email Address'] || r['Company Email'] || r['Work Email'] || r['User Email'] || r['HOD Email'] || r['HR Email'] || '';
  };

  const getRowName = (r) => {
    return r.Name || r.EmployeeName || r['Employee Name'] || r['Staff Name'] || r['HOD Name'] || r['HR Name'] || r.FullName || r['Full Name'] || '';
  };

  const getRowDept = (r) => {
    return r.Department || r.CostCentre || r['Cost Centre'] || r.Dept || r.Section || '';
  };

  const getRowPos = (r) => {
    return r.Position || r.JobTitle || r.PositionTitle || r['Position Title'] || r['Job Title'] || r.Designation || '';
  };

  // 1. Primary Lookup: Per-Training Spreadsheet (Participants & Assigned Supervisors)
  if (cleanTId) {
    try {
      const ss = getTrainingDataSpreadsheet(cleanTId);
      const tpSheet = ss ? (ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants') || ss.getSheetByName('ParticipantList')) : null;
      if (tpSheet) {
        const tpRows = sheetToJson(tpSheet);
        
        // 1a. Check Participant columns
        const tpEmp = tpRows.find(r => 
          isSameEmployeeId(getRowEmpId(r), cleanInput) ||
          (getRowEmail(r) && String(getRowEmail(r)).trim().toLowerCase() === cleanInputLower) ||
          (getRowName(r) && String(getRowName(r)).toLowerCase().trim() === cleanInputLower)
        );
        if (tpEmp) {
          return {
            valid: true,
            employee: {
              ID: getRowEmpId(tpEmp) || cleanInput,
              Name: getRowName(tpEmp) || cleanInput,
              Department: getRowDept(tpEmp),
              Position: getRowPos(tpEmp) || 'Participant',
              Email: getRowEmail(tpEmp)
            }
          };
        }

        // 1b. Check Assigned Supervisor columns
        const supEmp = tpRows.find(r => 
          isSameEmployeeId(r.SupervisorID || '', cleanInput) ||
          (r.SupervisorEmail && String(r.SupervisorEmail).trim().toLowerCase() === cleanInputLower) ||
          (r.SupervisorName && String(r.SupervisorName).trim().toLowerCase() === cleanInputLower)
        );
        if (supEmp) {
          return {
            valid: true,
            employee: {
              ID: supEmp.SupervisorID || cleanInput,
              Name: supEmp.SupervisorName || cleanInput,
              Email: supEmp.SupervisorEmail || (cleanInput.includes('@') ? cleanInput : ''),
              Department: getRowDept(supEmp) || 'Supervisor / PIC',
              Position: 'Supervisor'
            }
          };
        }
      }
    } catch (e) {
      Logger.log('Per-training lookup error: ' + e.message);
    }

    // Check ParticipantList JSON on training record
    try {
      const tCheck = getValidTraining(cleanTId);
      if (tCheck.valid && tCheck.training) {
        const rawJson = tCheck.training.ParticipantList || tCheck.training.participants;
        if (rawJson) {
          const list = Array.isArray(rawJson) ? rawJson : JSON.parse(rawJson);
          if (Array.isArray(list)) {
            const m = list.find(p => 
              isSameEmployeeId(getRowEmpId(p) || p.SupervisorID || '', cleanInput) ||
              (p.SupervisorEmail && String(p.SupervisorEmail).trim().toLowerCase() === cleanInputLower) ||
              (p.SupervisorName && String(p.SupervisorName).trim().toLowerCase() === cleanInputLower) ||
              (getRowEmail(p) && String(getRowEmail(p)).trim().toLowerCase() === cleanInputLower) ||
              (getRowName(p) && String(getRowName(p)).trim().toLowerCase() === cleanInputLower)
            );
            if (m) {
              return {
                valid: true,
                employee: {
                  ID: getRowEmpId(m) || m.SupervisorID || cleanInput,
                  Name: getRowName(m) || m.SupervisorName || cleanInput,
                  Department: getRowDept(m),
                  Position: getRowPos(m) || '',
                  Email: getRowEmail(m) || m.SupervisorEmail || ''
                }
              };
            }
          }
        }
      }
    } catch(jErr) {}
  }

  // 2. Employee Master Directory lookup (Checking all possible employee sheets)
  const masterSheetTabs = [
    SHEET_NAMES.employees,
    'Employees',
    'For IT',
    'FOR IT',
    'Staff',
    'Staff List',
    'Employee',
    'Master Employees',
    'Cost Centre',
    'HOD email',
    'HOD Email',
    'HR email',
    'HR Email',
    'Csuite email',
    'C-Suite email'
  ];

  for (const tabName of masterSheetTabs) {
    if (!tabName) continue;
    try {
      const empSheet = getSheet(tabName);
      if (empSheet) {
        const rows = sheetToJson(empSheet);
        const emp = rows.find(r => 
          isSameEmployeeId(getRowEmpId(r), cleanInput) ||
          (getRowEmail(r) && String(getRowEmail(r)).trim().toLowerCase() === cleanInputLower) ||
          (getRowName(r) && String(getRowName(r)).trim().toLowerCase() === cleanInputLower)
        );
        if (emp) {
          return {
            valid: true,
            employee: {
              ID: getRowEmpId(emp) || cleanInput,
              Name: getRowName(emp) || cleanInput,
              Department: getRowDept(emp) || 'Company Staff',
              Position: getRowPos(emp) || 'Employee',
              Email: getRowEmail(emp) || (cleanInput.includes('@') ? cleanInput : '')
            }
          };
        }
      }
    } catch(sheetErr) {
      Logger.log(`Error checking sheet tab ${tabName}: ` + sheetErr.message);
    }
  }

  // 3. Fallback: Search all per-training participant rosters for assigned supervisor / participant
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet) {
      const allTrainings = sheetToJson(tSheet);
      for (const t of allTrainings) {
        const tId = t.ID || t.TrainingID || t.Code;
        if (!tId) continue;
        const ss = getTrainingDataSpreadsheet(tId);
        if (!ss) continue;
        const tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
        if (!tpSheet) continue;
        const tpRows = sheetToJson(tpSheet);
        const foundRow = tpRows.find(r => 
          isSameEmployeeId(r.SupervisorID || '', cleanInput) ||
          isSameEmployeeId(getRowEmpId(r), cleanInput) ||
          (r.SupervisorEmail && String(r.SupervisorEmail).trim().toLowerCase() === cleanInputLower) ||
          (r.SupervisorName && String(r.SupervisorName).trim().toLowerCase() === cleanInputLower) ||
          (getRowEmail(r) && String(getRowEmail(r)).trim().toLowerCase() === cleanInputLower) ||
          (getRowName(r) && String(getRowName(r)).trim().toLowerCase() === cleanInputLower)
        );
        if (foundRow) {
          const isSup = isSameEmployeeId(foundRow.SupervisorID || '', cleanInput) ||
                        (foundRow.SupervisorEmail && String(foundRow.SupervisorEmail).trim().toLowerCase() === cleanInputLower) ||
                        (foundRow.SupervisorName && String(foundRow.SupervisorName).trim().toLowerCase() === cleanInputLower);
          return {
            valid: true,
            employee: {
              ID: isSup ? (foundRow.SupervisorID || cleanInput) : (getRowEmpId(foundRow) || cleanInput),
              Name: isSup ? (foundRow.SupervisorName || cleanInput) : (getRowName(foundRow) || cleanInput),
              Department: getRowDept(foundRow) || (isSup ? 'Supervisor / Evaluator' : ''),
              Position: isSup ? 'Supervisor' : (getRowPos(foundRow) || 'Participant'),
              Email: isSup ? (foundRow.SupervisorEmail || '') : getRowEmail(foundRow)
            }
          };
        }
      }
    }
  } catch(tScanErr) {
    Logger.log('Training scan fallback error: ' + tScanErr.message);
  }

  if (!getConfigProperty('SPREADSHEET_ID', '')) {
    return { valid: false, message: 'Spreadsheet ID not configured. Please set SPREADSHEET_ID in Apps Script Project Settings.' };
  }

  return { valid: false, message: `Evaluator (${cleanInput}) is not registered in the employee directory.` };
}

// ─── 3. Check Employee Enrollment for Training ──────────────────────────────────
function validateParticipantEnrollment(trainingId, employeeId) {
  const cleanTId   = String(trainingId || '').trim();
  const cleanEmpId = String(employeeId || '').trim();

  // A. Check per-training spreadsheet TrainingParticipants / Participants tab
  const ss = getTrainingDataSpreadsheet(cleanTId);
  const tpSheet = ss ? (ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants') || ss.getSheetByName('ParticipantList')) : null;
  if (tpSheet) {
    const tpRows = sheetToJson(tpSheet);
    const enrolled = tpRows.find(r => 
      isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId) ||
      (r.EmployeeName && String(r.EmployeeName).toLowerCase().includes(cleanEmpId.toLowerCase()))
    );
    if (enrolled) return { valid: true, participant: enrolled };
  }

  // B. Check ParticipantList JSON on training record
  try {
    const tCheck = getValidTraining(cleanTId);
    if (tCheck.valid && tCheck.training) {
      const rawJson = tCheck.training.ParticipantList || tCheck.training.participants;
      if (rawJson) {
        const list = Array.isArray(rawJson) ? rawJson : JSON.parse(rawJson);
        if (Array.isArray(list)) {
          const m = list.find(p => isSameEmployeeId(p.EmployeeID || p.EmployeeNo || p.ID || '', cleanEmpId));
          if (m) return { valid: true, participant: m };
        }
      }
    }
  } catch(jErr) {}

  // C. Check central Participants sheet in Main DB
  try {
    const dbPartSheet = getSheet('Participants') || getSheet('TrainingParticipants');
    if (dbPartSheet) {
      const dbParts = sheetToJson(dbPartSheet);
      const m = dbParts.find(p => {
        const pTid = String(p.TrainingID || p.TrainingCode || p.ID || '').trim().toLowerCase();
        return (pTid === cleanTId.toLowerCase()) && isSameEmployeeId(p.EmployeeID || p.EmployeeNo || p.ID || '', cleanEmpId);
      });
      if (m) return { valid: true, participant: m };
    }
  } catch(dbErr) {}

  return { valid: false, message: `Employee ID (${cleanEmpId}) is not enrolled in this training.` };
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

    // Check if SessionDate + EndTime has passed (only for valid current dates, ignoring 1899 date epoch)
    if (session.SessionDate && qrStatus !== 'Active') {
      try {
        const currentDate = new Date();
        const endTimeStr = session.EndTime || '17:00';
        const sDate = new Date(session.SessionDate);
        if (!isNaN(sDate.getTime()) && sDate.getFullYear() > 2020) {
          const [hours, minutes] = String(endTimeStr).split(':').map(Number);
          sDate.setHours(hours || 17, minutes || 0, 0, 0);

          if (currentDate > sDate) {
            if (sSheet) {
              const row = findRowById(sSheet, cleanSessionId);
              if (row !== -1) sSheet.getRange(row, 9).setValue('Expired');
            }
            return { valid: false, message: `Attendance registration for this session is closed.` };
          }
        }
      } catch (err) {
        Logger.log('Session time check error: ' + err.message);
      }
    }

    const statusLower = qrStatus.toLowerCase();
    if (statusLower === 'inactive' || statusLower === 'deleted') {
      return { valid: false, message: 'This QR attendance session is no longer active. Attendance cannot be recorded.' };
    }
    if (statusLower === 'expired') {
      return { valid: false, message: 'Attendance registration for this session is closed (Expired).' };
    }
    if (statusLower === 'scheduled' || statusLower === 'draft') {
      return { valid: false, message: 'Attendance check-in has not opened yet for this session.' };
    }

    const trainingId = session.TrainingID;

    // B. Validate Training Existence
    const tCheck = getValidTraining(trainingId);
    if (!tCheck.valid) return tCheck;

    // C. Validate Employee Existence (Checking training participant list first)
    const empCheck = getValidEmployee(cleanEmpId, trainingId);
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
  const cleanEmpId = String(employeeId).trim();

  const ss = getTrainingDataSpreadsheet(cleanTId);
  if (!ss) {
    return { eligible: true, message: 'Training data sheet is unavailable, bypassing attendance lock.' };
  }

  const attSheet = ss.getSheetByName('Attendance');
  if (!attSheet) {
    return { eligible: true, message: 'Attendance tab not created yet.' };
  }

  const attRows = sheetToJson(attSheet);
  if (attRows.length === 0) {
    return { eligible: true, message: 'No attendance rows recorded yet.' };
  }

  const userAttRecord = attRows.find(r => {
    const rEmpNo = String(r.EmployeeNo || r.EmployeeID || r.ID || '').trim();
    return isSameEmployeeId(rEmpNo, cleanEmpId);
  });

  if (userAttRecord) {
    const status = String(userAttRecord.Status || 'Present').trim().toLowerCase();
    if (status === 'absent') {
      return {
        eligible: false,
        message: 'Your attendance for this training was marked as Absent.'
      };
    }
    return { eligible: true, attendanceRecord: userAttRecord };
  }

  const enrollCheck = validateParticipantEnrollment(cleanTId, cleanEmpId);
  if (enrollCheck.valid) {
    return { eligible: true, participant: enrollCheck.participant };
  }

  return {
    eligible: false,
    message: 'Our records show that you do not have an attendance record for this training.'
  };
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

function isTrainingFinished(training) {
  if (!training) return false;
  const status = String(training.Status || '').trim().toLowerCase();
  const stage  = String(training.Stage || '').trim().toLowerCase();

  if (status === 'completed' || stage === 'training completed' || stage === 'evaluation completed' || stage === 'waiting for 3-month review' || stage === 'programme closed') {
    return true;
  }

  const dateVal = training.EndDate || training.StartDate;
  if (!dateVal) return true;

  let compDate = null;
  if (dateVal instanceof Date) {
    compDate = isNaN(dateVal.getTime()) ? null : new Date(dateVal);
  } else {
    const str = String(dateVal).trim();
    const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (ymd) {
      compDate = new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]), 23, 59, 59, 999);
    } else {
      const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if (dmy) {
        compDate = new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]), 23, 59, 59, 999);
      } else {
        const d = new Date(str);
        if (!isNaN(d.getTime())) {
          d.setHours(23, 59, 59, 999);
          compDate = d;
        }
      }
    }
  }

  if (!compDate) return true;

  const now = new Date();
  return now.getTime() >= compDate.getTime();
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

    // B. Validate Employee Existence (checking training participant list first)
    const empCheck = getValidEmployee(cleanEmpId, cleanTId);
    if (!empCheck.valid) return err(empCheck.message);

    // C. Validate Employee Enrollment
    const enrollCheck = validateParticipantEnrollment(cleanTId, cleanEmpId);
    if (!enrollCheck.valid) return err(enrollCheck.message);

    // D. Validate Training Completion: Training must be finished before evaluation is allowed
    if (!isTrainingFinished(tCheck.training)) {
      const compDateStr = formatMinimalistDate(tCheck.training.EndDate || tCheck.training.StartDate);
      return ok({
        eligible: false,
        notFinished: true,
        completionDate: compDateStr,
        message: `Training evaluation is not available yet. This evaluation can only be submitted after the training programme has finished (Scheduled completion: ${compDateStr}).`,
        training: tCheck.training,
        employee: empCheck.employee
      });
    }

    // E. Validate Attendance Eligibility (CRITICAL REQUIREMENT: Must have attended at least 1 session)
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
    const evalSheet = ss ? (ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval')) : null;
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
      return err('Supervisor / PIC Employee ID or Email is required.');
    }
    const cleanEvalEmpId = String(evaluatorEmployeeId).trim();
    let trnIdFilter = String(trainingId || '').trim();

    // 1. Verify Evaluator Employee Record from Directory or Per-Training Data
    const empCheck = getValidEmployee(cleanEvalEmpId, trnIdFilter);
    if (!empCheck.valid) {
      return err(`Evaluator (${cleanEvalEmpId}) is not registered in the employee directory.`);
    }

    const evaluator = empCheck.employee;
    const evalEmpId = evaluator.EmployeeID || evaluator.ID || cleanEvalEmpId;
    const evalName  = evaluator.Name || evaluator.EmployeeName || cleanEvalEmpId;
    const evalDept  = evaluator.CostCentre || evaluator.Department || '';
    const evalEmail = evaluator.Email || evaluator.EmailAddress || '';

    // 2. Fetch all trainings from Master Sheet
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return err('Trainings sheet unavailable.');
    const allTrainings = sheetToJson(tSheet);

    const targetTrainingsList = trnIdFilter
      ? allTrainings.filter(r => {
          const id = String(r.ID || '').trim().toLowerCase();
          const code = String(r.Code || '').trim().toLowerCase();
          const tId = String(r.TrainingID || '').trim().toLowerCase();
          return id === trnIdFilter.toLowerCase() || code === trnIdFilter.toLowerCase() || tId === trnIdFilter.toLowerCase();
        })
      : allTrainings;

    const trainingCards = [];
    let overallPendingCount = 0;
    let overallCompletedCount = 0;
    let overallAssignedCount = 0;
    let isSelfParticipantAnywhere = false;
    let anySupervisorConfiguredAnywhere = false;

    targetTrainingsList.forEach(t => {
      const tId = t.ID || t.TrainingID || t.Code;
      if (!tId) return;

      const ss = getTrainingDataSpreadsheet(tId);
      if (!ss) return;

      const tpSheet   = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
      const postSheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');

      const tpRows   = tpSheet ? sheetToJson(tpSheet) : [];
      const postRows = postSheet ? sheetToJson(postSheet) : [];
      const completedEmpIds = postRows.map(r => String(r.EmployeeID || '').trim().toLowerCase());

      const tPending = [];
      const tCompleted = [];

      tpRows.forEach(p => {
        const empId = String(p.EmployeeID || p.ID || '').trim();
        if (!empId) return;

        const pSupId = String(p.SupervisorID || '').trim();
        const pSupEmail = String(p.SupervisorEmail || '').trim().toLowerCase();
        const pSupName = String(p.SupervisorName || '').trim().toLowerCase();

        if (pSupId || pSupEmail || pSupName) {
          anySupervisorConfiguredAnywhere = true;
        }

        // Rule: Block self-evaluation
        if (isSameEmployeeId(evalEmpId, empId) || isSameEmployeeId(cleanEvalEmpId, empId)) {
          isSelfParticipantAnywhere = true;
          return;
        }

        // Strict supervisor assignment match
        const isExplicitSupervisor = (
          (pSupId && (isSameEmployeeId(pSupId, evalEmpId) || isSameEmployeeId(pSupId, cleanEvalEmpId))) ||
          (pSupEmail && evalEmail && pSupEmail === evalEmail.toLowerCase()) ||
          (pSupName && evalName && pSupName === evalName.toLowerCase())
        );

        if (!isExplicitSupervisor) return;

        const pObj = {
          EmployeeID: empId,
          Name: p.EmployeeName || p.Name || empId,
          Department: p.CostCentre || p.Department || '',
          Position: p.Position || p.JobTitle || 'Participant',
          TrainingID: tId,
          TrainingName: t.Name || ''
        };

        if (completedEmpIds.includes(empId.toLowerCase())) {
          tCompleted.push(pObj);
        } else {
          tPending.push(pObj);
        }
      });

      if (tPending.length > 0 || tCompleted.length > 0) {
        const lockInfo = computeTrainingLockInfo(t);
        trainingCards.push({
          ID: t.ID,
          Code: t.Code || t.ID,
          Name: t.Name || '',
          Category: t.Category || '',
          Trainer: t.Trainer || '',
          StartDate: formatMinimalistDate(t.StartDate),
          EndDate: formatMinimalistDate(t.EndDate),
          lockInfo: lockInfo,
          totalParticipants: tPending.length + tCompleted.length,
          pendingCount: tPending.length,
          completedCount: tCompleted.length,
          pendingParticipants: tPending,
          completedParticipants: tCompleted
        });

        overallPendingCount += tPending.length;
        overallCompletedCount += tCompleted.length;
        overallAssignedCount += (tPending.length + tCompleted.length);
      }
    });

    if (trainingCards.length === 0) {
      if (trnIdFilter) {
        return err(`Access Denied: You (${evalName}) are not assigned as the Supervisor / PIC for any participants in this training programme.`);
      }
      if (isSelfParticipantAnywhere && overallAssignedCount === 0) {
        return err(`Access Notice: You (${evalName}) are registered as a participant. Participants cannot evaluate themselves. 3-Month post-evaluations must be conducted by your assigned Supervisor or HOD.`);
      }
      // When accessing general dashboard with no specific training filter, allow showing empty dashboard
    }

    const primaryTrainingCard = trnIdFilter
      ? (trainingCards.find(c => String(c.ID).toLowerCase() === trnIdFilter.toLowerCase() || String(c.Code).toLowerCase() === trnIdFilter.toLowerCase()) || trainingCards[0] || null)
      : (trainingCards[0] || null);

    return ok({
      evaluator: {
        EmployeeID: evalEmpId,
        Name: evalName,
        Department: evalDept || 'Supervisor / Evaluator',
        Email: evalEmail
      },
      trainings: trainingCards,
      overallStats: {
        totalTrainings: trainingCards.length,
        totalPending: overallPendingCount,
        totalCompleted: overallCompletedCount,
        totalAssigned: overallAssignedCount
      },
      // Backward-compatible fields
      training: primaryTrainingCard ? {
        ID: primaryTrainingCard.ID,
        Code: primaryTrainingCard.Code,
        Name: primaryTrainingCard.Name,
        StartDate: primaryTrainingCard.StartDate,
        EndDate: primaryTrainingCard.EndDate
      } : null,
      lockInfo: primaryTrainingCard ? primaryTrainingCard.lockInfo : null,
      pendingParticipants: primaryTrainingCard ? primaryTrainingCard.pendingParticipants : [],
      completedCount: primaryTrainingCard ? primaryTrainingCard.completedCount : 0
    });

  } catch (e) {
    Logger.log('verifyEvaluatorByEmployeeId error: ' + e.message);
    return err('Evaluator verification error: ' + e.message);
  }
}

/**
 * Computes 3-Month Post Evaluation lock info & countdown target date
 */
function computeTrainingLockInfo(training) {
  const startDateStr = training.StartDate || training.EndDate || new Date();
  const completionDateStr = training.EndDate || training.StartDate || new Date();
  const completionDate = new Date(completionDateStr);
  completionDate.setHours(23, 59, 59, 999);

  const unlockTargetDate = new Date(completionDate);
  unlockTargetDate.setMonth(unlockTargetDate.getMonth() + 3);

  const now = new Date();
  const isTrainingCompleted = now.getTime() >= completionDate.getTime();
  const isUnlocked = isTrainingCompleted && (now.getTime() >= unlockTargetDate.getTime());
  const countdownActive = isTrainingCompleted && !isUnlocked;
  const remainingMs = Math.max(0, unlockTargetDate.getTime() - now.getTime());

  let phase = 'UNLOCKED';
  if (!isTrainingCompleted) {
    phase = 'NOT_STARTED';
  } else if (!isUnlocked) {
    phase = 'COUNTDOWN_ACTIVE';
  }

  return {
    phase: phase,
    isTrainingCompleted: isTrainingCompleted,
    isUnlocked: isUnlocked,
    countdownActive: countdownActive,
    startDateFormatted: formatMinimalistDate(startDateStr),
    completionDateFormatted: formatMinimalistDate(completionDateStr),
    unlockTargetIso: unlockTargetDate.toISOString(),
    unlockTargetDateFormatted: formatMinimalistDate(unlockTargetDate),
    remainingMs: remainingMs
  };
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

    // B. Validate Employee Existence (checking training participant list first)
    const empCheck = getValidEmployee(cleanEmpId, cleanTId);
    if (!empCheck.valid) return empCheck;

    // C. Validate Employee Enrollment
    const enrollCheck = validateParticipantEnrollment(cleanTId, cleanEmpId);
    if (!enrollCheck.valid) return enrollCheck;

    // D1. Validate Training Completion: Training must be finished before evaluation is allowed
    if (!isTrainingFinished(tCheck.training)) {
      const compDateStr = formatMinimalistDate(tCheck.training.EndDate || tCheck.training.StartDate);
      return {
        valid: false,
        message: `Submission Rejected: Training evaluation can only be submitted after the training programme has finished (Scheduled completion: ${compDateStr}).`
      };
    }

    // D2. Attendance Verification: Disallow evaluation for participants without attendance
    const attCheck = checkEmployeeAttendanceEligibility(cleanTId, cleanEmpId);
    if (!attCheck.eligible) {
      return {
        valid: false,
        message: attCheck.message
      };
    }

    // E. Prevent Duplicate Evaluation Submission
    const ss = getTrainingDataSpreadsheet(cleanTId);
    const evalSheet = ss ? (ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval')) : null;
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
function validatePublicPostEvaluation(trainingId, employeeId, token, evaluatorId, evaluatorName) {
  try {
    const cleanTId   = String(trainingId || '').trim();
    const cleanEmpId = String(employeeId || '').trim();
    const cleanToken = String(token || '').trim();
    const cleanEvalId = String(evaluatorId || '').trim();

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

    // RULE 1: Self-evaluation is STRICTLY PROHIBITED
    if (cleanEvalId && isSameEmployeeId(cleanEvalId, effectiveEmpId)) {
      return {
        valid: false,
        message: 'Submission Rejected: Participants are NOT permitted to fill 3-Month post evaluations on themselves. This evaluation must be conducted by an assigned supervisor or person in charge.'
      };
    }

    // A. Validate Training Existence
    const tCheck = getValidTraining(effectiveTId);
    if (!tCheck.valid) return tCheck;

    // B. Validate Employee Existence (checking training participant list first)
    const empCheck = getValidEmployee(effectiveEmpId, effectiveTId);
    if (!empCheck.valid) return empCheck;

    // C. Validate Employee Enrollment
    const enrollCheck = validateParticipantEnrollment(effectiveTId, effectiveEmpId);
    if (!enrollCheck.valid) return enrollCheck;

    // D. Validate Evaluator Authorization against Assigned Supervisor
    const ss = getTrainingDataSpreadsheet(effectiveTId);
    if (ss) {
      const tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
      if (tpSheet) {
        const tpRows = sheetToJson(tpSheet);
        const p = tpRows.find(r => isSameEmployeeId(r.EmployeeID || r.ID || '', effectiveEmpId));
        if (p) {
          const pSupId = String(p.SupervisorID || '').trim();
          const pSupEmail = String(p.SupervisorEmail || '').trim().toLowerCase();
          const pSupName = String(p.SupervisorName || '').trim().toLowerCase();

          if (!pSupId && !pSupEmail && !pSupName) {
            return {
              valid: false,
              message: 'Submission Rejected: No supervisor or person in charge (PIC) has been assigned by the Admin yet for this participant.'
            };
          }

          const evalCheck = cleanEvalId ? getValidEmployee(cleanEvalId, effectiveTId) : { valid: false };
          const evalObj = evalCheck.valid ? evalCheck.employee : {};
          const evalEmail = String(evalObj.Email || '').trim().toLowerCase();
          const evalName = String(evalObj.Name || evaluatorName || '').trim().toLowerCase();

          const isAssigned = (
            (pSupId && isSameEmployeeId(pSupId, cleanEvalId)) ||
            (pSupEmail && evalEmail && pSupEmail === evalEmail) ||
            (pSupName && evalName && pSupName === evalName)
          );

          if (!isAssigned) {
            return {
              valid: false,
              message: `Submission Rejected: You are not assigned as the supervisor for this participant (${p.EmployeeName || effectiveEmpId}).`
            };
          }
        }
      }
    }

    // E. 3-Month Lock Enforcement (Must unlock AFTER 3 months of training completion)
    const endDateStr = tCheck.training.EndDate || tCheck.training.StartDate || new Date();
    const completionDate = new Date(endDateStr);
    completionDate.setHours(23, 59, 59, 999);
    const now = new Date();

    if (now.getTime() < completionDate.getTime()) {
      return {
        valid: false,
        message: `Submission Rejected: Training has not been completed yet (Scheduled completion: ${formatMinimalistDate(endDateStr)}). 3-Month countdown begins upon course completion.`
      };
    }

    const unlockTargetDate = new Date(completionDate);
    unlockTargetDate.setMonth(unlockTargetDate.getMonth() + 3);
    if (now.getTime() < unlockTargetDate.getTime()) {
      return {
        valid: false,
        message: `Submission Rejected: 3-Month Post-Training Evaluation is locked until ${formatMinimalistDate(unlockTargetDate)} (evaluation unlocks 3 months after course completion on ${formatMinimalistDate(endDateStr)}).`
      };
    }

    // F. Prevent Duplicate Post-Evaluation Submission
    const postSheet = ss ? (ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval')) : null;
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
