/**
 * ValidationService.gs — Server-Side Security & Data Validation Engine for Public Participant Portal
 * 
 * Enforces strict validation rules on the server for all public submissions.
 * Rejects invalid requests before writing to the database.
 */

// ─── 1. Check Training Existence ────────────────────────────────────────────────
function getValidTraining(trainingId) {
  if (!trainingId || String(trainingId).trim() === '') {
    return { valid: true, training: { ID: '', Name: 'Training Programme', TrainingTitle: 'Training Programme' } };
  }
  const cleanId = String(trainingId).trim();
  const tSheet = getSheet(SHEET_NAMES.trainings);
  if (!tSheet) {
    return { valid: true, training: { ID: cleanId, Name: 'Training Programme', TrainingTitle: 'Training Programme' } };
  }

  const rows = sheetToJson(tSheet);
  const training = rows.find(r => {
    const id = String(r.ID || r['Training ID'] || r.TrainingID || '').trim();
    const code = String(r.Code || r['Training Code'] || r.TrainingCode || '').trim();
    const tId = String(r.TrainingID || '').trim();
    return isSameTrainingId(id, cleanId) || isSameTrainingId(code, cleanId) || isSameTrainingId(tId, cleanId);
  });

  if (!training) {
    return { valid: true, training: { ID: cleanId, Name: 'Training Programme', TrainingTitle: 'Training Programme' } };
  }
  return { valid: true, training: training };
}

// ─── 2. Participant Validation (Attendance & Level 1 Evaluation) ─────────────
/**
 * Validate and look up a PARTICIPANT for training attendance and Level 1 evaluation.
 * Strictly uses participant, attendance, and employee directory records.
 * NEVER checks supervisor assignment columns or evaluator/HOD routing tables.
 * 
 * @param {string} employeeId - Employee ID or Badge Number
 * @param {string} [trainingId] - Training ID (optional, to verify enrollment)
 * @param {Spreadsheet} [optionalSpreadsheet] - Pre-opened spreadsheet object (e.g. from session lookup)
 * @returns {Object} { valid: boolean, enrolled: boolean, employee?: Object, message?: string }
 */
function getValidParticipant(employeeId, trainingId, optionalSpreadsheet) {
  if (!employeeId || String(employeeId).trim() === '') {
    return { valid: false, message: 'Employee ID is required.' };
  }
  const cleanInput = String(employeeId).trim();
  const cleanInputLower = cleanInput.toLowerCase();
  const cleanTId   = trainingId ? String(trainingId).trim() : '';

  // Universal participant matching predicate across diverse sheet schemas
  const matchesParticipant = (r) => {
    if (!r) return false;
    const possibleIds = [
      r.EmployeeID, r['Employee ID'], r.Employee_ID,
      r.EmployeeNo, r['Employee No'], r['Employee Number'],
      r.EmpID, r['Emp ID'], r.EmpNo, r['Emp No'],
      r.StaffID, r['Staff ID'], r.StaffNo, r['Staff No'],
      r.BadgeNo, r['Badge No'], r.BadgeNumber,
      r.No, r['No.'],
      r.ID
    ];

    for (let i = 0; i < possibleIds.length; i++) {
      const pid = possibleIds[i];
      if (pid !== undefined && pid !== null && String(pid).trim() !== '' && isSameEmployeeId(pid, cleanInput)) {
        return true;
      }
    }

    const email = r.Email || r.EmailAddress || r['Email Address'] || r['Company Email'] || r['Work Email'];
    if (email && String(email).trim().toLowerCase() === cleanInputLower) return true;

    const name = r.EmployeeName || r['Employee Name'] || r.Name || r['Staff Name'] || r.FullName || r['Full Name'];
    if (name && String(name).trim().toLowerCase() === cleanInputLower) return true;

    return false;
  };

  // Universal participant details normalizer (Standard schema: ID, Name, Department, Position, Email)
  const extractParticipantDetails = (r) => {
    const empId = r.EmployeeID || r['Employee ID'] || r.Employee_ID ||
                  r.EmployeeNo || r['Employee No'] || r.EmpID || r['Emp ID'] ||
                  r.StaffID || r['Staff ID'] || r.BadgeNo || r.ID || cleanInput;

    const empName = r.EmployeeName || r['Employee Name'] || r.Name || r['Staff Name'] ||
                    r.FullName || r['Full Name'] || cleanInput;

    const empDept = r.Department || r.CostCentre || r['Cost Centre'] || r.Dept || r.Section || '';
    const empPos  = r.Position || r.JobTitle || r.PositionTitle || r['Position Title'] || r['Job Title'] || r.Designation || 'Participant';
    const empEmail = r.Email || r.EmailAddress || r['Email Address'] || r['Company Email'] || '';

    return {
      ID: String(empId).trim(),
      Name: String(empName).trim(),
      Department: String(empDept).trim(),
      Position: String(empPos).trim(),
      Email: String(empEmail).trim()
    };
  };

  // 1. Primary Lookup: Per-Training Spreadsheet(s)
  const candidateSheets = [];
  if (optionalSpreadsheet) candidateSheets.push(optionalSpreadsheet);
  if (cleanTId) {
    try {
      const ss = getTrainingDataSpreadsheet(cleanTId);
      if (ss && !candidateSheets.includes(ss)) candidateSheets.push(ss);
    } catch (e) {
      Logger.log('Per-training spreadsheet lookup error: ' + e.message);
    }
  }

  for (const ss of candidateSheets) {
    if (!ss) continue;
    try {
      // 1a. Check Participants Tab (ONLY participant fields, ignoring supervisor columns)
      const tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants') || ss.getSheetByName('ParticipantList');
      if (tpSheet && tpSheet.getLastRow() > 1) {
        const tpRows = sheetToJson(tpSheet);
        const tpEmp = tpRows.find(matchesParticipant);
        if (tpEmp) {
          return {
            valid: true,
            enrolled: true,
            employee: extractParticipantDetails(tpEmp)
          };
        }
      }

      // 1b. Check Attendance Tab in this Training
      const attSheet = ss.getSheetByName('Attendance');
      if (attSheet && attSheet.getLastRow() > 1) {
        const attRows = sheetToJson(attSheet);
        const attEmp = attRows.find(matchesParticipant);
        if (attEmp) {
          return {
            valid: true,
            enrolled: true,
            employee: extractParticipantDetails(attEmp)
          };
        }
      }
    } catch(ssErr) {
      Logger.log('Error reading candidate spreadsheet: ' + ssErr.message);
    }
  }

  // 1c. Check ParticipantList JSON on training record
  if (cleanTId) {
    try {
      const tCheck = getValidTraining(cleanTId);
      if (tCheck.valid && tCheck.training) {
        const rawJson = tCheck.training.ParticipantList || tCheck.training.participants || tCheck.training.Participants;
        if (rawJson) {
          const list = Array.isArray(rawJson) ? rawJson : JSON.parse(rawJson);
          if (Array.isArray(list)) {
            const m = list.find(matchesParticipant);
            if (m) {
              return {
                valid: true,
                enrolled: true,
                employee: extractParticipantDetails(m)
              };
            }
          }
        }
      }
    } catch(jErr) {}
  }

  // 2. Check Central Database sheet 'Participants' / 'TrainingParticipants'
  try {
    const mainSs = getSpreadsheet();
    if (mainSs) {
      const dbPartSheet = mainSs.getSheetByName('Participants') || mainSs.getSheetByName('TrainingParticipants');
      if (dbPartSheet && dbPartSheet.getLastRow() > 1) {
        const dbParts = sheetToJson(dbPartSheet);
        const m = dbParts.find(p => {
          const pTid = String(p.TrainingID || p.TrainingCode || p.ID || '').trim().toLowerCase();
          if (cleanTId && pTid && pTid !== cleanTId.toLowerCase()) return false;
          return matchesParticipant(p);
        });
        if (m) {
          return {
            valid: true,
            enrolled: true,
            employee: extractParticipantDetails(m)
          };
        }
      }
    }
  } catch(dbErr) {}

  // 3. Employee Master Directory Lookup (To verify if employee is in the company roster)
  const masterSheetTabs = [
    SHEET_NAMES.employees,
    'For IT',
    'FOR IT',
    'Employees',
    'EMPLOYEES',
    'Staff',
    'Staff List',
    'Employee',
    'Master Employees'
  ];

  let foundInDirectory = null;
  const directorySpreadsheets = [];
  try {
    const empSs = getEmployeeSpreadsheet();
    if (empSs) directorySpreadsheets.push(empSs);
  } catch(e) {}
  try {
    const mainSs = getSpreadsheet();
    if (mainSs && !directorySpreadsheets.includes(mainSs)) directorySpreadsheets.push(mainSs);
  } catch(e) {}

  for (const ssObj of directorySpreadsheets) {
    if (!ssObj) continue;
    for (const tabName of masterSheetTabs) {
      if (!tabName) continue;
      try {
        const empSheet = ssObj.getSheetByName(tabName);
        if (empSheet && empSheet.getLastRow() > 1) {
          const rows = sheetToJson(empSheet);
          const emp = rows.find(matchesParticipant);
          if (emp) {
            foundInDirectory = extractParticipantDetails(emp);
            break;
          }
        }
      } catch(sheetErr) {}
    }
    if (foundInDirectory) break;
  }

  if (foundInDirectory) {
    if (cleanTId) {
      let hasExplicitParticipants = false;
      for (const ss of candidateSheets) {
        if (!ss) continue;
        const tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants') || ss.getSheetByName('ParticipantList');
        if (tpSheet && tpSheet.getLastRow() > 1) {
          hasExplicitParticipants = true;
          break;
        }
      }

      if (hasExplicitParticipants) {
        return {
          valid: false,
          enrolled: false,
          employee: foundInDirectory,
          message: `Employee ID (${cleanInput}) - ${foundInDirectory.Name} is not enrolled in this training programme.`
        };
      }
    }

    return {
      valid: true,
      enrolled: true,
      employee: foundInDirectory
    };
  }

  // 4. Safety Fallback across all per-training spreadsheets if employee was enrolled in any training
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet && tSheet.getLastRow() > 1) {
      const allTrainings = sheetToJson(tSheet);
      for (const t of allTrainings) {
        const ss = getTrainingDataSpreadsheet(t);
        if (!ss) continue;
        const pSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
        if (pSheet && pSheet.getLastRow() > 1) {
          const pRows = sheetToJson(pSheet);
          const matched = pRows.find(matchesParticipant);
          if (matched) {
            const empDetails = extractParticipantDetails(matched);
            if (cleanTId) {
              const currentTCode = String(t.ID || t.Code || '').trim().toLowerCase();
              if (currentTCode === cleanTId.toLowerCase()) {
                return {
                  valid: true,
                  enrolled: true,
                  employee: empDetails
                };
              }
            } else {
              return {
                valid: true,
                enrolled: true,
                employee: empDetails
              };
            }
          }
        }
      }
    }
  } catch(fbErr) {}

  return {
    valid: false,
    enrolled: false,
    message: `Employee ID (${cleanInput}) is not registered in the employee directory.`
  };
}

// ─── 3. Check Employee Enrollment for Training ──────────────────────────────────
function validateParticipantEnrollment(trainingId, employeeId) {
  const cleanTId   = String(trainingId || '').trim();
  const cleanEmpId = String(employeeId || '').trim();

  const partCheck = getValidParticipant(cleanEmpId, cleanTId);
  if (partCheck.valid && partCheck.enrolled) {
    return { valid: true, participant: partCheck.employee };
  }

  return {
    valid: false,
    message: partCheck.message || `Employee ID (${cleanEmpId}) is not enrolled in this training programme.`
  };
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
    const qrStatus = String(session.QRStatus || 'Active').trim();
    const statusLower = qrStatus.toLowerCase();

    if (statusLower === 'deactivate' || statusLower === 'deactivated' || statusLower === 'inactive' || statusLower === 'deleted' || statusLower === 'expired') {
      return { valid: false, message: 'This QR attendance session is deactivated. Attendance cannot be recorded.' };
    }
    if (statusLower === 'scheduled' || statusLower === 'draft') {
      return { valid: false, message: 'Attendance check-in has not opened yet for this session.' };
    }

    const trainingId = session.TrainingID;

    // B. Validate Training Existence
    const tCheck = getValidTraining(trainingId);
    if (!tCheck.valid) return tCheck;

    // C. Validate Participant Existence & Enrollment (Using Attendance/Participant DB only)
    const partCheck = getValidParticipant(cleanEmpId, trainingId, found.spreadsheet);
    if (!partCheck.valid || !partCheck.enrolled) {
      return {
        valid: false,
        message: partCheck.message || `Employee ID (${cleanEmpId}) is not enrolled in this training programme.`
      };
    }

    // D. Prevent Duplicate Attendance Submission (Using Attendance Tab in Training Data)
    const ss = found.spreadsheet;
    const attSheet = ss ? ss.getSheetByName('Attendance') : null;
    if (attSheet) {
      const attRows = sheetToJson(attSheet);
      const duplicate = attRows.find(r => {
        const rSessionId = String(r.SessionID || '').trim();
        const rEmpNo     = String(r.EmployeeNo || r.EmployeeID || r.ID || '').trim();
        return rSessionId === cleanSessionId && isSameEmployeeId(rEmpNo, cleanEmpId);
      });

      if (duplicate) {
        const isAbsent = String(duplicate.Status || '').trim().toLowerCase() === 'absent';
        if (!isAbsent) {
          return {
            valid: false,
            duplicate: true,
            scanTime: duplicate.ScanTime || duplicate.EditedAt || 'Earlier',
            message: `Attendance has already been recorded for Employee ID (${cleanEmpId}) for session ${cleanSessionId}.`,
            employee: partCheck.employee,
            session: session
          };
        }
      }
    }

    return {
      valid: true,
      session: session,
      training: tCheck.training || found.training,
      employee: partCheck.employee,
      spreadsheet: found.spreadsheet
    };
  } catch (e) {
    Logger.log('validatePublicAttendance error: ' + e.message);
    return { valid: false, message: 'Attendance validation failure: ' + e.message };
  }
}

// ─── 4.1 Attendance Eligibility Check for Level 1 Evaluation ────────────────────
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
    return { eligible: true, message: 'Training data sheet is unavailable.' };
  }

  const attSheet = ss.getSheetByName('Attendance');
  if (!attSheet) {
    return { eligible: true, message: 'Attendance tab not created yet.' };
  }

  const attRows = sheetToJson(attSheet);
  if (attRows.length === 0) {
    // If no attendance records exist yet for this entire training, verify enrollment
    const partCheck = getValidParticipant(cleanEmpId, cleanTId);
    if (partCheck.valid && partCheck.enrolled) {
      return { eligible: true, participant: partCheck.employee };
    }
    return { eligible: false, message: `Employee ID (${cleanEmpId}) is not enrolled in this training.` };
  }

  // Attendance records exist: Participant MUST have an attendance log
  const userAttRecord = attRows.find(r => {
    const rEmpNo = String(r.EmployeeNo || r.EmployeeID || r.ID || '').trim();
    return isSameEmployeeId(rEmpNo, cleanEmpId);
  });

  if (userAttRecord) {
    const status = String(userAttRecord.Status || 'Present').trim().toLowerCase();
    if (status === 'absent' || status === 'did not attend' || status === 'not attend') {
      return {
        eligible: false,
        message: 'Your attendance for this training was marked as Absent. Evaluation is not permitted.'
      };
    }
    return { eligible: true, attendanceRecord: userAttRecord };
  }

  // Attendance taken but this employee has no record
  return {
    eligible: false,
    noAttendance: true,
    message: 'Our records show that you do not have an attendance record for this training. Please attend the training session first.'
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

function isTrainingFinished(training, employeeId) {
  if (!training) return false;
  const status = String(training.Status || '').trim().toLowerCase();
  const stage  = String(training.Stage || '').trim().toLowerCase();

  // 1. Status or stage indicates active or completed
  if (status === 'completed' || status === 'in progress' || status === 'ongoing' ||
      stage === 'training completed' || stage === 'evaluation completed' || 
      stage === 'attendance in progress' || stage === 'waiting for 3-month review' || stage === 'programme closed') {
    return true;
  }

  // 2. If participant already attended, training is definitely active/evaluable
  if (employeeId && (training.ID || training.TrainingID || training.Code)) {
    try {
      const ss = getTrainingDataSpreadsheet(training.ID || training.TrainingID || training.Code);
      if (ss) {
        const attSheet = ss.getSheetByName('Attendance');
        if (attSheet) {
          const attRows = sheetToJson(attSheet);
          const hasAttended = attRows.some(r => isSameEmployeeId(r.EmployeeNo || r.EmployeeID || r.ID || '', employeeId) && String(r.Status).toLowerCase() !== 'absent');
          if (hasAttended) return true;
        }
      }
    } catch(e) {}
  }

  // 3. Compare start date by calendar day (00:00:00) so same-day evaluations are unlocked
  const dateVal = training.StartDate || training.EndDate;
  if (!dateVal) return true;

  let startDateObj = null;
  if (dateVal instanceof Date) {
    startDateObj = isNaN(dateVal.getTime()) ? null : new Date(dateVal);
  } else {
    const str = String(dateVal).trim();
    const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (ymd) {
      startDateObj = new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]), 0, 0, 0, 0);
    } else {
      const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if (dmy) {
        startDateObj = new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]), 0, 0, 0, 0);
      } else {
        const d = new Date(str);
        if (!isNaN(d.getTime())) {
          d.setHours(0, 0, 0, 0);
          startDateObj = d;
        }
      }
    }
  }

  if (!startDateObj) return true;

  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return now.getTime() >= startDateObj.getTime();
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

    // B. Validate Participant Existence & Enrollment (Using Participant/Attendance DB only)
    const partCheck = getValidParticipant(cleanEmpId, cleanTId);
    if (!partCheck.valid || !partCheck.enrolled) {
      return err(partCheck.message || `Employee ID (${cleanEmpId}) is not enrolled in this training programme.`);
    }

    // C. Validate Training Timeline: Unlocks on training date or once attended
    if (!isTrainingFinished(tCheck.training, cleanEmpId)) {
      const compDateStr = formatMinimalistDate(tCheck.training.StartDate || tCheck.training.EndDate);
      return ok({
        eligible: false,
        notFinished: true,
        completionDate: compDateStr,
        message: `Training evaluation is not available yet. This evaluation opens on the training date (Scheduled: ${compDateStr}).`,
        training: tCheck.training,
        employee: partCheck.employee
      });
    }

    // D. Validate Attendance Eligibility (Must have attended at least 1 session)
    const attCheck = checkEmployeeAttendanceEligibility(cleanTId, cleanEmpId);
    if (!attCheck.eligible) {
      return ok({
        eligible: false,
        noAttendance: true,
        message: attCheck.message,
        training: tCheck.training,
        employee: partCheck.employee
      });
    }

    // E. Check if Evaluation Already Submitted (Using isSameEmployeeId)
    const ss = getTrainingDataSpreadsheet(cleanTId);
    const evalSheet = ss ? (ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval')) : null;
    if (evalSheet) {
      const evalRows = sheetToJson(evalSheet);
      const duplicate = evalRows.find(r => 
        isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId)
      );

      if (duplicate) {
        return ok({
          eligible: false,
          alreadySubmitted: true,
          submittedAt: duplicate.SubmittedAt || 'Earlier',
          message: 'You have already completed the evaluation for this training.',
          training: tCheck.training,
          employee: partCheck.employee
        });
      }
    }

    return ok({
      eligible: true,
      valid: true,
      training: tCheck.training,
      employee: partCheck.employee
    });
  } catch (e) {
    return err('Evaluation verification failed: ' + e.message);
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

    // B. Validate Participant Existence & Enrollment (Using Participant/Attendance DB only)
    const partCheck = getValidParticipant(cleanEmpId, cleanTId);
    if (!partCheck.valid || !partCheck.enrolled) {
      return {
        valid: false,
        message: partCheck.message || `Employee ID (${cleanEmpId}) is not enrolled in this training programme.`
      };
    }

    // C. Validate Training Timeline: Unlocks on training date or once attended
    if (!isTrainingFinished(tCheck.training, cleanEmpId)) {
      const compDateStr = formatMinimalistDate(tCheck.training.StartDate || tCheck.training.EndDate);
      return {
        valid: false,
        message: `Submission Rejected: Training evaluation can only be submitted on or after the training date (Scheduled: ${compDateStr}).`
      };
    }

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
    const evalSheet = ss ? (ss.getSheetByName('Evaluation') || ss.getSheetByName('TrainingEval')) : null;
    if (evalSheet) {
      const evalRows = sheetToJson(evalSheet);
      const duplicate = evalRows.find(r => 
        isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', cleanEmpId)
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
      employee: partCheck.employee
    };
  } catch (e) {
    Logger.log('validatePublicEvaluation error: ' + e.message);
    return { valid: false, message: 'Evaluation validation failure: ' + e.message };
  }
}

// ─── 6. Evaluator Authentication & Post-Evaluation (Level 3 - Supervisor Review) ──
/**
 * Dedicated Evaluator / Supervisor Lookup (Used ONLY by 3-Month Post Evaluation).
 */
function getValidEvaluator(evaluatorInput) {
  if (!evaluatorInput || String(evaluatorInput).trim() === '') {
    return { valid: false, message: 'Supervisor / Evaluator Employee ID or Email is required.' };
  }
  const cleanInput = String(evaluatorInput).trim();
  const cleanInputLower = cleanInput.toLowerCase();

  const getRowEmpId = (r) => r.EmployeeID || r.EmployeeNo || r['Employee ID'] || r['Employee No'] || r.EmpID || r.StaffID || r.ID || '';
  const getRowEmail = (r) => r.Email || r.EmailAddress || r['Email Address'] || r['HOD Email'] || r['HR Email'] || '';
  const getRowName = (r) => r.Name || r.EmployeeName || r['HOD Name'] || r['HR Name'] || r.FullName || '';
  const getRowDept = (r) => r.Department || r.CostCentre || r['Cost Centre'] || '';

  // 1. Check Master Employee & HOD Tables
  const evaluatorTabs = ['Employees', 'For IT', 'HOD email', 'HR email', 'Csuite email'];
  for (const tabName of evaluatorTabs) {
    try {
      const sheet = getSheet(tabName);
      if (sheet) {
        const rows = sheetToJson(sheet);
        const emp = rows.find(r => 
          isSameEmployeeId(getRowEmpId(r), cleanInput) ||
          (getRowEmail(r) && String(getRowEmail(r)).trim().toLowerCase() === cleanInputLower) ||
          (getRowName(r) && String(getRowName(r)).trim().toLowerCase() === cleanInputLower) ||
          (r.HODEmail && String(r.HODEmail).trim().toLowerCase() === cleanInputLower) ||
          (r.HODName && String(r.HODName).trim().toLowerCase() === cleanInputLower)
        );
        if (emp) {
          return {
            valid: true,
            evaluator: {
              ID: getRowEmpId(emp) || cleanInput,
              Name: getRowName(emp) || emp.HODName || cleanInput,
              Department: getRowDept(emp) || 'Supervisor / Evaluator',
              Email: getRowEmail(emp) || emp.HODEmail || (cleanInput.includes('@') ? cleanInput : '')
            }
          };
        }
      }
    } catch(e) {}
  }

  // 2. Check Assigned Supervisor columns across Training Data spreadsheets
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
        const found = tpRows.find(r => 
          isSameEmployeeId(r.SupervisorID || '', cleanInput) ||
          (r.SupervisorEmail && String(r.SupervisorEmail).trim().toLowerCase() === cleanInputLower) ||
          (r.SupervisorName && String(r.SupervisorName).trim().toLowerCase() === cleanInputLower)
        );
        if (found) {
          return {
            valid: true,
            evaluator: {
              ID: found.SupervisorID || cleanInput,
              EmployeeID: found.SupervisorID || cleanInput,
              Name: found.SupervisorName || cleanInput,
              Department: getRowDept(found) || 'Supervisor / Evaluator',
              Email: found.SupervisorEmail || (cleanInput.includes('@') ? cleanInput : '')
            }
          };
        }
      }
    }
  } catch(e) {}

  return { valid: false, message: `Evaluator (${cleanInput}) is not registered in the employee directory.` };
}

function verifyEvaluatorByEmployeeId(evaluatorEmployeeId, trainingId) {
  try {
    if (!evaluatorEmployeeId || String(evaluatorEmployeeId).trim() === '') {
      return err('Supervisor / PIC Employee ID or Email is required.');
    }
    const cleanEvalEmpId = String(evaluatorEmployeeId).trim();
    let trnIdFilter = String(trainingId || '').trim();

    // 1. Verify Evaluator Employee Record
    const evalCheck = getValidEvaluator(cleanEvalEmpId);
    if (!evalCheck.valid) {
      return err(evalCheck.message);
    }

    const evaluator = evalCheck.evaluator;
    const evalEmpId = evaluator.EmployeeID || evaluator.ID || cleanEvalEmpId;
    const evalName  = evaluator.Name || cleanEvalEmpId;
    const evalDept  = evaluator.Department || '';
    const evalEmail = evaluator.Email || '';

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
        const empId = String(p.EmployeeID || p.EmployeeNo || p.ID || '').trim();
        if (!empId) return;

        const pSupId = String(p.SupervisorID || '').trim();
        const pSupEmail = String(p.SupervisorEmail || '').trim().toLowerCase();
        const pSupName = String(p.SupervisorName || '').trim().toLowerCase();

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

        if (completedEmpIds.some(cId => isSameEmployeeId(cId, empId))) {
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

// ─── 7. Public Post-Evaluation Validation (3-Month Supervisor Review) ────────────
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

    // B. Validate Participant Existence & Enrollment
    const partCheck = getValidParticipant(effectiveEmpId, effectiveTId);
    if (!partCheck.valid || !partCheck.enrolled) {
      return { valid: false, message: `Subordinate (${effectiveEmpId}) is not enrolled in this training programme.` };
    }

    // C. Validate Evaluator Authorization against Assigned Supervisor
    const ss = getTrainingDataSpreadsheet(effectiveTId);
    if (ss) {
      const tpSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants');
      if (tpSheet) {
        const tpRows = sheetToJson(tpSheet);
        const p = tpRows.find(r => isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', effectiveEmpId));
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

          const evalCheck = cleanEvalId ? getValidEvaluator(cleanEvalId) : { valid: false };
          const evalObj = evalCheck.valid ? evalCheck.evaluator : {};
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

    // D. 3-Month Lock Enforcement (Must unlock AFTER 3 months of training completion)
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

    // E. Prevent Duplicate Post-Evaluation Submission
    const postSheet = ss ? (ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval')) : null;
    if (postSheet) {
      const postRows = sheetToJson(postSheet);
      const duplicate = postRows.find(r =>
        isSameEmployeeId(r.EmployeeID || r.EmployeeNo || r.ID || '', effectiveEmpId)
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
      employee: partCheck.employee
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
    const res = getValidParticipant(employeeId);
    if (!res.valid) return err(res.message);
    return ok(res.employee);
  } catch (e) {
    return err(e.message);
  }
}
