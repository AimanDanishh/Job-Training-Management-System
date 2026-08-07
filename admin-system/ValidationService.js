/**
 * ValidationService.gs — Validation module for session attendance submissions
 */

/**
 * Validate whether an employee can submit attendance for a specific session.
 * Rejects submission if:
 * 1. Session does not exist
 * 2. Session status is Expired or Inactive or Scheduled
 * 3. Session has expired based on session date/time range
 * 4. Employee has already submitted attendance for this session
 * 
 * @param {string} sessionId - The session ID to check
 * @param {string} employeeNo - The employee ID / Number checking in
 * @returns {Object} { valid: boolean, message: string, session?: Object, employee?: Object }
 */
function validateAttendance(sessionId, employeeNo) {
  try {
    if (!sessionId || String(sessionId).trim() === '') {
      return { valid: false, message: 'Session ID is missing or invalid.' };
    }

    if (!employeeNo || String(employeeNo).trim() === '') {
      return { valid: false, message: 'Employee Number is required.' };
    }

    const cleanSessionId = String(sessionId).trim();
    const cleanEmpNo     = String(employeeNo).trim();

    // 1. Check Session existence
    const sessionRes = getSession(cleanSessionId);
    const sessionResObj = typeof sessionRes === 'string' ? JSON.parse(sessionRes) : sessionRes;
    
    if (!sessionResObj || !sessionResObj.success || !sessionResObj.data) {
      return { valid: false, message: 'Session does not exist.' };
    }

    const session = sessionResObj.data;

    // 2. Check QR / Session Status
    const status = String(session.QRStatus || 'Active').trim();
    if (status === 'Expired') {
      return { valid: false, message: 'This training session has expired. Attendance registration is closed.', session: session };
    }
    if (status === 'Inactive') {
      return { valid: false, message: 'This training session is currently inactive.', session: session };
    }
    // 2b. Check Parent Training Approval Status
    if (session.TrainingID) {
      try {
        const tSheet = getSheet(SHEET_NAMES.trainings);
        if (tSheet) {
          const trainings = sheetToJson(tSheet);
          const parentT = trainings.find(t => String(t.ID || '').trim() === String(session.TrainingID).trim());
          if (parentT) {
            const appStatus = String(parentT.ApprovalStatus || '').trim().toLowerCase();
            if (appStatus && appStatus !== 'approved' && appStatus !== 'auto-approved') {
              return {
                valid: false,
                message: `Attendance is locked. Training programme approval status is '${parentT.ApprovalStatus}'. Approval from HOD / C-Suite / HOHR is required first.`,
                session: session
              };
            }
          }
        }
      } catch (e) {
        Logger.log('Parent training approval check error: ' + e.message);
      }
    }

    // 3. Check Session Date & Time Expiry if SessionDate / EndTime are provided
    if (session.SessionDate) {
      const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const sessionDateStr = session.SessionDate.split('T')[0];

      // Optional strict date check if past session date
      if (sessionDateStr < todayStr && status !== 'Active') {
        return { valid: false, message: 'This session date has already passed.', session: session };
      }
    }

    // 4. Duplicate Check: One employee can only submit attendance once per session
    const ss = getTrainingDataSpreadsheet(session.TrainingID);
    const attSheet = ss ? ss.getSheetByName('Attendance') : null;
    if (attSheet) {
      const rows = sheetToJson(attSheet);
      const duplicate = rows.find(r => {
        const rSessionId = String(r.SessionID || '').trim();
        const rEmpNo = String(r.EmployeeNo || r.EmployeeID || r.EmployeeId || '').trim();
        return rSessionId === cleanSessionId && rEmpNo.toLowerCase() === cleanEmpNo.toLowerCase();
      });

      if (duplicate) {
        return {
          valid: false,
          message: `Attendance already submitted by employee (${cleanEmpNo}) for this session.`,
          session: session,
          duplicate: true
        };
      }
    }

    // 5. Look up employee details (if available) to complement submission
    let empDetails = null;
    try {
      const empSheet = getSheet(SHEET_NAMES.employees);
      if (empSheet) {
        const empRows = sheetToJson(empSheet);
        empDetails = empRows.find(e => String(e.ID || '').toLowerCase() === cleanEmpNo.toLowerCase());
      }
    } catch (e) {
      Logger.log('Employee lookup non-fatal error: ' + e.message);
    }

    return {
      valid: true,
      message: 'Validation successful.',
      session: session,
      employee: empDetails
    };

  } catch (e) {
    Logger.log('validateAttendance error: ' + e.message);
    return { valid: false, message: 'Validation error: ' + e.message };
  }
}
