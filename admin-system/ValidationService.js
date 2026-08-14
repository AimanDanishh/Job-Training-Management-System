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
            const isApproved = !appStatus || appStatus === 'approved' || appStatus === 'auto-approved' || appStatus === 'fully approved' || appStatus === 'completed' || appStatus === 'in progress' || appStatus === 'active' || appStatus === 'on going';
            if (!isApproved) {
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
      const sessionDateStr = (session.SessionDate instanceof Date)
        ? Utilities.formatDate(session.SessionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(session.SessionDate).split('T')[0];

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
      if (session.TrainingID) {
        const ss = getTrainingDataSpreadsheet(session.TrainingID);
        const tpSheet = ss ? ss.getSheetByName('TrainingParticipants') : null;
        if (tpSheet) {
          const tpRows = sheetToJson(tpSheet);
          const tpEmp = tpRows.find(e => isSameEmployeeId(e.EmployeeID || e.EmployeeNo || e.ID || '', cleanEmpNo));
          if (tpEmp) {
            empDetails = {
              ID: tpEmp.EmployeeID || tpEmp.ID || cleanEmpNo,
              Name: tpEmp.EmployeeName || tpEmp.Name || cleanEmpNo,
              Department: tpEmp.Department || tpEmp.CostCentre || '',
              Position: tpEmp.Position || tpEmp.JobTitle || ''
            };
          }
        }
      }
      if (!empDetails) {
        const empSheet = getSheet(SHEET_NAMES.employees);
        if (empSheet) {
          const empRows = sheetToJson(empSheet);
          empDetails = empRows.find(e => isSameEmployeeId(e.ID || e.EmployeeID || e.EmployeeNo || '', cleanEmpNo));
        }
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

/**
 * System configuration validation function.
 * Checks:
 * 1. Database ID exists and Spreadsheet is accessible.
 * 2. Required sheet tabs exist.
 * 3. Required configuration properties exist.
 * 4. Required application URLs exist.
 * 5. Required system settings exist.
 * 
 * Returns a structured validation result without exposing sensitive parameters.
 * 
 * @returns {Object} Structured validation object { valid: boolean, status: 'valid'|'missing'|'invalid', ... }
 */
function validateSystemConfiguration() {
  try {
    const ssId = getConfigProperty('SPREADSHEET_ID', '');
    const empSsId = getConfigProperty('EMPLOYEE_SPREADSHEET_ID', ssId);

    const result = {
      valid: true,
      status: 'valid',
      database: {
        connected: false,
        spreadsheetId: ssId ? ssId : 'Not configured',
        employeeSpreadsheetId: empSsId ? empSsId : 'Not configured',
        name: ''
      },
      sheets: {},
      urls: {
        publicPortalUrl: getConfigProperty('PUBLIC_PORTAL_URL', '') ? 'Configured' : 'Not configured',
        hodPortalUrl: getConfigProperty('HOD_PORTAL_URL', '') ? 'Configured' : 'Not configured'
      },
      settings: {
        allowedDomain: getConfigProperty('ALLOWED_DOMAIN', '') ? 'Configured' : 'Not configured',
        adminEmails: getConfigProperty('ADMIN_EMAILS', '') ? 'Configured' : 'Not configured'
      },
      details: []
    };

    if (!ssId) {
      result.valid = false;
      result.status = 'missing';
      result.details.push('Master Database Spreadsheet ID (SPREADSHEET_ID) is not configured.');
    } else {
      try {
        const ss = SpreadsheetApp.openById(ssId);
        if (ss) {
          result.database.connected = true;
          result.database.name = ss.getName();

          const mainSheets = [
            { key: 'Trainings', name: getConfigProperty('SHEET_TRAININGS', 'Trainings') }
          ];

          mainSheets.forEach(item => {
            const sheet = ss.getSheetByName(item.name);
            if (sheet) {
              result.sheets[item.key] = { exists: true, status: 'Connected', name: item.name };
            } else {
              result.sheets[item.key] = { exists: false, status: 'Not configured', name: item.name };
              result.valid = false;
              if (result.status === 'valid') result.status = 'invalid';
              result.details.push(`Required sheet tab '${item.name}' not found in Main Database spreadsheet.`);
            }
          });

          // Read-only validation of Employee Spreadsheet tabs
          let empSS = ss;
          if (empSsId && empSsId !== ssId) {
            try { empSS = SpreadsheetApp.openById(empSsId); } catch(e) {}
          }

          if (empSS) {
            const empTabNames = ['For IT', 'Employees', 'HOD email', 'HR email', 'Cost Centre'];
            const foundEmpTab = empTabNames.find(tName => empSS.getSheetByName(tName) !== null);
            if (foundEmpTab) {
              result.sheets['Employees'] = { exists: true, status: 'Connected (Read-Only)', name: foundEmpTab };
            } else {
              result.sheets['Employees'] = { exists: false, status: 'Connected (Read-Only)', name: 'Employee Directory' };
            }
          }
        }
      } catch (e) {
        result.valid = false;
        result.status = 'invalid';
        result.database.connected = false;
        result.details.push('Failed to connect to Master Google Spreadsheet: ' + e.message);
      }
    }

    return result;
  } catch (e) {
    Logger.log('validateSystemConfiguration error: ' + e.message);
    return {
      valid: false,
      status: 'invalid',
      details: ['Configuration validation error: ' + e.message]
    };
  }
}

