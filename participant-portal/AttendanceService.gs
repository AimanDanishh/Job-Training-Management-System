/**
 * AttendanceService.gs — Session Attendance Service for Public Participant Portal
 */

/**
 * Retrieve session details for displaying the attendance check-in form to public users
 */
function getSessionInfo(sessionId) {
  try {
    if (!sessionId || String(sessionId).trim() === '') {
      return err('Session ID parameter is missing.');
    }
    const cleanSessionId = String(sessionId).trim();

    const sSheet = getSheet(SHEET_NAMES.trainingSessions);
    if (!sSheet) return err('Could not access Training Sessions sheet.');

    const sessions = sheetToJson(sSheet);
    const session = sessions.find(s => String(s.SessionID || '').trim() === cleanSessionId);

    if (!session) return err('Session not found.');

    // Fetch parent training info
    let trainingTitle = 'Training Programme';
    let trainingCode = '';
    try {
      const tSheet = getSheet(SHEET_NAMES.trainings);
      if (tSheet) {
        const trainings = sheetToJson(tSheet);
        const parentT = trainings.find(t => String(t.ID || t.TrainingID || '').trim() === String(session.TrainingID).trim());
        if (parentT) {
          trainingTitle = parentT.Name || parentT.TrainingTitle || trainingTitle;
          trainingCode  = parentT.Code || '';
        }
      }
    } catch (e) {}

    return ok({
      SessionID:     session.SessionID,
      TrainingID:    session.TrainingID,
      SessionName:   session.SessionName || 'Session Check-In',
      SessionDate:   session.SessionDate || '—',
      StartTime:     session.StartTime || '09:00',
      EndTime:       session.EndTime || '17:00',
      QRStatus:      session.QRStatus || 'Active',
      TrainingTitle: trainingTitle,
      TrainingCode:  trainingCode
    });
  } catch (e) {
    Logger.log('getSessionInfo error: ' + e.message);
    return err('Failed to load session details: ' + e.message);
  }
}

/**
 * Submit attendance for a session after server-side validation.
 * 
 * Supports object input: { sessionId, employeeNo, employeeName, department }
 */
function submitAttendance(arg1, arg2, arg3, arg4) {
  try {
    let sessionId, employeeNo, employeeName, department;

    if (typeof arg1 === 'object' && arg1 !== null) {
      sessionId    = arg1.sessionId    || arg1.SessionID;
      employeeNo   = arg1.employeeNo   || arg1.EmployeeNo || arg1.employeeId || arg1.EmployeeID;
      employeeName = arg1.employeeName || arg1.EmployeeName;
      department   = arg1.department   || arg1.Department;
    } else {
      sessionId    = arg1;
      employeeNo   = arg2;
      employeeName = arg3;
      department   = arg4;
    }

    // 1. Run Server-Side Security & Enrollment Validations
    const validation = validatePublicAttendance(sessionId, employeeNo);
    if (!validation.valid) {
      if (validation.duplicate) {
        const empName = (validation.employee && validation.employee.Name) ? validation.employee.Name : (employeeName || employeeNo);
        return ok({
          alreadyRecorded: true,
          message: `Attendance has ALREADY been recorded for ${empName} (${employeeNo}) for this session!`,
          scanTime: validation.scanTime || 'Earlier',
          employeeName: empName,
          employeeNo: employeeNo,
          sessionId: sessionId
        });
      }
      return err(validation.message);
    }

    const session  = validation.session;
    const empInfo  = validation.employee;
    const training = validation.training;

    const cleanEmpNo   = String(employeeNo).trim();
    const finalEmpName = (empInfo && empInfo.Name) ? empInfo.Name : (employeeName || cleanEmpNo);
    const finalDept    = (empInfo && empInfo.Department) ? empInfo.Department : (department || '');
    const trainingCode = training ? (training.Code || '') : '';

    const attSheet = getSheet(SHEET_NAMES.attendance);
    if (!attSheet) return err('Could not open Attendance sheet.');

    const attId = generateId('ATT');
    const scanTime = now();
    const status = 'Present';

    // Append to existing Attendance sheet columns:
    // ['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName', 'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours', 'Remarks', 'EditedBy', 'EditedAt']
    const newRecord = [
      attId,
      session.SessionID,
      session.TrainingID,
      cleanEmpNo,
      finalEmpName,
      finalDept,
      scanTime,
      status,
      trainingCode,
      session.SessionName || '',
      session.SessionDate || '',
      0,
      'QR Code Public Check-In',
      'Public Portal',
      scanTime
    ];

    attSheet.appendRow(newRecord);

    return ok({
      message: `Attendance successfully recorded for ${finalEmpName} (${cleanEmpNo})!`,
      attendanceId: attId,
      sessionId: session.SessionID,
      sessionName: session.SessionName,
      trainingTitle: session.TrainingTitle || (training ? training.Name : ''),
      scanTime: scanTime
    });

  } catch (e) {
    Logger.log('submitAttendance error: ' + e.message);
    return err('Failed to submit attendance: ' + e.message);
  }
}

/**
 * Real-time employee lookup for attendance check-in
 */
function lookupEmployeeInfo(employeeNo) {
  return getEmployeeDetails(employeeNo);
}
