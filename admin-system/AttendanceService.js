/**
 * AttendanceService.gs — Session-based Attendance service
 */

/**
 * Submit attendance for a specific session by a participant.
 * 
 * Supports two parameter patterns:
 * 1. submitAttendance(sessionId, employeeNo, employeeName, department)
 * 2. submitAttendance({ sessionId, employeeNo, employeeName, department })
 * 
 * @returns {string} JSON response indicating success or failure reason
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

    // 1. Run Validation Rules
    const validation = validateAttendance(sessionId, employeeNo);
    if (!validation.valid) {
      return err(validation.message);
    }

    const session = validation.session;
    const empInfo = validation.employee;

    const cleanEmpNo = String(employeeNo).trim();
    const finalEmpName = employeeName || (empInfo ? empInfo.Name : cleanEmpNo);
    const finalDept    = department   || (empInfo ? empInfo.Department : '');

    const attSheet = getSheet(SHEET_NAMES.attendance);
    if (!attSheet) return err('Could not open Attendance sheet.');

    const attId = generateId('ATT');
    const scanTime = now();
    const status = 'Present';

    // Retrieve TrainingCode if available
    let trainingCode = '';
    try {
      const tSheet = getSheet(SHEET_NAMES.trainings);
      if (tSheet) {
        const tRows = sheetToJson(tSheet);
        const parentT = tRows.find(t => String(t.ID || t.TrainingID || '').trim() === String(session.TrainingID).trim());
        if (parentT) trainingCode = parentT.Code || '';
      }
    } catch (e) {}

    // Attendance Row format matching sheet headers:
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
      'QR Code Check-In',
      'System',
      scanTime
    ];

    attSheet.appendRow(newRecord);

    // Automatically update training stage to 'Attendance In Progress' if still 'Created' or 'Participants Imported'
    try {
      updateTrainingStage(session.TrainingID, 'Attendance In Progress');
    } catch (e) {}

    return ok({
      message: `Attendance successfully submitted for ${finalEmpName} (${cleanEmpNo})!`,
      attendanceId: attId,
      sessionId: session.SessionID,
      sessionName: session.SessionName,
      trainingTitle: session.TrainingTitle || '',
      scanTime: scanTime
    });

  } catch (e) {
    Logger.log('submitAttendance error: ' + e.message);
    return err('Failed to submit attendance: ' + e.message);
  }
}

/**
 * Get attendance records for a specific session ID.
 * 
 * @param {string} sessionId - Session ID
 * @returns {string} JSON response with array of attendance records
 */
function getAttendanceBySession(sessionId) {
  try {
    if (!sessionId) return err('Session ID is required.');

    const sheet = getSheet(SHEET_NAMES.attendance);
    if (!sheet) return ok([]);

    const rows = sheetToJson(sheet);
    const filtered = rows.filter(r => String(r.SessionID || '').trim() === String(sessionId).trim());

    return ok(filtered);
  } catch (e) {
    Logger.log('getAttendanceBySession error: ' + e.message);
    return err('Failed to get attendance records: ' + e.message);
  }
}

/**
 * Get attendance records for a training ID, grouped by session.
 * 
 * @param {string} trainingId - Training ID
 * @returns {string} JSON response with sessions and their attendance records
 */
function getAttendanceByTraining(trainingId) {
  try {
    if (!trainingId) return err('Training ID is required.');

    // Fetch sessions
    const sessionsRes = getSessions(trainingId);
    const sessionsObj = typeof sessionsRes === 'string' ? JSON.parse(sessionsRes) : sessionsRes;
    const sessions = (sessionsObj && sessionsObj.success && sessionsObj.data) ? sessionsObj.data : [];

    // Fetch all attendance for training
    const attSheet = getSheet(SHEET_NAMES.attendance);
    const allAtt = attSheet ? sheetToJson(attSheet).filter(r => String(r.TrainingID || '').trim() === String(trainingId).trim()) : [];

    const grouped = sessions.map(s => {
      const records = allAtt.filter(a => String(a.SessionID || '').trim() === String(s.SessionID).trim());
      return {
        session: s,
        records: records,
        totalRecords: records.length,
        presentCount: records.filter(r => String(r.Status).toLowerCase() === 'present').length
      };
    });

    return ok(grouped);
  } catch (e) {
    Logger.log('getAttendanceByTraining error: ' + e.message);
    return err('Failed to load training attendance: ' + e.message);
  }
}

/**
 * Update single attendance record's Status and Remarks.
 */
function updateAttendanceRecord(id, status, remarks) {
  try {
    if (!id) return err('Record ID is required.');
    const sheet = getSheet(SHEET_NAMES.attendance);
    if (!sheet) return err('Attendance sheet not found.');

    const row = findRowById(sheet, id);
    if (row === -1) return err('Attendance record not found.');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol   = headers.indexOf('Status') + 1;
    const remarksCol  = headers.indexOf('Remarks') + 1;
    const editedByCol = headers.indexOf('EditedBy') + 1;
    const editedAtCol = headers.indexOf('EditedAt') + 1;

    if (statusCol)  sheet.getRange(row, statusCol).setValue(status || 'Present');
    if (remarksCol) sheet.getRange(row, remarksCol).setValue(remarks || '');
    if (editedByCol) sheet.getRange(row, editedByCol).setValue('Admin');
    if (editedAtCol) sheet.getRange(row, editedAtCol).setValue(now());

    return ok({ message: 'Attendance record updated successfully.' });
  } catch (e) {
    Logger.log('updateAttendanceRecord error: ' + e.message);
    return err('Failed to update record: ' + e.message);
  }
}
