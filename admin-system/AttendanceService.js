/**
 * AttendanceService.gs - Session-based Attendance service
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

    const ss = getTrainingDataSpreadsheet(session.TrainingID);
    if (!ss) return err('Could not open per-training sheet for ID: ' + session.TrainingID);

    let attSheet = ss.getSheetByName('Attendance');
    if (!attSheet) {
      attSheet = ss.insertSheet('Attendance');
      attSheet.appendRow(['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName', 'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours', 'Remarks', 'EditedBy', 'EditedAt']);
      attSheet.getRange('A1:O1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      attSheet.setFrozenRows(1);
    }

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

    // Automatically update training stage to 'Attendance In Progress'
    try {
      updateTrainingStage(session.TrainingID, 'Attendance In Progress');
    } catch (e) {}

    invalidateTrainingCaches(session.TrainingID);

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

    const found = findTrainingBySessionId(sessionId);
    if (!found || !found.spreadsheet) return ok([]);

    const attSheet = found.spreadsheet.getSheetByName('Attendance');
    if (!attSheet) return ok([]);

    const rows = sheetToJson(attSheet);
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

    const sessionsRes = getSessions(trainingId);
    const sessionsObj = typeof sessionsRes === 'string' ? JSON.parse(sessionsRes) : sessionsRes;
    const sessions = (sessionsObj && sessionsObj.success && sessionsObj.data) ? sessionsObj.data : [];

    const ss = getTrainingDataSpreadsheet(trainingId);
    const attSheet = ss ? ss.getSheetByName('Attendance') : null;
    const allAtt = attSheet ? sheetToJson(attSheet) : [];

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

    // Locate training sheet containing this attendance ID
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return err('Trainings sheet unavailable.');
    const trainings = sheetToJson(tSheet);

    for (const t of trainings) {
      if (!t.ID) continue;
      const ss = getTrainingDataSpreadsheet(t.ID);
      if (!ss) continue;
      const attSheet = ss.getSheetByName('Attendance');
      if (!attSheet) continue;

      const row = findRowById(attSheet, id);
      if (row !== -1) {
        const headers = attSheet.getRange(1, 1, 1, attSheet.getLastColumn()).getValues()[0];
        const statusCol   = headers.indexOf('Status') + 1;
        const remarksCol  = headers.indexOf('Remarks') + 1;
        const editedByCol = headers.indexOf('EditedBy') + 1;
        const editedAtCol = headers.indexOf('EditedAt') + 1;

        if (statusCol)  attSheet.getRange(row, statusCol).setValue(status || 'Present');
        if (remarksCol) attSheet.getRange(row, remarksCol).setValue(remarks || '');
        if (editedByCol) attSheet.getRange(row, editedByCol).setValue('Admin');
        if (editedAtCol) attSheet.getRange(row, editedAtCol).setValue(now());

        invalidateTrainingCaches(t.ID);

        return ok({ message: 'Attendance record updated successfully.' });
      }
    }

    return err('Attendance record not found.');
  } catch (e) {
    Logger.log('updateAttendanceRecord error: ' + e.message);
    return err('Failed to update record: ' + e.message);
  }
}

/**
 * Read all attendance records for a training ID.
 */
function getAttendance(trainingId) {
  try {
    if (!trainingId) return ok([]);
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return ok([]);
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) return ok([]);

    const rows = sheetToJson(sheet);

    // Group by day
    const days = {};
    rows.forEach(r => {
      const d = r.Day || '1';
      if (!days[d]) days[d] = { day: d, date: r.Date, records: [] };
      days[d].records.push(r);
    });

    return ok(Object.values(days).sort((a, b) => Number(a.day) - Number(b.day)));
  } catch (e) {
    return err('Failed to load attendance: ' + e.message);
  }
}

/**
 * Attendance Summary for Dashboard / Reports.
 */
function getAttendanceSummary(trainingId) {
  try {
    if (!trainingId) return ok({ total: 0, present: 0, absent: 0, late: 0, pct: 0 });
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return ok({ total: 0, present: 0, absent: 0, late: 0, pct: 0 });
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) return ok({ total: 0, present: 0, absent: 0, late: 0, pct: 0 });

    const rows = sheetToJson(sheet);
    const total   = rows.length;
    const present = rows.filter(r => r.Status === 'Present').length;
    const absent  = rows.filter(r => r.Status === 'Absent').length;
    const late    = rows.filter(r => r.Status === 'Late').length;
    const pct     = total > 0 ? Math.round((present / total) * 100) : 0;
    return ok({ total, present, absent, late, pct });
  } catch (e) {
    return err(e.message);
  }
}

