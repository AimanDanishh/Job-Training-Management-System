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
    let sessionId, employeeNo, employeeName, department, remarks, editedBy, reqStatus;

    if (typeof arg1 === 'object' && arg1 !== null) {
      sessionId    = arg1.sessionId    || arg1.SessionID;
      employeeNo   = arg1.employeeNo   || arg1.EmployeeNo || arg1.employeeId || arg1.EmployeeID;
      employeeName = arg1.employeeName || arg1.EmployeeName;
      department   = arg1.department   || arg1.Department;
      remarks      = arg1.remarks      || arg1.Remarks;
      editedBy     = arg1.editedBy     || arg1.EditedBy;
      reqStatus    = arg1.status       || arg1.Status;
    } else {
      sessionId    = arg1;
      employeeNo   = arg2;
      employeeName = arg3;
      department   = arg4;
    }

    const finalRemarks = remarks || 'Manual Attendance';
    const finalEditedBy = editedBy || (finalRemarks.indexOf('Manual') !== -1 ? 'Admin' : 'System');

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
    
    // Determine status (Present, Late, or explicitly requested status e.g. Absent)
    let determinedStatus = reqStatus || 'Present';
    if (!reqStatus && session) {
      determinedStatus = determineAttendanceStatus(session, new Date());
    }

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

    // Check if an existing row for this participant and session exists (e.g. was marked Absent on deactivation)
    const existingRow = findRowByEmpAndSession(attSheet, cleanEmpNo, session.SessionID);
    if (existingRow !== -1) {
      const headers = attSheet.getRange(1, 1, 1, attSheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim().toLowerCase());
      const statusCol   = headers.findIndex(h => h === 'status') + 1;
      const scanTimeCol = headers.findIndex(h => h === 'scantime' || h === 'checkin') + 1;
      const remarksCol  = headers.findIndex(h => h === 'remarks') + 1;
      const editedByCol = headers.findIndex(h => h === 'editedby') + 1;
      const editedAtCol = headers.findIndex(h => h === 'editedat') + 1;

      if (statusCol)   attSheet.getRange(existingRow, statusCol).setValue(determinedStatus);
      if (scanTimeCol) attSheet.getRange(existingRow, scanTimeCol).setValue(scanTime);
      if (remarksCol)  attSheet.getRange(existingRow, remarksCol).setValue(finalRemarks);
      if (editedByCol) attSheet.getRange(existingRow, editedByCol).setValue(finalEditedBy);
      if (editedAtCol) attSheet.getRange(existingRow, editedAtCol).setValue(scanTime);
    } else {
      const newRecord = [
        attId,
        session.SessionID,
        session.TrainingID,
        cleanEmpNo,
        finalEmpName,
        finalDept,
        scanTime,
        determinedStatus,
        trainingCode,
        session.SessionName || '',
        session.SessionDate || '',
        0,
        finalRemarks,
        finalEditedBy,
        scanTime
      ];
      attSheet.appendRow(newRecord);
    }
    SpreadsheetApp.flush();

    // Automatically update training stage to 'Attendance In Progress'
    try {
      updateTrainingStage(session.TrainingID, 'Attendance In Progress');
    } catch (e) {}

    invalidateTrainingCaches(session.TrainingID);

    return ok({
      message: `Attendance successfully recorded for ${finalEmpName} (${cleanEmpNo})! Status: ${determinedStatus}`,
      attendanceId: attId,
      status: determinedStatus,
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

    try {
      const s = getSession(sessionId);
      const sObj = typeof s === 'string' ? JSON.parse(s) : s;
      if (sObj && sObj.data && isDeactivatedStatus(sObj.data.QRStatus)) {
        markUnscannedParticipantsAbsent(sessionId, sObj.data.TrainingID);
      }
    } catch(e) {}

    const attSheet = found.spreadsheet.getSheetByName('Attendance');
    if (!attSheet) return ok([]);

    const rows = sheetToJson(attSheet);
    const filtered = rows.filter(r => String(r.SessionID || '').trim().toLowerCase() === String(sessionId).trim().toLowerCase());

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

    // Ensure any deactivated sessions have unscanned participants recorded as Absent
    try {
      syncDeactivatedSessionsAbsent(trainingId);
    } catch (sErr) {
      Logger.log('syncDeactivatedSessionsAbsent in getAttendanceByTraining error: ' + sErr.message);
    }

    const sessionsRes = getSessions(trainingId);
    const sessionsObj = typeof sessionsRes === 'string' ? JSON.parse(sessionsRes) : sessionsRes;
    const sessions = (sessionsObj && sessionsObj.success && sessionsObj.data) ? sessionsObj.data : [];

    const ss = getTrainingDataSpreadsheet(trainingId);
    const attSheet = ss ? ss.getSheetByName('Attendance') : null;
    const allAtt = attSheet ? sheetToJson(attSheet) : [];

    const enrolledParticipants = getEnrolledParticipantsForTraining(trainingId);
    const totalEnrolled = enrolledParticipants.length;

    const grouped = sessions.map(s => {
      const records = allAtt.filter(a => String(a.SessionID || '').trim().toLowerCase() === String(s.SessionID || '').trim().toLowerCase());
      const presentCount = records.filter(r => {
        const st = String(r.Status || '').trim().toLowerCase();
        return st === 'present' || st === 'late';
      }).length;
      const totalExpected = Math.max(totalEnrolled, records.length);
      return {
        session: s,
        records: records,
        totalRecords: totalExpected,
        presentCount: presentCount
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
    const enrolledParts = getEnrolledParticipantsForTraining(trainingId);
    const enrolledCount = enrolledParts.length;

    const present = rows.filter(r => String(r.Status || '').trim().toLowerCase() === 'present').length;
    const late    = rows.filter(r => String(r.Status || '').trim().toLowerCase() === 'late').length;
    let absent    = rows.filter(r => String(r.Status || '').trim().toLowerCase() === 'absent').length;

    const sessionsRes = getSessions(trainingId);
    const sessionsObj = typeof sessionsRes === 'string' ? JSON.parse(sessionsRes) : sessionsRes;
    const sessions = (sessionsObj && sessionsObj.success && sessionsObj.data) ? sessionsObj.data : [];
    const sessionCount = Math.max(1, sessions.length);

    const totalExpected = enrolledCount > 0 ? (enrolledCount * sessionCount) : rows.length;
    const total = Math.max(totalExpected, rows.length);

    if (absent === 0 && total > (present + late)) {
      absent = total - (present + late);
    }

    const pct = total > 0 ? Math.round(((present + late) / total) * 100) : 0;
    return ok({ total, present, absent, late, pct });
  } catch (e) {
    return err(e.message);
  }
}

// --- Helper Functions for Attendance & Session Synchronization ---------------

/**
 * Checks if a session QR status represents deactivated/closed.
 */
function isDeactivatedStatus(status) {
  if (!status) return false;
  const clean = String(status).trim().toLowerCase();
  return clean === 'deactivate' || clean === 'deactivated' || clean === 'inactive' || clean === 'expired' || clean === 'disabled';
}

/**
 * Evaluates whether an attendance check-in is Late based on session start time and grace period.
 */
function determineAttendanceStatus(session, scanDate) {
  if (!session) return 'Present';
  const nowObj = scanDate || new Date();
  try {
    const startStr = String(session.StartTime || '').trim();
    if (!startStr) return 'Present';
    const match = startStr.match(/(\d{1,2}):(\d{2})(?:\s*([AaPp][Mm]))?/);
    if (!match) return 'Present';

    let startH = parseInt(match[1], 10);
    const startM = parseInt(match[2], 10);
    const ampm = match[3];
    if (ampm) {
      if (ampm.toLowerCase() === 'pm' && startH < 12) startH += 12;
      if (ampm.toLowerCase() === 'am' && startH === 12) startH = 0;
    }

    const startTotalMinutes = startH * 60 + startM;
    const scanTotalMinutes = nowObj.getHours() * 60 + nowObj.getMinutes();

    // If session date is in the past, mark as Late
    if (session.SessionDate) {
      const sDateStr = String(session.SessionDate).replace(/GMT.*$/, '').replace(/\(.*\)$/, '').trim();
      const sDate = new Date(sDateStr);
      if (!isNaN(sDate.getTime())) {
        const todayMid = new Date(nowObj.getFullYear(), nowObj.getMonth(), nowObj.getDate()).getTime();
        const sMid = new Date(sDate.getFullYear(), sDate.getMonth(), sDate.getDate()).getTime();
        if (todayMid > sMid) return 'Late';
      }
    }

    // Grace period of 15 minutes after session start time
    if (scanTotalMinutes > (startTotalMinutes + 15)) {
      return 'Late';
    }
  } catch (e) {
    Logger.log('determineAttendanceStatus error: ' + e.message);
  }
  return 'Present';
}

/**
 * Finds row number in Attendance sheet matching Employee ID and Session ID.
 */
function findRowByEmpAndSession(sheet, empNo, sessionId) {
  if (!sheet || sheet.getLastRow() < 2) return -1;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
  const empCol = headers.findIndex(h => h === 'employeeno' || h === 'employeeid' || h === 'empid' || h === 'id' || h === 'staffid');
  const sessCol = headers.findIndex(h => h === 'sessionid' || h === 'session' || h === 'sessioncode');
  if (empCol === -1 || sessCol === -1) return -1;

  const targetEmp = String(empNo).trim().toLowerCase();
  const targetSess = String(sessionId).trim().toLowerCase();
  for (let r = 1; r < data.length; r++) {
    const rEmp = String(data[r][empCol] || '').trim().toLowerCase();
    const rSess = String(data[r][sessCol] || '').trim().toLowerCase();
    if ((rEmp === targetEmp || (typeof isSameEmployeeId === 'function' && isSameEmployeeId(rEmp, targetEmp))) &&
        rSess === targetSess) {
      return r + 1;
    }
  }
  return -1;
}

/**
 * Reliably fetches enrolled participants for a training programme across all storage tiers.
 */
function getEnrolledParticipantsForTraining(trainingId) {
  if (!trainingId) return [];
  try {
    if (typeof getTrainingParticipantsList === 'function') {
      const list = getTrainingParticipantsList(trainingId);
      if (Array.isArray(list) && list.length > 0) return list;
    }
  } catch (e) {}

  try {
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (ss) {
      const partSheet = ss.getSheetByName('Participants') || ss.getSheetByName('TrainingParticipants') || ss.getSheetByName('Participant');
      if (partSheet && partSheet.getLastRow() >= 2) {
        const rows = sheetToJson(partSheet);
        if (rows && rows.length > 0) return rows;
      }
    }
  } catch (e) {}

  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet) {
      const tRows = sheetToJson(tSheet);
      const t = tRows.find(x => String(x.ID || '').trim().toLowerCase() === String(trainingId).trim().toLowerCase());
      if (t && t.ParticipantsSheetID) {
        const pSs = SpreadsheetApp.openById(t.ParticipantsSheetID);
        if (pSs) {
          const s = pSs.getSheets()[0];
          if (s && s.getLastRow() >= 2) {
            const rows = sheetToJson(s);
            if (rows && rows.length > 0) return rows;
          }
        }
      }
    }
  } catch (e) {}

  return [];
}

/**
 * Deduplicates attendance sheet rows to guarantee at most one record per (SessionID, EmployeeID).
 * Keeps Present/Late records over Absent records, and removes redundant duplicates.
 */
function cleanupDuplicateAttendanceRows(attSheet, trainingId) {
  if (!attSheet || attSheet.getLastRow() < 3) return 0;
  try {
    const data = attSheet.getDataRange().getValues();
    if (data.length < 3) return 0;

    const rawHeaders = data[0];
    const cleanHeaders = rawHeaders.map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));

    const sessCol = cleanHeaders.findIndex(h => h === 'sessionid' || h === 'session' || h === 'sessioncode');
    const empCol  = cleanHeaders.findIndex(h => h === 'employeeno' || h === 'employeeid' || h === 'empid' || h === 'staffid' || h === 'id');
    const statCol = cleanHeaders.findIndex(h => h === 'status');

    if (sessCol === -1 || empCol === -1) return 0;

    // Track seen entries: key => { rowNumber (1-based), status }
    const seen = new Map();
    const rowsToDelete = [];

    for (let r = 1; r < data.length; r++) {
      const sess = String(data[r][sessCol] || '').trim().toLowerCase();
      const emp  = String(data[r][empCol]  || '').trim().toLowerCase();
      const stat = statCol !== -1 ? String(data[r][statCol] || '').trim().toLowerCase() : '';
      if (!sess || !emp) continue;

      const key = `${sess}___${emp}`;
      if (!seen.has(key)) {
        seen.set(key, { rowNumber: r + 1, status: stat });
      } else {
        const existing = seen.get(key);
        // If current row is Present/Late and existing was Absent, keep current and delete existing
        if ((stat === 'present' || stat === 'late') && existing.status === 'absent') {
          rowsToDelete.push(existing.rowNumber);
          seen.set(key, { rowNumber: r + 1, status: stat });
        } else {
          // Otherwise current row is duplicate, mark for deletion
          rowsToDelete.push(r + 1);
        }
      }
    }

    if (rowsToDelete.length > 0) {
      // Delete from bottom to top to preserve correct row indexing
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach(rowNum => {
        try { attSheet.deleteRow(rowNum); } catch(e) {}
      });
      SpreadsheetApp.flush();
      if (trainingId) invalidateTrainingCaches(trainingId);
    }
    return rowsToDelete.length;
  } catch (e) {
    Logger.log('cleanupDuplicateAttendanceRows error: ' + e.message);
    return 0;
  }
}

/**
 * When a session is deactivated, mark all enrolled participants who haven't scanned as Absent.
 * Serialized via script lock and protected by duplicate detection.
 */
function markUnscannedParticipantsAbsent(sessionId, trainingIdHint) {
  if (!sessionId) return;
  const cleanSessionId = String(sessionId).trim();
  const lowerSessionId = cleanSessionId.toLowerCase();

  const lock = LockService.getScriptLock();
  let hasLock = false;
  try {
    hasLock = lock.tryLock(15000);
  } catch(lErr) {}

  try {
    let session = null;
    try {
      const sRes = getSession(cleanSessionId);
      const sObj = typeof sRes === 'string' ? JSON.parse(sRes) : sRes;
      if (sObj && sObj.success && sObj.data) session = sObj.data;
    } catch (e) {}

    const trainingId = (session && session.TrainingID) || trainingIdHint;
    if (!trainingId) return;

    const participants = getEnrolledParticipantsForTraining(trainingId);
    if (!participants || participants.length === 0) return;

    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return;

    let attSheet = ss.getSheetByName('Attendance');
    if (!attSheet) {
      attSheet = ss.insertSheet('Attendance');
      attSheet.appendRow(['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName', 'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours', 'Remarks', 'EditedBy', 'EditedAt']);
      attSheet.getRange('A1:O1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      attSheet.setFrozenRows(1);
    }

    // 1. Clean up any existing duplicate rows first
    cleanupDuplicateAttendanceRows(attSheet, trainingId);

    // 2. Read raw sheet data directly to avoid any key normalization mismatch
    const data = attSheet.getDataRange().getValues();
    if (data.length < 1) return;

    const rawHeaders = data[0];
    const cleanHeaders = rawHeaders.map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));

    const sessCol = cleanHeaders.findIndex(h => h === 'sessionid' || h === 'session' || h === 'sessioncode');
    const empCol  = cleanHeaders.findIndex(h => h === 'employeeno' || h === 'employeeid' || h === 'empid' || h === 'staffid' || h === 'id');

    const recordedEmpIds = new Set();
    if (sessCol !== -1 && empCol !== -1) {
      for (let r = 1; r < data.length; r++) {
        const rowSess = String(data[r][sessCol] || '').trim().toLowerCase();
        const rowEmp  = String(data[r][empCol]  || '').trim().toLowerCase();
        if (rowSess === lowerSessionId && rowEmp) {
          recordedEmpIds.add(rowEmp);
        }
      }
    }

    let trainingCode = (session && session.TrainingCode) || '';
    if (!trainingCode) {
      try {
        const tSheet = getSheet(SHEET_NAMES.trainings);
        if (tSheet) {
          const tRows = sheetToJson(tSheet);
          const pt = tRows.find(t => String(t.ID || '').trim().toLowerCase() === String(trainingId).trim().toLowerCase());
          if (pt) trainingCode = pt.Code || '';
        }
      } catch (e) {}
    }

    const sessionName = (session && session.SessionName) || 'Session';
    const sessionDate = (session && session.SessionDate) || '';

    let addedCount = 0;
    participants.forEach(p => {
      const pId = String(p.EmployeeID || p.ID || p.EmployeeNo || '').trim();
      if (!pId) return;

      const cleanPId = pId.toLowerCase();
      let alreadyRecorded = recordedEmpIds.has(cleanPId);
      if (!alreadyRecorded) {
        for (const recId of recordedEmpIds) {
          if (typeof isSameEmployeeId === 'function' && isSameEmployeeId(recId, cleanPId)) {
            alreadyRecorded = true;
            break;
          }
        }
      }
      if (alreadyRecorded) return;

      const pName = p.EmployeeName || p.Name || pId;
      const pDept = p.Department || p.CostCentre || p.Dept || '';

      const recordObj = {
        AttendanceID: generateId('ATT'),
        SessionID: cleanSessionId,
        TrainingID: trainingId,
        EmployeeNo: pId,
        EmployeeName: pName,
        Department: pDept,
        ScanTime: '',
        Status: 'Absent',
        TrainingCode: trainingCode,
        Day: sessionName,
        Date: sessionDate,
        Hours: 0,
        Remarks: 'Absent - Did Not Scan',
        EditedBy: 'System',
        EditedAt: now()
      };

      const row = [];
      cleanHeaders.forEach((h, colIdx) => {
        switch (h) {
          case 'attendanceid': case 'id': row.push(recordObj.AttendanceID); break;
          case 'sessionid': case 'session_id': row.push(recordObj.SessionID); break;
          case 'trainingid': case 'training_id': row.push(recordObj.TrainingID); break;
          case 'employeeno': case 'employeeid': case 'empid': row.push(recordObj.EmployeeNo); break;
          case 'employeename': case 'name': row.push(recordObj.EmployeeName); break;
          case 'department': case 'costcentre': case 'dept': row.push(recordObj.Department); break;
          case 'scantime': case 'checkin': case 'scan_time': row.push(recordObj.ScanTime); break;
          case 'status': row.push(recordObj.Status); break;
          case 'trainingcode': case 'code': row.push(recordObj.TrainingCode); break;
          case 'day': row.push(recordObj.Day); break;
          case 'date': row.push(recordObj.Date); break;
          case 'hours': row.push(recordObj.Hours); break;
          case 'remarks': row.push(recordObj.Remarks); break;
          case 'editedby': row.push(recordObj.EditedBy); break;
          case 'editedat': row.push(recordObj.EditedAt); break;
          default: row.push(''); break;
        }
      });

      attSheet.appendRow(row);
      recordedEmpIds.add(cleanPId);
      addedCount++;
    });

    if (addedCount > 0) {
      SpreadsheetApp.flush();
      invalidateTrainingCaches(trainingId);
    }
  } catch (e) {
    Logger.log('markUnscannedParticipantsAbsent error: ' + e.message);
  } finally {
    if (hasLock) {
      try { lock.releaseLock(); } catch(e) {}
    }
  }
}

/**
 * Scans all sessions for a training programme and ensures unscanned participants in deactivated sessions are marked Absent.
 */
function syncDeactivatedSessionsAbsent(trainingId) {
  try {
    if (!trainingId) return;

    // Clean any duplicates first
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (ss) {
      const attSheet = ss.getSheetByName('Attendance');
      if (attSheet) cleanupDuplicateAttendanceRows(attSheet, trainingId);
    }

    const sessionsRes = getSessions(trainingId);
    const sessionsObj = typeof sessionsRes === 'string' ? JSON.parse(sessionsRes) : sessionsRes;
    const sessions = (sessionsObj && sessionsObj.success && sessionsObj.data) ? sessionsObj.data : [];

    sessions.forEach(s => {
      const status = String(s.QRStatus || '').trim().toLowerCase();
      if (isDeactivatedStatus(status)) {
        markUnscannedParticipantsAbsent(s.SessionID, trainingId);
      }
    });
  } catch (e) {
    Logger.log('syncDeactivatedSessionsAbsent error: ' + e.message);
  }
}

