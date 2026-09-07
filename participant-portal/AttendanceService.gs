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

    const found = findTrainingBySessionId(cleanSessionId);
    if (!found || !found.session) {
      return err(`Session (${cleanSessionId}) not found in the database. Please verify the session QR code.`);
    }

    const session  = found.session;
    const training = found.training;

    const formatTimeClean = (val, defaultVal) => {
      if (!val) return defaultVal;
      const str = String(val).replace(/GMT.*$/, '').replace(/\(.*\)$/, '').trim();
      const simpleMatch = str.match(/^(\d{1,2}:\d{2})/);
      if (simpleMatch && !str.includes('1899') && !str.includes('Singapore') && !str.includes('Standard')) return simpleMatch[1];
      const d = new Date(str);
      if (!isNaN(d.getTime())) {
        return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
      }
      const match = str.match(/(\d{2}:\d{2})/);
      if (match) return match[1];
      return defaultVal;
    };

    return ok({
      SessionID:     session.SessionID || session['Session ID'] || session.SessionId || cleanSessionId,
      TrainingID:    session.TrainingID || session['Training ID'] || session.trainingId || (training ? (training.ID || training.TrainingID) : ''),
      SessionName:   session.SessionName || session['Session Name'] || 'Session Check-In',
      SessionDate:   session.SessionDate || session['Session Date'] || '—',
      StartTime:     formatTimeClean(session.StartTime || session['Start Time'], '09:00'),
      EndTime:       formatTimeClean(session.EndTime || session['End Time'], '17:00'),
      QRStatus:      session.QRStatus || session['QR Status'] || 'Active',
      TrainingTitle: training ? (training.Name || training.TrainingTitle || training['Training Title'] || 'Training Programme') : 'Training Programme',
      TrainingCode:  training ? (training.Code || training.TrainingCode || training['Training Code'] || '') : ''
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
    const finalDept    = (empInfo && (empInfo.Department || empInfo.CostCentre)) ? (empInfo.Department || empInfo.CostCentre) : (department || '');
    const trainingCode = training ? (training.Code || '') : '';

    const ss = (validation && validation.spreadsheet) ? validation.spreadsheet : (getTrainingDataSpreadsheet(session.TrainingID) || getSpreadsheet());
    if (!ss) return err('Could not open training data spreadsheet.');

    let attSheet = ss.getSheetByName('Attendance');
    if (!attSheet) {
      attSheet = ss.insertSheet('Attendance');
      attSheet.appendRow(['AttendanceID', 'SessionID', 'TrainingID', 'EmployeeNo', 'EmployeeName', 'Department', 'ScanTime', 'Status', 'TrainingCode', 'Day', 'Date', 'Hours', 'Remarks', 'EditedBy', 'EditedAt']);
      attSheet.getRange('A1:O1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      attSheet.setFrozenRows(1);
    } else {
      ensureAttendanceSheetColumns(attSheet);
    }

    const attId = generateId('ATT');
    const scanTime = now();
    
    // Evaluate if check-in is Late based on session start time
    let status = 'Present';
    try {
      if (session && session.StartTime) {
        const startStr = String(session.StartTime).trim();
        const match = startStr.match(/(\d{1,2}):(\d{2})(?:\s*([AaPp][Mm]))?/);
        if (match) {
          let startH = parseInt(match[1], 10);
          const startM = parseInt(match[2], 10);
          const ampm = match[3];
          if (ampm) {
            if (ampm.toLowerCase() === 'pm' && startH < 12) startH += 12;
            if (ampm.toLowerCase() === 'am' && startH === 12) startH = 0;
          }
          const startTotalMinutes = startH * 60 + startM;
          const scanMinutes = new Date().getHours() * 60 + new Date().getMinutes();
          if (scanMinutes > (startTotalMinutes + 15)) {
            status = 'Late';
          }
        }
      }
    } catch(tErr) {}

    // Check if an existing row for this participant was marked 'Absent'
    let updatedExisting = false;
    try {
      if (attSheet.getLastRow() >= 2) {
        const data = attSheet.getDataRange().getValues();
        const headers = data[0].map(h => String(h || '').trim().toLowerCase());
        const empCol = headers.findIndex(h => h === 'employeeno' || h === 'employeeid' || h === 'empid');
        const sessCol = headers.findIndex(h => h === 'sessionid');
        const statCol = headers.findIndex(h => h === 'status');
        const scanCol = headers.findIndex(h => h === 'scantime' || h === 'checkin');
        const remCol  = headers.findIndex(h => h === 'remarks');

        if (empCol !== -1 && sessCol !== -1) {
          for (let r = 1; r < data.length; r++) {
            if (String(data[r][empCol] || '').trim().toLowerCase() === cleanEmpNo.toLowerCase() &&
                String(data[r][sessCol] || '').trim().toLowerCase() === String(session.SessionID || '').trim().toLowerCase()) {
              if (statCol !== -1) attSheet.getRange(r + 1, statCol + 1).setValue(status);
              if (scanCol !== -1) attSheet.getRange(r + 1, scanCol + 1).setValue(scanTime);
              if (remCol !== -1)  attSheet.getRange(r + 1, remCol + 1).setValue('QR Code Public Check-In');
              updatedExisting = true;
              break;
            }
          }
        }
      }
    } catch(e) {}

    if (!updatedExisting) {
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
    }

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

