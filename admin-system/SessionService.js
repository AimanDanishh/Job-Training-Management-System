/**
 * SessionService.gs - Training Sessions management service
 */

/**
 * Generate sequential Session ID (e.g., SES0001, SES0002, ...)
 */
function generateSessionId() {
  let maxNum = 0;
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet) {
      const trainings = sheetToJson(tSheet);
      for (const t of trainings) {
        if (!t.ID) continue;
        const ss = getTrainingDataSpreadsheet(t.ID);
        if (!ss) continue;
        const sessSheet = ss.getSheetByName('TrainingSessions');
        if (!sessSheet) continue;
        const data = sessSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const id = String(data[i][0]).trim();
          if (id.startsWith('SES')) {
            const num = parseInt(id.replace('SES', ''), 10);
            if (!isNaN(num) && num > maxNum) {
              maxNum = num;
            }
          }
        }
      }
    }
  } catch (e) {}
  const nextNum = maxNum + 1;
  return 'SES' + String(nextNum).padStart(4, '0');
}

/**
 * Create a new Training Session and automatically generate Attendance URL & QR Code.
 * 
 * @param {Object} data - { TrainingID, SessionName, SessionDate, StartTime, EndTime, QRStatus }
 * @returns {string} JSON response with created session object
 */
function createSession(data) {
  try {
    if (!data.TrainingID) return err('Training ID is required.');
    if (!data.SessionName) return err('Session Name is required.');

    // Enforce Approval Check: QR session creation is only allowed for Approved trainings
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet) {
      const trainings = sheetToJson(tSheet);
      const training = trainings.find(t => String(t.ID || '').trim() === String(data.TrainingID).trim());
      if (training) {
        const appStatus = String(training.ApprovalStatus || '').trim().toLowerCase();
        if (appStatus && appStatus !== 'approved' && appStatus !== 'auto-approved') {
          return err(`QR session creation is only allowed for APPROVED training requisitions. Current approval status: '${training.ApprovalStatus || 'Pending Approval'}'.`);
        }
      }
    }

    const ss = getTrainingDataSpreadsheet(data.TrainingID);
    if (!ss) return err('Could not open per-training sheet for ID: ' + data.TrainingID);

    let sheet = ss.getSheetByName('TrainingSessions');
    if (!sheet) {
      sheet = ss.insertSheet('TrainingSessions');
      sheet.appendRow(['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate']);
      sheet.getRange('A1:J1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const sessionId = generateSessionId();
    const attendanceUrl = generateAttendanceURL(sessionId);
    const qrCodeUrl = generateQRCode(attendanceUrl);
    const timeNow = now();

    const newSession = {
      SessionID:     sessionId,
      TrainingID:    data.TrainingID,
      SessionName:   data.SessionName,
      SessionDate:   data.SessionDate || formatDate(new Date()),
      StartTime:     data.StartTime || '09:00',
      EndTime:       data.EndTime || '17:00',
      AttendanceURL: attendanceUrl,
      QRCodeURL:     qrCodeUrl,
      QRStatus:      data.QRStatus || 'Active',
      CreatedDate:   timeNow
    };

    const sessionRow = [
      newSession.SessionID,
      newSession.TrainingID,
      newSession.SessionName,
      newSession.SessionDate,
      newSession.StartTime,
      newSession.EndTime,
      newSession.AttendanceURL,
      newSession.QRCodeURL,
      newSession.QRStatus,
      newSession.CreatedDate
    ];

    sheet.appendRow(sessionRow);

    // Ensure ParticipantsSheetID and SessionsSheetID are persisted to Trainings row
    try {
      if (tSheet) {
        const headers = ensureTrainingSheetColumns(tSheet);
        const tRow = findRowById(tSheet, data.TrainingID);
        if (tRow !== -1) {
          const colIdx = headers.indexOf('ParticipantsSheetID') + 1;
          if (colIdx > 0 && !tSheet.getRange(tRow, colIdx).getValue()) {
            tSheet.getRange(tRow, colIdx).setValue(ss.getId());
          }
          const sColIdx = headers.indexOf('SessionsSheetID') + 1;
          if (sColIdx > 0 && !tSheet.getRange(tRow, sColIdx).getValue()) {
            tSheet.getRange(tRow, sColIdx).setValue(ss.getId());
          }
        }
      }
    } catch(persistErr) {
      Logger.log('Could not persist sheet ID to Trainings: ' + persistErr.message);
    }

    // Sync session record to central TrainingSessions tab in Main Database
    try {
      const mainSs = getSpreadsheet();
      if (mainSs) {
        let centralSheet = mainSs.getSheetByName('TrainingSessions');
        if (!centralSheet) {
          centralSheet = mainSs.insertSheet('TrainingSessions');
          centralSheet.appendRow(['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate']);
          centralSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
          centralSheet.setFrozenRows(1);
        }
        centralSheet.appendRow(sessionRow);
      }
    } catch(cErr) {
      Logger.log('Central TrainingSessions append error: ' + cErr.message);
    }

    // Automatically update lifecycle stage to 'Attendance In Progress' when session attendance is created
    try {
      updateTrainingStage(data.TrainingID, 'Attendance In Progress');
    } catch (e) {
      Logger.log('Auto update stage error: ' + e.message);
    }

    invalidateTrainingCaches(data.TrainingID);

    return ok(newSession);
  } catch (e) {
    Logger.log('createSession error: ' + e.message);
    return err('Failed to create session: ' + e.message);
  }
}

/**
 * Retrieve all sessions for a specific Training ID (or all sessions if trainingId omitted).
 * 
 * @param {string} [trainingId] - Optional Training ID filter
 * @returns {string} JSON response with array of session objects
 */
function getSessions(trainingId) {
  try {
    if (trainingId) {
      const ss = getTrainingDataSpreadsheet(trainingId);
      if (!ss) return ok([]);
      const sheet = ss.getSheetByName('TrainingSessions');
      if (!sheet) return ok([]);
      return ok(sheetToJson(sheet));
    }

    // Iterate across all trainings if trainingId omitted
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return ok([]);
    const trainings = sheetToJson(tSheet);
    let allSessions = [];
    trainings.forEach(t => {
      if (t.ID) {
        const ss = getTrainingDataSpreadsheet(t.ID);
        if (ss) {
          const sheet = ss.getSheetByName('TrainingSessions');
          if (sheet) {
            allSessions = allSessions.concat(sheetToJson(sheet));
          }
        }
      }
    });
    return ok(allSessions);
  } catch (e) {
    Logger.log('getSessions error: ' + e.message);
    return err('Failed to load sessions: ' + e.message);
  }
}

/**
 * Retrieve details for a single Session by SessionID, enriched with parent Training info.
 * 
 * @param {string} sessionId - The session ID to fetch
 * @returns {string} JSON response with session object and embedded training details
 */
function getSession(sessionId) {
  try {
    if (!sessionId) return err('Session ID is required.');

    const found = findTrainingBySessionId(sessionId);
    if (!found || !found.session) return err('Session not found.');

    const session = found.session;

    // Ensure AttendanceURL and QRCodeURL are present
    if (!session.AttendanceURL) {
      session.AttendanceURL = generateAttendanceURL(session.SessionID);
    }
    if (!session.QRCodeURL) {
      session.QRCodeURL = generateQRCode(session.AttendanceURL);
    }

    // Attach parent Training details
    if (found.training) {
      session.TrainingTitle = found.training.Name || found.training.TrainingTitle || '';
      session.Trainer       = found.training.Trainer || '';
      session.Venue         = found.training.Venue || '';
      session.Category      = found.training.Category || '';
    }

    return ok(session);
  } catch (e) {
    Logger.log('getSession error: ' + e.message);
    return err('Failed to get session: ' + e.message);
  }
}

/**
 * Update the QR status of a session (e.g. Active, Expired, Inactive)
 * 
 * @param {string} sessionId - Session ID
 * @param {string} status - New status (Active, Expired, Inactive)
 * @returns {string} JSON response
 */
function updateSessionQRStatus(sessionId, status) {
  try {
    if (!sessionId) return err('Session ID is required.');
    const found = findTrainingBySessionId(sessionId);
    if (!found || !found.sessionSheet) return err('Session not found.');

    const sheet = found.sessionSheet;
    const row = findRowById(sheet, sessionId);
    if (row === -1) return err('Session not found.');

    sheet.getRange(row, 9).setValue(status); // Column 9 = QRStatus

    if (found.training && found.training.ID) {
      invalidateTrainingCaches(found.training.ID);
    }

    return ok({ message: `Session status updated to: ${status}` });
  } catch (e) {
    return err('Failed to update session status: ' + e.message);
  }
}
