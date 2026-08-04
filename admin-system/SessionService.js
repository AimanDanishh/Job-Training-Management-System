/**
 * SessionService.gs — Training Sessions management service
 */

/**
 * Generate sequential Session ID (e.g., SES0001, SES0002, ...)
 */
function generateSessionId() {
  const sheet = getSheet(SHEET_NAMES.trainingSessions);
  if (!sheet) return 'SES' + Date.now();

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 'SES0001';

  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    if (id.startsWith('SES')) {
      const num = parseInt(id.replace('SES', ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
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

    const sheet = getSheet(SHEET_NAMES.trainingSessions);
    if (!sheet) return err('Could not open TrainingSessions sheet.');

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

    sheet.appendRow([
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
    ]);

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
    const sheet = getSheet(SHEET_NAMES.trainingSessions);
    if (!sheet) return ok([]);

    const rows = sheetToJson(sheet);
    if (trainingId) {
      const filtered = rows.filter(r => String(r.TrainingID).trim() === String(trainingId).trim());
      return ok(filtered);
    }
    return ok(rows);
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

    const sheet = getSheet(SHEET_NAMES.trainingSessions);
    if (!sheet) return err('TrainingSessions sheet not found.');

    const rows = sheetToJson(sheet);
    const session = rows.find(r => String(r.SessionID || '').trim() === String(sessionId).trim());

    if (!session) return err('Session not found.');

    // Ensure AttendanceURL and QRCodeURL are present
    if (!session.AttendanceURL) {
      session.AttendanceURL = generateAttendanceURL(session.SessionID);
    }
    if (!session.QRCodeURL) {
      session.QRCodeURL = generateQRCode(session.AttendanceURL);
    }

    // Attach parent Training details
    try {
      const tSheet = getSheet(SHEET_NAMES.trainings);
      if (tSheet) {
        const tRows = sheetToJson(tSheet);
        const parentTraining = tRows.find(t => String(t.ID || t.TrainingID || '').trim() === String(session.TrainingID).trim());
        if (parentTraining) {
          session.TrainingTitle = parentTraining.Name || parentTraining.TrainingTitle || '';
          session.Trainer       = parentTraining.Trainer || '';
          session.Venue         = parentTraining.Venue || '';
          session.Category      = parentTraining.Category || '';
        }
      }
    } catch (e) {
      Logger.log('Non-fatal parent training lookup error: ' + e.message);
    }

    return ok(session);
  } catch (e) {
    Logger.log('getSession error: ' + e.message);
    return err('Failed to get session: ' + e.message);
  }
}

/**
 * Update an existing session.
 * 
 * @param {string} sessionId - The session ID to update
 * @param {Object} data - Updated session fields
 * @returns {string} JSON response
 */
function updateSession(sessionId, data) {
  try {
    if (!sessionId) return err('Session ID is required.');

    const sheet = getSheet(SHEET_NAMES.trainingSessions);
    const row = findRowById(sheet, sessionId);
    if (row === -1) return err('Session not found.');

    const dataRange = sheet.getRange(row, 1, 1, 10).getValues()[0];

    const updatedSession = [
      sessionId,
      data.TrainingID  || dataRange[1],
      data.SessionName || dataRange[2],
      data.SessionDate || dataRange[3],
      data.StartTime   || dataRange[4],
      data.EndTime     || dataRange[5],
      dataRange[6]     || generateAttendanceURL(sessionId),
      dataRange[7]     || generateQRCode(dataRange[6] || generateAttendanceURL(sessionId)),
      data.QRStatus    || dataRange[8] || 'Active',
      dataRange[9]     || now()
    ];

    sheet.getRange(row, 1, 1, 10).setValues([updatedSession]);
    return ok({ message: 'Session updated successfully.' });
  } catch (e) {
    return err('Failed to update session: ' + e.message);
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
    const sheet = getSheet(SHEET_NAMES.trainingSessions);
    const row = findRowById(sheet, sessionId);
    if (row === -1) return err('Session not found.');

    sheet.getRange(row, 9).setValue(status); // Column 9 = QRStatus
    return ok({ message: `Session status updated to: ${status}` });
  } catch (e) {
    return err('Failed to update session status: ' + e.message);
  }
}

/**
 * Delete a session by SessionID.
 * 
 * @param {string} sessionId - Session ID to delete
 * @returns {string} JSON response
 */
function deleteSession(sessionId) {
  try {
    if (!sessionId) return err('Session ID is required.');
    const sheet = getSheet(SHEET_NAMES.trainingSessions);
    const row = findRowById(sheet, sessionId);
    if (row === -1) return err('Session not found.');

    sheet.deleteRow(row);
    return ok({ message: 'Session deleted successfully.' });
  } catch (e) {
    return err('Failed to delete session: ' + e.message);
  }
}
