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
      const training = trainings.find(t => String(t.ID || '').trim() === String(data.TrainingID).trim() || String(t.Code || '').trim() === String(data.TrainingID).trim());
      if (training) {
        const appStatus = String(training.ApprovalStatus || '').trim().toLowerCase();
        const isApproved = !appStatus || appStatus === 'approved' || appStatus === 'auto-approved' || appStatus === 'completed' || appStatus === 'in progress' || appStatus === 'active' || appStatus === 'on going';
        if (!isApproved) {
          return err(`QR session creation is only allowed for APPROVED training requisitions. Current approval status: '${training.ApprovalStatus || 'Pending Approval'}'.`);
        }
      }
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

    // 1. Try to append to per-training spreadsheet
    let ss = getTrainingDataSpreadsheet(data.TrainingID);
    if (!ss && typeof createTrainingWorkspace === 'function') {
      try {
        const wsRes = createTrainingWorkspace(data.TrainingID);
        if (wsRes && wsRes.spreadsheetId) {
          ss = SpreadsheetApp.openById(wsRes.spreadsheetId);
        }
      } catch(wsErr) {}
    }

    if (ss) {
      let sheet = ss.getSheetByName('TrainingSessions') || ss.getSheetByName('Sessions') || ss.getSheetByName('Training Sessions');
      if (!sheet) {
        sheet = ss.insertSheet('TrainingSessions');
        sheet.appendRow(['SessionID', 'TrainingID', 'SessionName', 'SessionDate', 'StartTime', 'EndTime', 'AttendanceURL', 'QRCodeURL', 'QRStatus', 'CreatedDate']);
        sheet.getRange('A1:J1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
      }
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
    }

    // 2. Sync session record to central TrainingSessions tab in Main Database
    try {
      const mainSs = getSpreadsheet();
      if (mainSs) {
        let centralSheet = mainSs.getSheetByName('TrainingSessions') || mainSs.getSheetByName('Sessions') || mainSs.getSheetByName('Training Sessions');
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

    try {
      SpreadsheetApp.flush();
    } catch(fErr) {}

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
    const sessionsMap = new Map();

    // 1. Gather all training aliases if trainingId is provided
    let targetTraining = null;
    const targetAliases = new Set();

    const tSheet = getSheet(SHEET_NAMES.trainings);
    const trainings = tSheet ? sheetToJson(tSheet) : [];

    if (trainingId) {
      const cleanInput = String(trainingId).trim().toLowerCase();
      targetAliases.add(cleanInput);
      targetTraining = trainings.find(t => {
        const id = String(t.ID || '').trim().toLowerCase();
        const code = String(t.Code || '').trim().toLowerCase();
        const tId = String(t.TrainingID || '').trim().toLowerCase();
        const name = String(t.Name || '').trim().toLowerCase();
        return id === cleanInput || code === cleanInput || tId === cleanInput || name === cleanInput;
      });

      if (targetTraining) {
        if (targetTraining.ID) targetAliases.add(String(targetTraining.ID).trim().toLowerCase());
        if (targetTraining.Code) targetAliases.add(String(targetTraining.Code).trim().toLowerCase());
        if (targetTraining.TrainingID) targetAliases.add(String(targetTraining.TrainingID).trim().toLowerCase());
      }
    }

    const matchesFilter = (sTId) => {
      if (!trainingId || targetAliases.size === 0) return true;
      const clean = String(sTId || '').trim().toLowerCase();
      return targetAliases.has(clean);
    };

    // 2. Check central TrainingSessions tab in Main Database
    try {
      const mainSs = getSpreadsheet();
      if (mainSs) {
        const cSheet = mainSs.getSheetByName('TrainingSessions') || mainSs.getSheetByName('Sessions') || mainSs.getSheetByName('Training Sessions');
        if (cSheet && cSheet.getLastRow() > 1) {
          const allC = sheetToJson(cSheet);
          allC.forEach(s => {
            const sId = String(s.SessionID || s.ID || '').trim().toLowerCase();
            const sTId = String(s.TrainingID || '').trim();
            if (matchesFilter(sTId)) {
              if (sId) sessionsMap.set(sId, s);
            }
          });
        }
      }
    } catch(eC) {
      Logger.log('getSessions central sheet load error: ' + eC.message);
    }

    // 3. Check per-training sheets
    if (trainingId) {
      const targetInput = targetTraining || trainingId;
      const ss = getTrainingDataSpreadsheet(targetInput);
      if (ss) {
        const sheet = ss.getSheetByName('TrainingSessions') || ss.getSheetByName('Sessions') || ss.getSheetByName('Training Sessions') || ss.getSheetByName('Session');
        if (sheet && sheet.getLastRow() > 1) {
          const pSessions = sheetToJson(sheet);
          pSessions.forEach(ps => {
            const sId = String(ps.SessionID || ps.ID || '').trim().toLowerCase();
            if (sId) {
              const existing = sessionsMap.get(sId) || {};
              sessionsMap.set(sId, Object.assign({}, ps, existing));
            }
          });
        }
      }
    } else {
      trainings.forEach(t => {
        if (t.ID || t.Code) {
          const ss = getTrainingDataSpreadsheet(t);
          if (ss) {
            const sheet = ss.getSheetByName('TrainingSessions') || ss.getSheetByName('Sessions') || ss.getSheetByName('Training Sessions') || ss.getSheetByName('Session');
            if (sheet && sheet.getLastRow() > 1) {
              const pSessions = sheetToJson(sheet);
              pSessions.forEach(ps => {
                const sId = String(ps.SessionID || ps.ID || '').trim().toLowerCase();
                if (sId) {
                  const existing = sessionsMap.get(sId) || {};
                  sessionsMap.set(sId, Object.assign({}, ps, existing));
                }
              });
            }
          }
        }
      });
    }

    // 4. Attach parent training details if missing
    const results = Array.from(sessionsMap.values()).map(s => {
      if (!s.TrainingTitle) {
        const pT = trainings.find(t => {
          const sTId = String(s.TrainingID || '').trim().toLowerCase();
          return sTId === String(t.ID || '').trim().toLowerCase() ||
                 sTId === String(t.Code || '').trim().toLowerCase() ||
                 sTId === String(t.TrainingID || '').trim().toLowerCase();
        });
        if (pT) {
          s.TrainingTitle = pT.Name || pT.TrainingTitle || '';
          if (!s.Trainer) s.Trainer = pT.Trainer || '';
          if (!s.Venue)   s.Venue   = pT.Venue || '';
        }
      }
      return s;
    });

    return ok(results);
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
  return updateSession(sessionId, { QRStatus: status });
}

/**
 * Update an existing Training Session record in both per-training and central spreadsheets.
 * Preserves SessionID, TrainingID, AttendanceURL, QRCodeURL, and CreatedDate.
 * 
 * @param {string} sessionId - The Session ID to update (e.g. SES0001)
 * @param {Object} data - { SessionName, SessionDate, StartTime, EndTime, QRStatus }
 * @returns {string} JSON response with updated session object
 */
function updateSession(sessionId, data) {
  try {
    if (!sessionId) return err('Session ID is required.');
    if (!data) return err('Session update data is required.');

    const cleanSessionId = String(sessionId).trim();
    const found = findTrainingBySessionId(cleanSessionId);
    if (!found) return err(`Session not found for ID: ${cleanSessionId}`);

    const training = found.training || {};
    const trainingId = training.ID || (found.session && found.session.TrainingID) || (data && data.TrainingID) || '';

    const setColValFlexible = (targetSheet, rIdx, colHeaders, colPatterns, defaultHeaderName, value) => {
      if (rIdx < 2 || !targetSheet) return;
      for (let pattern of colPatterns) {
        const cIdx = colHeaders.findIndex(h => pattern.test(String(h).trim())) + 1;
        if (cIdx > 0) {
          targetSheet.getRange(rIdx, cIdx).setValue(value);
          return;
        }
      }
      // If column does not exist in header row, add it
      const newColIdx = targetSheet.getLastColumn() + 1;
      targetSheet.getRange(1, newColIdx).setValue(defaultHeaderName);
      targetSheet.getRange(rIdx, newColIdx).setValue(value);
      colHeaders.push(defaultHeaderName);
    };

    const findRowInSheet = (targetSheet, sId) => {
      if (!targetSheet) return -1;
      const lastRow = targetSheet.getLastRow();
      if (lastRow < 2) return -1;
      const lastCol = targetSheet.getLastColumn();
      if (lastCol < 1) return -1;
      const vals = targetSheet.getRange(1, 1, lastRow, lastCol).getValues();
      const cleanSId = String(sId).trim().toLowerCase();
      for (let r = 1; r < vals.length; r++) {
        for (let c = 0; c < vals[r].length; c++) {
          if (String(vals[r][c]).trim().toLowerCase() === cleanSId) {
            return r + 1;
          }
        }
      }
      return -1;
    };

    const applyUpdatesToSheet = (targetSheet) => {
      if (!targetSheet) return;
      let row = findRowInSheet(targetSheet, cleanSessionId);
      const headers = targetSheet.getRange(1, 1, 1, Math.max(1, targetSheet.getLastColumn())).getValues()[0].map(h => String(h).trim());

      if (row < 2) {
        // Append session if missing in this sheet
        const fullSession = (found && found.session) ? found.session : {};
        const merged = Object.assign({}, fullSession, data, { SessionID: cleanSessionId, TrainingID: trainingId });
        const newRow = headers.map(h => {
          const matchKey = Object.keys(merged).find(k => k.toLowerCase() === h.toLowerCase().replace(/[\s_]/g, ''));
          return matchKey ? merged[matchKey] : '';
        });
        targetSheet.appendRow(newRow);
        row = targetSheet.getLastRow();
      }

      if (data.SessionName !== undefined && data.SessionName !== null) {
        setColValFlexible(targetSheet, row, headers, [/^sessionname$/i, /^session name$/i, /^name$/i, /^title$/i], 'SessionName', String(data.SessionName).trim());
      }
      if (data.SessionDate !== undefined && data.SessionDate !== null) {
        setColValFlexible(targetSheet, row, headers, [/^sessiondate$/i, /^session date$/i, /^date$/i], 'SessionDate', String(data.SessionDate).trim());
      }
      if (data.StartTime !== undefined && data.StartTime !== null) {
        setColValFlexible(targetSheet, row, headers, [/^starttime$/i, /^start time$/i, /^time start$/i], 'StartTime', String(data.StartTime).trim());
      }
      if (data.EndTime !== undefined && data.EndTime !== null) {
        setColValFlexible(targetSheet, row, headers, [/^endtime$/i, /^end time$/i, /^time end$/i], 'EndTime', String(data.EndTime).trim());
      }
      if (data.QRStatus !== undefined && data.QRStatus !== null) {
        const validStatuses = ['Active', 'Inactive', 'Expired'];
        const cleanStatus = String(data.QRStatus).trim();
        const finalStatus = validStatuses.find(s => s.toLowerCase() === cleanStatus.toLowerCase()) || cleanStatus;
        setColValFlexible(targetSheet, row, headers, [/^qrstatus$/i, /^qr status$/i, /^status$/i, /^state$/i], 'QRStatus', finalStatus);
      }
    };

    // 1. Update in per-training sheet (if individual Drive sheet exists)
    try {
      const targetTrainingInput = (training && (training.ID || training.Code)) ? training : (data && data.TrainingID) || (found.session && found.session.TrainingID);
      if (targetTrainingInput) {
        const perSs = getTrainingDataSpreadsheet(targetTrainingInput);
        if (perSs) {
          const perSessSheet = perSs.getSheetByName('TrainingSessions') || 
                               perSs.getSheetByName('Sessions') || 
                               perSs.getSheetByName('Training Sessions') || 
                               perSs.getSheetByName('Session');
          if (perSessSheet) {
            applyUpdatesToSheet(perSessSheet);
          }
        }
      }
    } catch(perErr) {
      Logger.log('Per-training sheet session update error: ' + perErr.message);
    }

    // 2. Update in central TrainingSessions tab in Main Database
    try {
      const mainSs = getSpreadsheet();
      if (mainSs) {
        const centralSheet = mainSs.getSheetByName('TrainingSessions') || 
                             mainSs.getSheetByName('Sessions') || 
                             mainSs.getSheetByName('Training Sessions') || 
                             mainSs.getSheetByName('Session');
        if (centralSheet) {
          applyUpdatesToSheet(centralSheet);
        }
      }
    } catch (cErr) {
      Logger.log('Central TrainingSessions update error: ' + cErr.message);
    }

    try {
      SpreadsheetApp.flush();
    } catch(fErr) {}

    // Invalidate server caches for all ID aliases
    const idsToInvalidate = new Set();
    if (trainingId) idsToInvalidate.add(String(trainingId));
    if (data && data.TrainingID) idsToInvalidate.add(String(data.TrainingID));
    if (training && training.ID) idsToInvalidate.add(String(training.ID));
    if (training && training.Code) idsToInvalidate.add(String(training.Code));
    if (found.session && found.session.TrainingID) idsToInvalidate.add(String(found.session.TrainingID));

    idsToInvalidate.forEach(id => {
      invalidateTrainingCaches(id);
    });

    // Fetch and return the updated session
    const updatedRes = getSession(cleanSessionId);
    const updatedObj = typeof updatedRes === 'string' ? JSON.parse(updatedRes) : updatedRes;
    return ok(updatedObj && updatedObj.data ? updatedObj.data : { SessionID: cleanSessionId, message: 'Session updated successfully.' });
  } catch (e) {
    Logger.log('updateSession error: ' + e.message);
    return err('Failed to update session: ' + e.message);
  }
}

/**
 * Alias for updateSession
 */
function updateQRSession(sessionId, data) {
  return updateSession(sessionId, data);
}

/**
 * Soft-delete a Training Session by marking its QRStatus as 'Inactive' in both spreadsheets.
 * Preserves all historical attendance records intact.
 * 
 * @param {string} sessionId - The session ID to deactivate
 * @returns {string} JSON response
 */
function deleteSession(sessionId) {
  try {
    if (!sessionId) return err('Session ID is required.');
    const cleanSessionId = String(sessionId).trim();

    const res = updateSession(cleanSessionId, { QRStatus: 'Inactive' });
    const resObj = typeof res === 'string' ? JSON.parse(res) : res;
    if (resObj && !resObj.success) {
      return err(resObj.error || 'Failed to delete session.');
    }

    return ok({ message: `QR session ${cleanSessionId} deleted successfully (marked Inactive).`, sessionId: cleanSessionId });
  } catch (e) {
    Logger.log('deleteSession error: ' + e.message);
    return err('Failed to delete session: ' + e.message);
  }
}

/**
 * Alias for deleteSession
 */
function deleteQRSession(sessionId) {
  return deleteSession(sessionId);
}

/**
 * Alias for getSessions
 */
function getQRSessions(trainingId) {
  return getSessions(trainingId);
}

/**
 * Server-side validator for a QR session's eligibility for attendance marking.
 * Reads fresh session record directly from the database sheet.
 * 
 * @param {string} sessionId - The session ID to check
 * @returns {string} JSON response
 */
function validateQRSessionForAttendance(sessionId) {
  try {
    if (!sessionId) return err('Session ID is required.');
    const cleanSessionId = String(sessionId).trim();

    const found = findTrainingBySessionId(cleanSessionId);
    if (!found || !found.session) return err('Invalid session ID. Session does not exist.');

    const status = String(found.session.QRStatus || 'Active').trim();
    if (status.toLowerCase() === 'inactive') {
      return err('This QR attendance session is no longer active. Attendance cannot be recorded.');
    }
    if (status.toLowerCase() === 'expired') {
      return err('Attendance registration for this session is closed (Expired).');
    }

    return ok(found.session);
  } catch (e) {
    Logger.log('validateQRSessionForAttendance error: ' + e.message);
    return err('Failed to validate session: ' + e.message);
  }
}
