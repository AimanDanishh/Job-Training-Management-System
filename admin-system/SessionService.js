/**
 * SessionService.gs - Training Sessions management service
 */

/**
 * Record a decommissioned, deleted, or deactivated session ID in Script Properties
 * so it can NEVER be reused by future session generators.
 */
function recordDecommissionedSessionId(sessionId) {
  if (!sessionId) return;
  try {
    const cleanId = String(sessionId).trim().toUpperCase();
    const raw = getConfigProperty('DECOMMISSIONED_SESSION_IDS', '[]');
    let list = [];
    try { list = JSON.parse(raw); } catch(e) { list = []; }
    if (!Array.isArray(list)) list = [];
    if (!list.includes(cleanId)) {
      list.push(cleanId);
      setConfigProperty('DECOMMISSIONED_SESSION_IDS', JSON.stringify(list));
    }
  } catch(e) {
    Logger.log('recordDecommissionedSessionId warning: ' + e.message);
  }
}

/**
 * Scan all databases to collect all session IDs that have EVER existed:
 * - Central Main Database (all session tabs & attendance tabs)
 * - All Per-Training Spreadsheets (Sessions & Attendance tabs)
 * - Persistent decommissioned / deleted list
 */
function getAllKnownSessionIds() {
  const allIds = new Set();
  const sessionNumPattern = /SES-?(\d+)/i;

  const extractIdsFromSheet = (sheet) => {
    if (!sheet) return;
    try {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol < 1) return;
      
      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const headers = values[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
      let targetCols = [];
      headers.forEach((h, idx) => {
        if (['sessionid', 'sessioncode', 'id', 'session', 'sessid'].includes(h)) {
          targetCols.push(idx);
        }
      });
      if (targetCols.length === 0) {
        targetCols = Array.from({ length: lastCol }, (_, i) => i);
      }

      for (let r = 1; r < values.length; r++) {
        for (const c of targetCols) {
          const cellStr = String(values[r][c] || '').trim();
          if (cellStr && (sessionNumPattern.test(cellStr) || cellStr.toUpperCase().startsWith('SES'))) {
            allIds.add(cellStr.toUpperCase());
          }
        }
      }
    } catch(e) {}
  };

  // 1. Scan Central Main Spreadsheet
  try {
    const mainSs = getSpreadsheet();
    if (mainSs) {
      const allSheets = mainSs.getSheets();
      allSheets.forEach(s => {
        const name = s.getName().toLowerCase();
        if (name.includes('session') || name.includes('attendance')) {
          extractIdsFromSheet(s);
        }
      });
    }
  } catch(e) {}

  // 2. Scan All Per-Training Spreadsheets
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet && tSheet.getLastRow() >= 2) {
      const trainings = sheetToJson(tSheet);
      for (const t of trainings) {
        if (!t.ID && !t.Code) continue;
        try {
          const ss = getTrainingDataSpreadsheet(t);
          if (ss) {
            const allSheets = ss.getSheets();
            allSheets.forEach(s => {
              const name = s.getName().toLowerCase();
              if (name.includes('session') || name.includes('attendance')) {
                extractIdsFromSheet(s);
              }
            });
          }
        } catch(ssErr) {}
      }
    }
  } catch(e) {}

  // 3. Scan Decommissioned / Deleted Session IDs stored in Script Properties
  try {
    const rawDecom = getConfigProperty('DECOMMISSIONED_SESSION_IDS', '');
    if (rawDecom) {
      const parsed = JSON.parse(rawDecom);
      if (Array.isArray(parsed)) {
        parsed.forEach(id => allIds.add(String(id).trim().toUpperCase()));
      }
    }
  } catch(e) {}

  return allIds;
}

/**
 * Generate guaranteed-unique sequential Session ID (e.g., SES0001, SES0002, ...)
 * 
 * STRICT GUARANTEE: The generated ID will NEVER match:
 * - Any active session in central or per-training sheets
 * - Any deactivated or inactive session
 * - Any historically deleted session
 * - Any session ID ever recorded in attendance logs
 * 
 * Uses a monotonically increasing persistent sequence counter (LAST_SESSION_SEQ).
 */
function generateSessionId() {
  const allIds = getAllKnownSessionIds();
  let maxNum = 0;

  // 1. Find the highest numerical ID across all existing, deactivated, and deleted sessions
  allIds.forEach(id => {
    const m = id.match(/SES-?(\d+)/i);
    if (m && m[1]) {
      const n = parseInt(m[1], 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  });

  // 2. Check persistent monotonic sequence counter in Script Properties
  try {
    const storedSeq = parseInt(getConfigProperty('LAST_SESSION_SEQ', '0'), 10);
    if (!isNaN(storedSeq) && storedSeq > maxNum) {
      maxNum = storedSeq;
    }
  } catch(e) {}

  // 3. Increment and ensure the candidate ID has NEVER been used anywhere
  let nextNum = maxNum + 1;
  let candidateId = 'SES' + String(nextNum).padStart(4, '0');
  while (allIds.has(candidateId) || allIds.has('SES-' + String(nextNum).padStart(4, '0'))) {
    nextNum++;
    candidateId = 'SES' + String(nextNum).padStart(4, '0');
  }

  // 4. Update persistent monotonic sequence counter so it NEVER drops even if rows are deleted
  try {
    setConfigProperty('LAST_SESSION_SEQ', String(nextNum));
  } catch(e) {}

  Logger.log(`[SESSION ID GENERATOR] Generated guaranteed-unique Session ID: ${candidateId} (maxNum: ${maxNum})`);
  return candidateId;
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

    const valRes = typeof validateSessionData === 'function' ? validateSessionData(data) : { valid: true };
    if (!valRes.valid) {
      return err(valRes.message);
    }

    // Enforce Approval Check: QR session creation is only allowed for Approved trainings
    let trainingObj = null;
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet) {
      const trainings = sheetToJson(tSheet);
      trainingObj = trainings.find(t => String(t.ID || '').trim() === String(data.TrainingID).trim() || String(t.Code || '').trim() === String(data.TrainingID).trim());
      if (trainingObj) {
        const appStatus = String(trainingObj.ApprovalStatus || '').trim().toLowerCase();
        const isApproved = !appStatus || appStatus === 'approved' || appStatus === 'auto-approved' || appStatus === 'completed' || appStatus === 'in progress' || appStatus === 'active' || appStatus === 'on going';
        if (!isApproved) {
          return err(`QR session creation is only allowed for APPROVED training requisitions. Current approval status: '${trainingObj.ApprovalStatus || 'Pending Approval'}'.`);
        }
      }
    }

    let defaultSessionDate = data.SessionDate;
    if (!defaultSessionDate && trainingObj && trainingObj.StartDate) {
      defaultSessionDate = formatDate(trainingObj.StartDate);
    }
    if (!defaultSessionDate) {
      defaultSessionDate = formatDate(new Date());
    }

    const sessionId = generateSessionId();
    const attendanceUrl = generateAttendanceURL(sessionId);
    const qrCodeUrl = generateQRCode(attendanceUrl);
    const timeNow = now();

    const newSession = {
      SessionID:     sessionId,
      TrainingID:    data.TrainingID,
      SessionName:   data.SessionName,
      SessionDate:   defaultSessionDate,
      StartTime:     data.StartTime || '09:00',
      EndTime:       data.EndTime || '16:00',
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

    if (data.StartTime || data.EndTime) {
      const valRes = typeof validateSessionData === 'function' ? validateSessionData({
        TrainingID: trainingId || 'TRN',
        SessionName: data.SessionName || (found.session && found.session.SessionName) || 'Session',
        StartTime: data.StartTime || (found.session && found.session.StartTime) || '09:00',
        EndTime: data.EndTime || (found.session && found.session.EndTime) || '16:00'
      }) : { valid: true };
      if (!valRes.valid) {
        return err(valRes.message);
      }
    }

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
        const cleanStatus = String(data.QRStatus).trim().toLowerCase();
        const finalStatus = (cleanStatus === 'deactivate' || cleanStatus === 'deactivated' || cleanStatus === 'inactive' || cleanStatus === 'expired' || cleanStatus === 'disabled') ? 'Deactivate' : 'Active';
        setColValFlexible(targetSheet, row, headers, [/^qrstatus$/i, /^qr status$/i, /^status$/i, /^state$/i], 'QRStatus', finalStatus);
        if (finalStatus === 'Deactivate') {
          try {
            if (typeof markUnscannedParticipantsAbsent === 'function') {
              markUnscannedParticipantsAbsent(cleanSessionId, (training && training.ID) || (found.session && found.session.TrainingID));
            }
          } catch(mErr) {
            Logger.log('markUnscannedParticipantsAbsent in updateSession error: ' + mErr.message);
          }
        }
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
 * Soft-delete a Training Session by marking its QRStatus as 'Deactivate' in both spreadsheets.
 * Preserves all historical attendance records intact.
 * 
 * @param {string} sessionId - The session ID to deactivate
 * @returns {string} JSON response
 */
function deleteSession(sessionId) {
  try {
    if (!sessionId) return err('Session ID is required.');
    const cleanSessionId = String(sessionId).trim();
    const lowerSessionId = cleanSessionId.toLowerCase();

    // 0. Permanently blacklist this session ID from ever being regenerated
    recordDecommissionedSessionId(cleanSessionId);

    // 1. Direct deactivation in central Main Database Sessions tabs
    try {
      const mainSs = getSpreadsheet();
      if (mainSs) {
        const sessionSheetNames = ['TrainingSessions', 'Sessions', 'Training Sessions', 'Session'];
        sessionSheetNames.forEach(sheetName => {
          const sh = mainSs.getSheetByName(sheetName);
          if (sh && sh.getLastRow() >= 2) {
            const data = sh.getDataRange().getValues();
            const headers = data[0].map(h => String(h || '').trim().toLowerCase());
            const sidCol = headers.findIndex(h => h === 'sessionid' || h === 'session_id' || h === 'session id' || h === 'id' || h === 'sessioncode');
            const qrCol = headers.findIndex(h => h === 'qrstatus' || h === 'qr status' || h === 'status' || h === 'state');

            if (sidCol !== -1) {
              for (let r = 1; r < data.length; r++) {
                if (String(data[r][sidCol] || '').trim().toLowerCase() === lowerSessionId) {
                  if (qrCol !== -1) {
                    sh.getRange(r + 1, qrCol + 1).setValue('Deactivate');
                  }
                }
              }
            }
          }
        });
      }
    } catch(cErr) {
      Logger.log('Central sheet deleteSession error: ' + cErr.message);
    }

    // 2. Direct deactivation across all per-training spreadsheets
    try {
      const tSheet = getSheet(SHEET_NAMES.trainings);
      if (tSheet && tSheet.getLastRow() >= 2) {
        const trainings = sheetToJson(tSheet);
        for (const t of trainings) {
          try {
            const perSs = getTrainingDataSpreadsheet(t);
            if (perSs) {
              const sessionSheetNames = ['TrainingSessions', 'Sessions', 'Training Sessions', 'Session'];
              sessionSheetNames.forEach(sheetName => {
                const sh = perSs.getSheetByName(sheetName);
                if (sh && sh.getLastRow() >= 2) {
                  const data = sh.getDataRange().getValues();
                  const headers = data[0].map(h => String(h || '').trim().toLowerCase());
                  const sidCol = headers.findIndex(h => h === 'sessionid' || h === 'session_id' || h === 'session id' || h === 'id' || h === 'sessioncode');
                  const qrCol = headers.findIndex(h => h === 'qrstatus' || h === 'qr status' || h === 'status' || h === 'state');

                  if (sidCol !== -1) {
                    for (let r = 1; r < data.length; r++) {
                      if (String(data[r][sidCol] || '').trim().toLowerCase() === lowerSessionId) {
                        if (qrCol !== -1) {
                          sh.getRange(r + 1, qrCol + 1).setValue('Deactivate');
                        }
                      }
                    }
                  }
                }
              });
            }
          } catch(perTErr) {}
        }
      }
    } catch(tErr) {
      Logger.log('Per-training deleteSession error: ' + tErr.message);
    }

    // 3. Also run updateSession to trigger cache invalidations and hooks
    try {
      updateSession(cleanSessionId, { QRStatus: 'Deactivate' });
    } catch(uErr) {}

    // 4. Mark all enrolled participants who haven't scanned as Absent
    try {
      if (typeof markUnscannedParticipantsAbsent === 'function') {
        markUnscannedParticipantsAbsent(cleanSessionId);
      }
    } catch(mErr) {
      Logger.log('markUnscannedParticipantsAbsent in deleteSession error: ' + mErr.message);
    }

    invalidateTrainingCaches();

    return ok({ message: `QR session ${cleanSessionId} deleted successfully (marked Deactivate).`, sessionId: cleanSessionId });
  } catch (e) {
    Logger.log('deleteSession error: ' + e.message);
    return err('Failed to delete session: ' + e.message);
  }
}

/**
 * Generates a brand-new QR Code with a fresh, guaranteed-unique Session ID for an existing session.
 * Permanently retires and deactivates the old Session ID and QR code,
 * ensuring the old QR code can NEVER be used for attendance.
 * 
 * @param {string} oldSessionId - The session ID to renew/replace
 * @returns {string} JSON response with the updated session containing the new QR code
 */
function renewSessionQRCode(oldSessionId) {
  try {
    if (!oldSessionId) return err('Session ID is required.');
    const cleanOldId = String(oldSessionId).trim();
    const found = findTrainingBySessionId(cleanOldId);
    if (!found || !found.session) return err(`Session ${cleanOldId} not found.`);

    const oldSession = found.session;
    const trainingId = oldSession.TrainingID;

    // 1. Permanently blacklist old session ID
    recordDecommissionedSessionId(cleanOldId);

    // 2. Mark old session row as deactivated
    deleteSession(cleanOldId);

    // 3. Generate a brand new, strictly unique Session ID
    const newSessionId = generateSessionId();
    const newAttendanceUrl = generateAttendanceURL(newSessionId);
    const newQRCodeUrl = generateQRCode(newAttendanceUrl);
    const timeNow = now();

    // 4. Create new session with the same session metadata but active with fresh QR code
    const newSession = {
      SessionID:     newSessionId,
      TrainingID:    trainingId,
      SessionName:   oldSession.SessionName || 'Session Check-In',
      SessionDate:   oldSession.SessionDate || '',
      StartTime:     oldSession.StartTime || '09:00',
      EndTime:       oldSession.EndTime || '16:00',
      AttendanceURL: newAttendanceUrl,
      QRCodeURL:     newQRCodeUrl,
      QRStatus:      'Active',
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

    // Append to per-training spreadsheet
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (ss) {
      let sheet = ss.getSheetByName('TrainingSessions') || ss.getSheetByName('Sessions') || ss.getSheetByName('Training Sessions');
      if (sheet) sheet.appendRow(sessionRow);
    }

    // Append to central Main Database
    const mainSs = getSpreadsheet();
    if (mainSs) {
      let cSheet = mainSs.getSheetByName('TrainingSessions') || mainSs.getSheetByName('Sessions') || mainSs.getSheetByName('Training Sessions');
      if (cSheet) cSheet.appendRow(sessionRow);
    }

    invalidateTrainingCaches(trainingId);

    Logger.log(`[QR RENEWAL] Replaced session ${cleanOldId} with fresh session ${newSessionId} and new QR code.`);
    return ok({
      message: `New QR Code generated successfully! Session ID: ${newSessionId}. Old session ${cleanOldId} is permanently deactivated.`,
      oldSessionId: cleanOldId,
      newSession: newSession
    });
  } catch (e) {
    Logger.log('renewSessionQRCode error: ' + e.message);
    return err('Failed to generate new QR code: ' + e.message);
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
 * Checks existence and that status is Active.
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

    const status = String(found.session.QRStatus || 'Active').trim().toLowerCase();
    if (status === 'deactivate' || status === 'deactivated' || status === 'inactive' || status === 'expired' || status === 'disabled') {
      return err('This QR attendance session is deactivated. Attendance cannot be recorded.');
    }

    return ok(found.session);
  } catch (e) {
    Logger.log('validateQRSessionForAttendance error: ' + e.message);
    return err('Failed to validate session: ' + e.message);
  }
}
