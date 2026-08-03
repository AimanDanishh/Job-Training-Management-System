/**
 * TrainingService.gs — Training programme lifecycle management service
 */

/**
 * Create a new Training programme and optionally generate initial sessions.
 * 
 * @param {Object} data - Training programme details
 * @returns {string} JSON response with created training ID and initial sessions info
 */
function createTraining(data) {
  try {
    // Invoke underlying addTraining logic from Training.gs
    const resultJson = addTraining(data);
    const result = typeof resultJson === 'string' ? JSON.parse(resultJson) : resultJson;

    if (!result || !result.success) return resultJson;

    const trainingId = result.data.id;
    const duration = parseInt(data.Duration || 1, 10);

    // Automatically create default session(s) based on training duration / schedule
    const createdSessions = [];

    if (data.sessions && Array.isArray(data.sessions) && data.sessions.length > 0) {
      // Custom session definitions supplied
      data.sessions.forEach(s => {
        const sRes = createSession({
          TrainingID: trainingId,
          SessionName: s.name || s.SessionName || 'Session 1',
          SessionDate: s.date || s.SessionDate || data.StartDate,
          StartTime:   s.startTime || s.StartTime || '09:00',
          EndTime:     s.endTime || s.EndTime || '17:00',
          QRStatus:    'Active'
        });
        const sObj = typeof sRes === 'string' ? JSON.parse(sRes) : sRes;
        if (sObj && sObj.success) createdSessions.push(sObj.data);
      });
    } else {
      // Auto-generate sessions based on Duration
      if (duration === 1) {
        // One-day training -> 1 session (Full Day)
        const sRes = createSession({
          TrainingID: trainingId,
          SessionName: 'Full Day',
          SessionDate: data.StartDate || formatDate(new Date()),
          StartTime:   '09:00',
          EndTime:     '17:00',
          QRStatus:    'Active'
        });
        const sObj = typeof sRes === 'string' ? JSON.parse(sRes) : sRes;
        if (sObj && sObj.success) createdSessions.push(sObj.data);
      } else {
        // Multi-day training -> 1 session per day
        const startDate = new Date(data.StartDate || new Date());
        for (let i = 1; i <= duration; i++) {
          const sDate = new Date(startDate);
          sDate.setDate(startDate.getDate() + (i - 1));
          const dateStr = Utilities.formatDate(sDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

          const sRes = createSession({
            TrainingID: trainingId,
            SessionName: `Day ${i}`,
            SessionDate: dateStr,
            StartTime:   '09:00',
            EndTime:     '17:00',
            QRStatus:    'Active'
          });
          const sObj = typeof sRes === 'string' ? JSON.parse(sRes) : sRes;
          if (sObj && sObj.success) createdSessions.push(sObj.data);
        }
      }
    }

    result.data.sessions = createdSessions;
    return ok(result.data);
  } catch (e) {
    Logger.log('createTraining error: ' + e.message);
    return err('Failed to create training: ' + e.message);
  }
}
