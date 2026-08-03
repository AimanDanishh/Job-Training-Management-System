/**
 * Attendance.gs — Per-day attendance recording and editing
 */

// ─── Read All for a Training ────────────────────────────────────────────────────
function getAttendance(trainingId) {
  try {
    const sheet = getSheet(SHEET_NAMES.attendance);
    const rows = sheetToJson(sheet);
    const filtered = rows.filter(r => r.TrainingID === trainingId);

    // Group by day
    const days = {};
    filtered.forEach(r => {
      const d = r.Day;
      if (!days[d]) days[d] = { day: d, date: r.Date, records: [] };
      days[d].records.push(r);
    });

    return ok(Object.values(days).sort((a, b) => Number(a.day) - Number(b.day)));
  } catch (e) {
    return err('Failed to load attendance: ' + e.message);
  }
}

// ─── Bulk Save (all records for a day) ─────────────────────────────────────────
function saveAttendance(trainingId, trainingCode, dayNumber, date, records) {
  try {
    const sheet = getSheet(SHEET_NAMES.attendance);
    const allData = sheet.getDataRange().getValues();

    // Remove existing rows for this training + day (rebuild from scratch)
    for (let i = allData.length - 1; i >= 1; i--) {
      if (String(allData[i][1]) === String(trainingId) &&
          String(allData[i][6]) === String(dayNumber)) {
        sheet.deleteRow(i + 1);
      }
    }

    // Append updated records
    records.forEach(r => {
      sheet.appendRow([
        generateId('ATT'),
        trainingId,
        trainingCode,
        r.EmployeeID   || '',
        r.EmployeeName || '',
        r.Department   || '',
        dayNumber,
        date,
        r.CheckIn      || '',
        r.CheckOut     || '',
        r.Hours        || 0,
        r.Status       || 'Present',
        r.Remarks      || '',
        r.EditedBy     || '',
        r.EditedAt     || ''
      ]);
    });

    syncTrainingRequisitionParticipants(trainingId);

    return ok({ message: 'Attendance saved for Day ' + dayNumber });
  } catch (e) {
    return err('Failed to save attendance: ' + e.message);
  }
}

// ─── Update Single Record ───────────────────────────────────────────────────────
function updateAttendanceRecord(id, updatedFields) {
  try {
    const sheet = getSheet(SHEET_NAMES.attendance);
    const row = findRowById(sheet, id);
    if (row === -1) return err('Attendance record not found.');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const current = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

    // Merge: only update provided fields
    const fieldMap = {};
    headers.forEach((h, i) => { fieldMap[h] = i; });

    if (updatedFields.CheckIn  !== undefined) current[fieldMap['CheckIn']]  = updatedFields.CheckIn;
    if (updatedFields.CheckOut !== undefined) current[fieldMap['CheckOut']] = updatedFields.CheckOut;
    if (updatedFields.Hours    !== undefined) current[fieldMap['Hours']]    = updatedFields.Hours;
    if (updatedFields.Status   !== undefined) current[fieldMap['Status']]   = updatedFields.Status;
    if (updatedFields.Remarks  !== undefined) current[fieldMap['Remarks']]  = updatedFields.Remarks;
    current[fieldMap['EditedBy']] = updatedFields.EditedBy || 'Admin';
    current[fieldMap['EditedAt']] = now();

    sheet.getRange(row, 1, 1, headers.length).setValues([current]);
    return ok({ message: 'Record updated.' });
  } catch (e) {
    return err('Failed to update record: ' + e.message);
  }
}

// ─── Attendance Summary for Dashboard ──────────────────────────────────────────
function getAttendanceSummary(trainingId) {
  try {
    const sheet = getSheet(SHEET_NAMES.attendance);
    const rows = sheetToJson(sheet).filter(r => r.TrainingID === trainingId);
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
