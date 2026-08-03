/**
 * Report.gs — Report generation (returns JSON data for frontend rendering)
 */

// ─── Full Programme Report ──────────────────────────────────────────────────────
function generateReport(trainingId) {
  try {
    // Training details
    const tResult  = JSON.parse(getTrainingById(trainingId));
    if (!tResult.success) return err('Training not found.');
    const training = tResult.data;

    // Attendance summary
    const attSummary = JSON.parse(getAttendanceSummary(trainingId)).data || {};

    // Evaluation summary
    const evalSummary = JSON.parse(getEvaluationSummary(trainingId)).data || {};

    // Full attendance data
    const attData = JSON.parse(getAttendance(trainingId)).data || [];

    // Evaluation data
    const evalData  = JSON.parse(getTrainingEvaluations(trainingId)).data || [];
    const postData  = JSON.parse(getPostEvaluations(trainingId)).data    || [];

    return ok({
      training:      training,
      attendance:    { summary: attSummary, days: attData },
      evaluation:    { summary: evalSummary, responses: evalData },
      postEvaluation: postData,
      generatedAt:   now()
    });
  } catch (e) {
    return err('Failed to generate report: ' + e.message);
  }
}

// ─── Dashboard Overview Report ──────────────────────────────────────────────────
function getDashboardReport() {
  try {
    const trainingSummary = JSON.parse(getTrainingSummary()).data || {};

    // Monthly breakdown (last 6 months)
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const tRows  = sheetToJson(tSheet);
    const monthly = buildMonthlyData(tRows);

    // Dept participation from attendance sheet
    const aSheet = getSheet(SHEET_NAMES.attendance);
    const aRows  = sheetToJson(aSheet);
    const deptData = buildDeptData(aRows);

    return ok({
      summary: trainingSummary,
      monthly: monthly,
      departments: deptData
    });
  } catch (e) {
    return err('Dashboard report error: ' + e.message);
  }
}

// ─── Export Attendance to Sheet (returns a link to a new sheet) ─────────────────
function exportAttendanceSheet(trainingId) {
  try {
    const tResult = JSON.parse(getTrainingById(trainingId));
    if (!tResult.success) return err('Training not found.');
    const training = tResult.data;

    const attData = JSON.parse(getAttendance(trainingId)).data || [];

    const ss    = getSpreadsheet();
    const name  = 'Attendance_' + training.Code + '_' + new Date().getTime();
    const sheet = ss.insertSheet(name);

    // Header
    const headers = ['Employee ID', 'Employee Name', 'Department',
                     'Day', 'Date', 'Check In', 'Check Out', 'Hours', 'Status', 'Remarks'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');

    // Rows
    attData.forEach(day => {
      (day.records || []).forEach(r => {
        sheet.appendRow([
          r.EmployeeID, r.EmployeeName, r.Department,
          r.Day, r.Date, r.CheckIn, r.CheckOut, r.Hours, r.Status, r.Remarks
        ]);
      });
    });

    sheet.autoResizeColumns(1, headers.length);
    return ok({ sheetName: name, message: 'Export sheet created: ' + name });
  } catch (e) {
    return err('Export failed: ' + e.message);
  }
}

// ─── Helpers ────────────────────────────────────────────────────────────────────
function buildMonthlyData(rows) {
  const counts = {};
  rows.forEach(r => {
    if (!r.StartDate) return;
    const d = new Date(r.StartDate);
    if (isNaN(d)) return;
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yyyy');
    counts[key] = (counts[key] || 0) + 1;
  });
  return Object.entries(counts).map(([month, count]) => ({ month, count }));
}

function buildDeptData(rows) {
  const counts = {};
  rows.forEach(r => {
    if (!r.Department) return;
    counts[r.Department] = (counts[r.Department] || 0) + 1;
  });
  return Object.entries(counts)
    .map(([dept, count]) => ({ dept, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 8);
}