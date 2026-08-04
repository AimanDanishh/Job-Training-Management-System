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

    let requisitionFormUrl = '';
    if (training.RequisitionFormFileID) {
      requisitionFormUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(training.RequisitionFormFileID)}/edit`;
    }

    return ok({
      training:           training,
      requisitionFormUrl: requisitionFormUrl,
      attendance:         { summary: attSummary, days: attData },
      evaluation:         { summary: evalSummary, responses: evalData },
      postEvaluation:     postData,
      generatedAt:        now()
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

// ─── Export Full Report to Excel (.xlsx / Sheet) ──────────────────────────────
function exportReportExcel(trainingId) {
  try {
    const reportRes = JSON.parse(generateReport(trainingId));
    if (!reportRes.success) return err(reportRes.message);
    const rep = reportRes.data;
    const t = rep.training;

    const ssName = `${t.Code}_Full_Report_${Date.now()}`;
    const ss = SpreadsheetApp.create(ssName);

    // Sheet 1: Programme Summary
    const sSummary = ss.getActiveSheet();
    sSummary.setName('Programme Overview');
    sSummary.appendRow(['TRAINING PROGRAMME SUMMARY REPORT']);
    sSummary.appendRow(['Generated Date:', rep.generatedAt]);
    sSummary.appendRow([]);
    sSummary.appendRow(['Programme Code', t.Code]);
    sSummary.appendRow(['Programme Name', t.Name]);
    sSummary.appendRow(['Category', t.Category]);
    sSummary.appendRow(['Trainer / Facilitator', t.Trainer]);
    sSummary.appendRow(['Venue', t.Venue || '—']);
    sSummary.appendRow(['Course Fee (RM)', t.CourseFee || '0.00']);
    sSummary.appendRow(['Department', t.Department || 'All Departments']);
    sSummary.appendRow(['Start Date', t.StartDate]);
    sSummary.appendRow(['End Date', t.EndDate || '—']);
    sSummary.appendRow(['Duration (Days)', t.Duration]);
    sSummary.appendRow(['Total Hours', t.TotalHours]);
    sSummary.appendRow(['Status', t.Status]);
    sSummary.appendRow(['Lifecycle Stage', t.Stage]);
    sSummary.appendRow(['Enrolled Participants', t.Participants]);
    sSummary.appendRow(['Objectives', t.Objectives || '—']);
    sSummary.appendRow([]);
    sSummary.appendRow(['KEY PERFORMANCE METRICS']);
    sSummary.appendRow(['Attendance Rate (%)', rep.attendance.summary.pct || 0]);
    sSummary.appendRow(['Present Records Count', rep.attendance.summary.present || 0]);
    sSummary.appendRow(['Absent Records Count', rep.attendance.summary.absent || 0]);
    sSummary.appendRow(['Evaluations Completed', rep.evaluation.summary.evalCompleted || 0]);
    sSummary.appendRow(['Average Evaluation Score', rep.evaluation.summary.avgScore || 'N/A']);
    sSummary.getRange('A1:B1').setFontWeight('bold').setFontSize(14);
    sSummary.getRange('A20:B20').setFontWeight('bold').setFontSize(12);

    // Sheet 2: Daily Attendance
    const sAtt = ss.insertSheet('Attendance Records');
    const attHeaders = ['Day', 'Date', 'Employee ID', 'Employee Name', 'Department', 'Check In', 'Check Out', 'Hours', 'Status', 'Remarks'];
    sAtt.appendRow(attHeaders);
    sAtt.getRange(1, 1, 1, attHeaders.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
    (rep.attendance.days || []).forEach(day => {
      (day.records || []).forEach(r => {
        sAtt.appendRow([r.Day, r.Date, r.EmployeeID, r.EmployeeName, r.Department, r.CheckIn, r.CheckOut, r.Hours, r.Status, r.Remarks]);
      });
    });

    // Sheet 3: Training Evaluations
    const sEval = ss.insertSheet('Training Evaluations');
    const evalHeaders = ['Evaluation ID', 'Employee ID', 'Employee Name', 'Q1 (Obj)', 'Q2 (Rel)', 'Q3 (Mat)', 'Q4 (Trn)', 'Q5 (Eng)', 'Q6 (Dur)', 'Q7 (App)', 'Avg Score', 'Section B1', 'Section B2', 'Section B3', 'Submitted At'];
    sEval.appendRow(evalHeaders);
    sEval.getRange(1, 1, 1, evalHeaders.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
    (rep.evaluation.responses || []).forEach(ev => {
      sEval.appendRow([ev.ID, ev.EmployeeID, ev.EmployeeName, ev.Q1, ev.Q2, ev.Q3, ev.Q4, ev.Q5, ev.Q6, ev.Q7, ev.AvgScore, ev.SectionB1, ev.SectionB2, ev.SectionB3, ev.SubmittedAt]);
    });

    // Sheet 4: 6-Month Post Evaluation
    const sPost = ss.insertSheet('6-Month Post Evaluations');
    const postHeaders = ['Post Eval ID', 'Employee ID', 'Evaluator Name', 'Evaluator ID', 'Competency Before', 'Competency After', 'Improvement', 'Can Apply', 'Further Training', 'Comments', 'Submitted At'];
    sPost.appendRow(postHeaders);
    sPost.getRange(1, 1, 1, postHeaders.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
    (rep.postEvaluation || []).forEach(p => {
      sPost.appendRow([p.ID, p.EmployeeID, p.EvaluatorName, p.EvaluatorID, p.CompetencyBefore, p.CompetencyAfter, p.Improvement, p.CanApply, p.FurtherTraining, p.Comments, p.SubmittedAt]);
    });

    SpreadsheetApp.flush();

    const fileId = ss.getId();
    const sheetUrl = ss.getUrl();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    return ok({
      fileId: fileId,
      fileName: ssName,
      sheetUrl: sheetUrl,
      downloadUrl: downloadUrl,
      message: 'Excel report generated successfully.'
    });
  } catch (e) {
    return err('Failed to export Excel report: ' + e.message);
  }
}

// ─── Export Full Report to PDF (.pdf) ──────────────────────────────────────────
function exportReportPdf(trainingId) {
  try {
    const reportRes = JSON.parse(generateReport(trainingId));
    if (!reportRes.success) return err(reportRes.message);
    const rep = reportRes.data;
    const t = rep.training;

    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          body { font-family: 'Helvetica Neue', Arial, sans-serif; color: #1E293B; margin: 20px; line-height: 1.5; font-size: 12px; }
          .header { border-bottom: 2px solid #2563EB; padding-bottom: 12px; margin-bottom: 20px; }
          .title { font-size: 20px; font-weight: bold; color: #1E293B; margin: 0 0 4px; }
          .subtitle { font-size: 12px; color: #64748B; margin: 0; }
          .section { margin-bottom: 20px; }
          .section-title { font-size: 14px; font-weight: bold; color: #2563EB; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px; margin-bottom: 10px; }
          .grid { display: table; width: 100%; margin-bottom: 10px; }
          .row { display: table-row; }
          .cell { display: table-cell; padding: 4px 8px; font-size: 11px; }
          .cell-label { font-weight: bold; color: #64748B; width: 30%; }
          .kpi-container { width: 100%; margin-bottom: 15px; }
          .kpi-card { display: inline-block; width: 18%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 8px; text-align: center; margin-right: 1.5%; box-sizing: border-box; }
          .kpi-value { font-size: 16px; font-weight: bold; color: #2563EB; }
          .kpi-label { font-size: 9px; color: #64748B; margin-top: 2px; }
          table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 10px; }
          th { background: #2563EB; color: white; padding: 6px; text-align: left; font-size: 10px; }
          td { border-bottom: 1px solid #E2E8F0; padding: 6px; text-align: left; }
          tr:nth-child(even) { background: #F8FAFC; }
          .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 9px; font-weight: bold; background: #EFF6FF; color: #2563EB; }
          .footer { margin-top: 30px; font-size: 9px; color: #94A3B8; border-top: 1px solid #E2E8F0; padding-top: 8px; text-align: right; }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="subtitle">JOB TRAINING MANAGEMENT SYSTEM</div>
          <div class="title">${t.Name} [${t.Code}]</div>
          <div class="subtitle">Generated on ${rep.generatedAt}</div>
        </div>

        <div class="section">
          <div class="section-title">Programme Overview</div>
          <div class="grid">
            <div class="row"><div class="cell cell-label">Category:</div><div class="cell">${t.Category}</div><div class="cell cell-label">Trainer:</div><div class="cell">${t.Trainer}</div></div>
            <div class="row"><div class="cell cell-label">Start Date:</div><div class="cell">${t.StartDate}</div><div class="cell cell-label">End Date:</div><div class="cell">${t.EndDate || '—'}</div></div>
            <div class="row"><div class="cell cell-label">Duration:</div><div class="cell">${t.Duration} Day(s) (${t.TotalHours || 8} Hours)</div><div class="cell cell-label">Venue:</div><div class="cell">${t.Venue || '—'}</div></div>
            <div class="row"><div class="cell cell-label">Department:</div><div class="cell">${t.Department || 'All'}</div><div class="cell cell-label">Status / Stage:</div><div class="cell"><span class="badge">${t.Status}</span> (${t.Stage || 'Created'})</div></div>
            <div class="row"><div class="cell cell-label">Course Fee:</div><div class="cell">RM ${t.CourseFee || '0.00'}</div><div class="cell cell-label">Enrolled Pax:</div><div class="cell">${t.Participants || 0}</div></div>
          </div>
          ${t.Objectives ? `<div style="margin-top:6px;font-size:11px"><strong>Objectives:</strong> ${t.Objectives}</div>` : ''}
        </div>

        <div class="section">
          <div class="section-title">Key Performance Indicators</div>
          <div class="kpi-container">
            <div class="kpi-card"><div class="kpi-value">${rep.attendance.summary.pct || 0}%</div><div class="kpi-label">Attendance Rate</div></div>
            <div class="kpi-card"><div class="kpi-value">${rep.attendance.summary.present || 0}</div><div class="kpi-label">Present Count</div></div>
            <div class="kpi-card"><div class="kpi-value">${rep.attendance.summary.absent || 0}</div><div class="kpi-label">Absent Count</div></div>
            <div class="kpi-card"><div class="kpi-value">${rep.evaluation.summary.evalCompleted || 0}</div><div class="kpi-label">Evaluations Done</div></div>
            <div class="kpi-card"><div class="kpi-value">${rep.evaluation.summary.avgScore || 'N/A'}</div><div class="kpi-label">Avg Eval Score</div></div>
          </div>
        </div>

        <div class="section">
          <div class="section-title">Attendance Summary by Day</div>
          <table>
            <thead>
              <tr><th>Day</th><th>Date</th><th>Total Records</th><th>Present</th><th>Absent</th><th>Late</th><th>Rate (%)</th></tr>
            </thead>
            <tbody>
              ${(rep.attendance.days || []).map(d => {
                const recs = d.records || [];
                const pres = recs.filter(r => r.Status === 'Present').length;
                const abs  = recs.filter(r => r.Status === 'Absent').length;
                const late = recs.filter(r => r.Status === 'Late').length;
                const rate = recs.length ? Math.round(pres / recs.length * 100) : 0;
                return `<tr>
                  <td>Day ${d.day}</td>
                  <td>${d.date}</td>
                  <td>${recs.length}</td>
                  <td style="color:#16A34A;font-weight:bold">${pres}</td>
                  <td style="color:#DC2626;font-weight:bold">${abs}</td>
                  <td style="color:#D97706">${late}</td>
                  <td><strong>${rate}%</strong></td>
                </tr>`;
              }).join('') || '<tr><td colspan="7">No attendance records registered.</td></tr>'}
            </tbody>
          </table>
        </div>

        <div class="section">
          <div class="section-title">Evaluation & Post-Training Overview</div>
          <div class="grid">
            <div class="row">
              <div class="cell cell-label">Evaluations Submitted:</div>
              <div class="cell">${rep.evaluation.summary.evalCompleted || 0} participant(s)</div>
              <div class="cell cell-label">Average Score:</div>
              <div class="cell"><strong>${rep.evaluation.summary.avgScore || 'N/A'} / 5.0</strong></div>
            </div>
            <div class="row">
              <div class="cell cell-label">6-Month Post Evaluations:</div>
              <div class="cell">${rep.evaluation.summary.postCompleted || 0} completed</div>
              <div class="cell cell-label">Workspace Folder ID:</div>
              <div class="cell">${t.FolderID || 'N/A'}</div>
            </div>
          </div>
        </div>

        <div class="footer">
          TrainHub Job Training Management System &bull; Official Programme Summary PDF
        </div>
      </body>
      </html>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
    blob.setName(`${t.Code}_Training_Report.pdf`);

    let pdfUrl = '';
    let fileId = '';
    if (t.FolderID) {
      try {
        const folder = DriveApp.getFolderById(t.FolderID);
        const pdfFile = folder.createFile(blob);
        pdfUrl = pdfFile.getUrl();
        fileId = pdfFile.getId();
      } catch (e) {
        Logger.log('Could not save PDF to training folder: ' + e.message);
      }
    }

    const base64Data = Utilities.base64Encode(blob.getBytes());
    const dataUri = 'data:application/pdf;base64,' + base64Data;

    return ok({
      fileId: fileId,
      fileName: `${t.Code}_Training_Report.pdf`,
      pdfUrl: pdfUrl,
      dataUri: dataUri,
      message: 'PDF report generated successfully.'
    });
  } catch (e) {
    return err('Failed to export PDF report: ' + e.message);
  }
}