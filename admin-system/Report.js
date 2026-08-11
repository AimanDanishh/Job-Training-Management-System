/**
 * Report.gs — Comprehensive Report Generation & Excel Export Engine
 */

// ─── Full Programme Report (Single Training) ──────────────────────────────────
function generateReport(trainingId) {
  try {
    const tResult  = JSON.parse(getTrainingById(trainingId));
    if (!tResult.success) return err('Training not found.');
    const training = tResult.data;

    const attSummary  = JSON.parse(getAttendanceSummary(trainingId)).data  || {};
    const evalSummary = JSON.parse(getEvaluationSummary(trainingId)).data  || {};
    const attData     = JSON.parse(getAttendance(trainingId)).data         || [];
    const evalData    = JSON.parse(getTrainingEvaluations(trainingId)).data || [];
    const postData    = JSON.parse(getPostEvaluations(trainingId)).data    || [];

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
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const tRows  = tSheet ? sheetToJson(tSheet) : [];

    const monthly = buildMonthlyData(tRows);

    let aRows = [];
    tRows.forEach(t => {
      if (t.ID) {
        const ss = getTrainingDataSpreadsheet(t.ID);
        if (ss) {
          const sheet = ss.getSheetByName('Attendance');
          if (sheet) aRows = aRows.concat(sheetToJson(sheet));
        }
      }
    });
    const deptData = buildDeptData(aRows);

    // Compute Multi-Tier Approval Status Pipeline breakdown
    let pendingHodCount = 0;
    let pendingCsuiteCount = 0;
    let pendingHohrCount = 0;
    let approvedCount = 0;
    let rejectedCount = 0;

    tRows.forEach(r => {
      const appStatus = String(r.ApprovalStatus || '').trim();
      if (appStatus === 'Pending HOD Approval') pendingHodCount++;
      else if (appStatus === 'Pending C-Suite Approval') pendingCsuiteCount++;
      else if (appStatus === 'Pending HOHR Approval') pendingHohrCount++;
      else if (appStatus === 'Rejected') rejectedCount++;
      else approvedCount++; // 'Approved', 'Auto-Approved', or default
    });

    const approvalSummary = {
      pendingHod:    pendingHodCount,
      pendingCsuite: pendingCsuiteCount,
      pendingHohr:   pendingHohrCount,
      approved:      approvedCount,
      rejected:      rejectedCount,
      total:         tRows.length
    };

    return ok({
      summary:     trainingSummary,
      approval:    approvalSummary,
      monthly:     monthly,
      departments: deptData
    });
  } catch (e) {
    return err('Dashboard report error: ' + e.message);
  }
}

// ─── Unified Filtered Reports Engine ──────────────────────────────────────────
/**
 * Generates structured report data based on report type and filter parameters.
 * Types: 'hours', 'cost', 'title', 'employee', 'atp', 'single'
 */
function getFilteredReportData(reportType, filters) {
  try {
    const opts = filters || {};
    const selectedYear  = String(opts.year || '').trim();
    const selectedMonth = String(opts.month || '').trim(); // e.g. '01', '08', or 'Jan', 'Feb'
    const selectedTitle = String(opts.title || '').trim().toLowerCase();
    const selectedDept  = String(opts.costCentre || opts.department || '').trim().toLowerCase();
    const selectedCat   = String(opts.category || '').trim().toLowerCase();

    const tSheet = getSheet(SHEET_NAMES.trainings);
    const tRows  = tSheet ? sheetToJson(tSheet) : [];
    const empSheet = getSheet(SHEET_NAMES.employees);
    const empRows = empSheet ? sheetToJson(empSheet) : [];

    // Filter trainings by year, month, title, category, department
    const filteredTrainings = tRows.filter(t => {
      const startDateStr = t.StartDate || '';
      let matchYear = true;
      let matchMonth = true;

      if (startDateStr) {
        const d = parseDateObj(startDateStr);
        if (d) {
          const yStr = String(d.getFullYear());
          const mNumStr = String(d.getMonth() + 1).padStart(2, '0');
          const mNameShort = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM');

          if (selectedYear && selectedYear !== 'All') {
            matchYear = (yStr === selectedYear);
          }
          if (selectedMonth && selectedMonth !== 'All') {
            matchMonth = (mNumStr === selectedMonth || mNameShort.toLowerCase() === selectedMonth.toLowerCase() || String(d.getMonth() + 1) === selectedMonth);
          }
        }
      }

      const matchTitle = !selectedTitle || String(t.Name || t.ID || t.Code || '').toLowerCase().includes(selectedTitle);
      const matchDept  = !selectedDept  || String(t.Department || '').toLowerCase().includes(selectedDept);
      const matchCat   = !selectedCat   || String(t.Category || '').toLowerCase().includes(selectedCat);

      return matchYear && matchMonth && matchTitle && matchDept && matchCat;
    });

    if (reportType === 'hours') {
      return ok(buildHoursReportData(filteredTrainings, empRows, selectedYear || '2026'));
    } else if (reportType === 'cost') {
      return ok(buildCostReportData(filteredTrainings, selectedYear || '2026'));
    } else if (reportType === 'title') {
      return ok(buildTrainingTitleReportData(filteredTrainings, empRows));
    } else if (reportType === 'employee') {
      return ok(buildEmployeeReportData(filteredTrainings, empRows, selectedDept));
    } else if (reportType === 'atp') {
      return ok(buildAnnualTrainingPlanData(filteredTrainings));
    } else {
      return err('Invalid report type specified.');
    }
  } catch (e) {
    return err('Failed to generate filtered report: ' + e.message);
  }
}

// ─── 1. Training Hours Report Data Builder (Cost Centre vs Month) ─────────────
function buildHoursReportData(trainings, employees, year) {
  const yrSuffix = year && year.length === 4 ? year.slice(-2) : '26';
  const monthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].map(m => `${m}-${yrSuffix}`);

  // Fetch unique cost centres
  const costCentresSet = new Set();
  employees.forEach(e => {
    const cc = String(e.CostCentre || e['Cost Centre'] || e.Department || '').trim();
    if (cc) costCentresSet.add(cc);
  });
  trainings.forEach(t => {
    const cc = String(t.Department || '').trim();
    if (cc && cc !== 'All Departments / Cost Centres') costCentresSet.add(cc);
  });
  if (costCentresSet.size === 0) {
    costCentresSet.add('General');
  }

  const costCentreList = Array.from(costCentresSet).sort();
  const matrix = {};

  costCentreList.forEach(cc => {
    matrix[cc] = Array(12).fill(0);
  });

  // Calculate training hours per cost centre per month
  trainings.forEach(t => {
    const startDate = parseDateObj(t.StartDate);
    if (!startDate) return;
    const monthIdx = startDate.getMonth(); // 0 to 11
    const hours = Number(t.TotalHours || t.Duration * 8 || 8);

    // Get participants for this training to attribute hours accurately by cost centre
    const participants = getTrainingParticipantsList(t.ID);
    if (participants.length > 0) {
      participants.forEach(p => {
        const cc = String(p.Department || p.CostCentre || t.Department || 'General').trim();
        if (!matrix[cc]) {
          matrix[cc] = Array(12).fill(0);
        }
        matrix[cc][monthIdx] += hours;
      });
    } else {
      const cc = String(t.Department || 'General').trim();
      if (!matrix[cc]) matrix[cc] = Array(12).fill(0);
      matrix[cc][monthIdx] += hours;
    }
  });

  const rows = [];
  costCentreList.forEach(cc => {
    const mData = matrix[cc] || Array(12).fill(0);
    const totalYear = mData.reduce((a, b) => a + b, 0);
    const totalCumulative = totalYear; // Can accumulate across years if expanded

    rows.push({
      costCentre: cc,
      months: mData,
      totalYear: totalYear,
      totalHours: totalCumulative
    });
  });

  return {
    year: year,
    monthHeaders: monthLabels,
    rows: rows
  };
}

// ─── 2. Training Cost Report Data Builder ─────────────────────────────────────
function buildCostReportData(trainings, year) {
  const rows = trainings.map(t => {
    return {
      trainingTitle: t.Name,
      dateFrom: formatDate(t.StartDate),
      dateTo: formatDate(t.EndDate || t.StartDate),
      totalParticipant: Number(t.Participants || 0),
      // Cost items left empty for user input / HR verification as requested
      trainingFees: '',
      meal: '',
      subsistanceAllowance: '',
      hotelFees: '',
      mileageClaim: '',
      taxiFees: '',
      tollFees: '',
      flight: '',
      totalCost: t.CourseFee ? Number(t.CourseFee).toFixed(2) : '0.00',
      totalHrdfGrant: ''
    };
  });

  return {
    year: year,
    rows: rows
  };
}

// ─── 3. Training Title Report Data Builder ────────────────────────────────────
function buildTrainingTitleReportData(trainings, employees) {
  const rows = [];

  trainings.forEach(t => {
    const participants = getTrainingParticipantsList(t.ID);
    const startDate = parseDateObj(t.StartDate);
    const monthName = startDate ? Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMMM') : '—';
    const totalHours = Number(t.TotalHours || t.Duration * 8 || 8);

    if (participants.length > 0) {
      participants.forEach(p => {
        rows.push({
          name: p.Name || p.EmployeeName || '—',
          empNo: p.ID || p.EmployeeID || p.EmployeeNo || '—',
          trainingTitle: t.Name,
          costCentreDesc: p.Department || t.Department || '—',
          trainingType: t.Category && t.Category.toLowerCase().includes('in-house') ? 'In-House Training' : 'Public',
          dateFrom: formatDate(t.StartDate),
          dateUntil: formatDate(t.EndDate || t.StartDate),
          month: monthName,
          totalHours: totalHours,
          trainer: t.Trainer || '—',
          trainingProvider: t.Trainer || 'Internal / External',
          expiryDate: 'N/A'
        });
      });
    } else {
      rows.push({
        name: '—',
        empNo: '—',
        trainingTitle: t.Name,
        costCentreDesc: t.Department || '—',
        trainingType: 'Public',
        dateFrom: formatDate(t.StartDate),
        dateUntil: formatDate(t.EndDate || t.StartDate),
        month: monthName,
        totalHours: totalHours,
        trainer: t.Trainer || '—',
        trainingProvider: t.Trainer || 'Internal / External',
        expiryDate: 'N/A'
      });
    }
  });

  return { rows: rows };
}

// ─── 4. Employee Report Data Builder ──────────────────────────────────────────
function buildEmployeeReportData(trainings, employees, filterDept) {
  const empMap = {};

  employees.forEach((e, idx) => {
    const id = String(e.ID || e.EmployeeID || e.EmployeeNo || '').trim();
    if (!id) return;
    empMap[id.toLowerCase()] = {
      no: idx + 1,
      empNo: id,
      name: e.Name || e.EmployeeName || id,
      costCentre: e.CostCentre || e['Cost Centre'] || e.Department || '—',
      position: e.Position || e.JobTitle || e.PositionTitle || 'Staff',
      totalTrainings: 0,
      totalHours: 0,
      attendedList: [],
      status: e.Status || 'Active'
    };
  });

  trainings.forEach(t => {
    const participants = getTrainingParticipantsList(t.ID);
    const hours = Number(t.TotalHours || t.Duration * 8 || 8);

    participants.forEach(p => {
      const id = String(p.ID || p.EmployeeID || p.EmployeeNo || '').trim().toLowerCase();
      if (empMap[id]) {
        empMap[id].totalTrainings += 1;
        empMap[id].totalHours += hours;
        empMap[id].attendedList.push(t.Name);
      }
    });
  });

  let list = Object.values(empMap);
  if (filterDept) {
    list = list.filter(e => e.costCentre.toLowerCase().includes(filterDept));
  }

  return { rows: list };
}

// ─── 5. Annual Training Plan (ATP) Data Builder ───────────────────────────────
function buildAnnualTrainingPlanData(trainings) {
  const rows = trainings.map((t, idx) => {
    return {
      no: idx + 1,
      trainingTitle: t.Name,
      trainingCategory: t.Category || 'General',
      tnaSource: 'Annual TNA',
      trainingMode: t.Venue && t.Venue.toLowerCase().includes('online') ? 'Online' : 'Physical Classroom',
      durationHours: Number(t.TotalHours || t.Duration * 8 || 8),
      trainer: t.Trainer || '—',
      department: t.Department || 'All Departments',
      positionEmpNo: 'All Eligible Staff',
      totalPax: Number(t.Participants || 0),
      plannedDate: formatDate(t.StartDate),
      actualDateFrom: formatDate(t.StartDate),
      actualDateTo: formatDate(t.EndDate || t.StartDate),
      trainingStatus: t.Status || 'Draft',
      remarks: t.Stage || 'Planned'
    };
  });

  return { rows: rows };
}

// ─── Helper to fetch participants for a training ──────────────────────────────
function getTrainingParticipantsList(trainingId) {
  try {
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (ss) {
      const sheet = ss.getSheetByName('TrainingParticipants');
      if (sheet) return sheetToJson(sheet);
    }
  } catch (e) {}
  return [];
}

function parseDateObj(str) {
  if (!str) return null;
  const d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  const dmy = String(str).match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) return new Date(dmy[3], dmy[2] - 1, dmy[1]);
  return null;
}

// ─── Export Filtered Report to Excel / Google Sheets ──────────────────────────
/**
 * Creates formatted Google Sheet & returns XLSX download link matching user's exact header specs.
 * Supports single report exports as well as 'master' / 'all' multi-tab exports.
 */
function exportFilteredReportExcel(reportType, filters) {
  try {
    if (reportType === 'master' || reportType === 'all') {
      return exportAllInOneMasterReportExcel(filters);
    }

    const res = getFilteredReportData(reportType, filters);
    if (!res.success) return err(res.message);
    const data = res.data;
    const year = (filters && filters.year) ? filters.year : '2026';
    const timestamp = Date.now();

    let titleName = 'Report';
    if (reportType === 'hours') titleName = `Training_Hours_CostCentre_${year}_${timestamp}`;
    else if (reportType === 'cost') titleName = `Training_Cost_${year}_${timestamp}`;
    else if (reportType === 'title') titleName = `Training_Title_Report_${timestamp}`;
    else if (reportType === 'employee') titleName = `Employee_Training_Report_${timestamp}`;
    else if (reportType === 'atp') titleName = `Annual_Training_Plan_${year}_${timestamp}`;

    const ss = SpreadsheetApp.create(titleName);
    const sheet = ss.getActiveSheet();
    sheet.setName(titleName.slice(0, 30));

    if (reportType === 'hours') {
      formatHoursReportSheet(sheet, data, year);
    } else if (reportType === 'cost') {
      formatCostReportSheet(sheet, data, year);
    } else if (reportType === 'title') {
      formatTitleReportSheet(sheet, data);
    } else if (reportType === 'employee') {
      formatEmployeeReportSheet(sheet, data);
    } else if (reportType === 'atp') {
      formatAtpReportSheet(sheet, data);
    }

    SpreadsheetApp.flush();

    const fileId = ss.getId();
    const sheetUrl = ss.getUrl();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    return ok({
      fileId: fileId,
      fileName: `${titleName}.xlsx`,
      sheetUrl: sheetUrl,
      downloadUrl: downloadUrl,
      message: 'Excel report generated successfully!'
    });
  } catch (e) {
    return err('Failed to export Excel report: ' + e.message);
  }
}

/**
 * Creates a single Master Google Spreadsheet containing ALL 5 reports on separate tabs:
 * 1. Training Hours (Cost Centre vs Month)
 * 2. Training Cost
 * 3. Training Title
 * 4. Employee Report
 * 5. Annual Training Plan (ATP)
 */
function exportAllInOneMasterReportExcel(filters) {
  try {
    const year = (filters && filters.year) ? filters.year : '2026';
    const timestamp = Date.now();
    const fileName = `Master_Annual_Training_Report_${year}_${timestamp}`;

    const ss = SpreadsheetApp.create(fileName);

    // Tab 1: Training Hours
    const sHours = ss.getActiveSheet();
    sHours.setName('Training Hours');
    const hoursRes = getFilteredReportData('hours', filters);
    if (hoursRes.success) formatHoursReportSheet(sHours, hoursRes.data, year);

    // Tab 2: Training Cost
    const sCost = ss.insertSheet('Training Cost');
    const costRes = getFilteredReportData('cost', filters);
    if (costRes.success) formatCostReportSheet(sCost, costRes.data, year);

    // Tab 3: Training Title
    const sTitle = ss.insertSheet('Training Title');
    const titleRes = getFilteredReportData('title', filters);
    if (titleRes.success) formatTitleReportSheet(sTitle, titleRes.data);

    // Tab 4: Employee Report
    const sEmp = ss.insertSheet('Employee Report');
    const empRes = getFilteredReportData('employee', filters);
    if (empRes.success) formatEmployeeReportSheet(sEmp, empRes.data);

    // Tab 5: Annual Training Plan (ATP)
    const sAtp = ss.insertSheet('Annual Training Plan (ATP)');
    const atpRes = getFilteredReportData('atp', filters);
    if (atpRes.success) formatAtpReportSheet(sAtp, atpRes.data);

    SpreadsheetApp.flush();

    const fileId = ss.getId();
    const sheetUrl = ss.getUrl();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    return ok({
      fileId: fileId,
      fileName: `${fileName}.xlsx`,
      sheetUrl: sheetUrl,
      downloadUrl: downloadUrl,
      message: `Master All-in-One Report Workbook created with 5 tabs (Training Hours, Training Cost, Training Title, Employee, ATP)!`
    });
  } catch (e) {
    return err('Failed to export Master Report Workbook: ' + e.message);
  }
}

// ─── Header Formatters for Excel Export ──────────────────────────────────────

// 1. Hours Header Formatter:
// Row 1: A1-A2: (Cost Centre/Month) | B-M (row 1: Training Hours) | N1-N2: Total 2026 | O1-O2: Total Training Hours 2026
// Row 2: B-M (row 2: Jan-26, Feb-26, ...)
function formatHoursReportSheet(sheet, data, year) {
  const yrSuffix = year && year.length === 4 ? year.slice(-2) : '26';
  const monthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].map(m => `${m}-${yrSuffix}`);

  // Row 1
  const row1 = ['Cost Centre/Month', 'Training Hours', '', '', '', '', '', '', '', '', '', '', '', `Total ${year}`, `Total Training Hours ${year}`];
  sheet.appendRow(row1);

  // Row 2
  const row2 = ['', ...monthLabels, '', ''];
  sheet.appendRow(row2);

  // Merge headers
  sheet.getRange('A1:A2').merge().setValue('Cost Centre/Month');
  sheet.getRange('B1:M1').merge().setValue('Training Hours');
  sheet.getRange('N1:N2').merge().setValue(`Total ${year}`);
  sheet.getRange('O1:O2').merge().setValue(`Total Training Hours ${year}`);

  // Header Styling
  const headerRange = sheet.getRange('A1:O2');
  headerRange.setFontWeight('bold')
    .setBackground('#2563EB')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // Append data rows
  (data.rows || []).forEach(r => {
    sheet.appendRow([
      r.costCentre,
      ...r.months,
      r.totalYear,
      r.totalHours
    ]);
  });

  sheet.autoResizeColumns(1, 15);
}

// 2. Cost Header Formatter:
// Row 1: A1-A2: Training Title | B1-B2: Training Date (From) | C1-C2: Training Date (To) | D1-D2: Total Participant | E1-L1: Training Cost (RM) | M1-M2: Total Cost 2026 | N1-N2: Total HRDF Grant (RM)
// Row 2: E2: Training Fees | F2: Meal | G2: Subsistance Allowance | H2: Hotel Fees | I2: Mileage Claim | J2: Taxi Fees | K2: Toll Fees | L2: Flight
function formatCostReportSheet(sheet, data, year) {
  // Row 1
  const row1 = [
    'Training Title', 'Training Date (From)', 'Training Date (To)', 'Total Participant',
    'Training Cost (RM)', '', '', '', '', '', '', '',
    `Total Cost ${year}`, 'Total HRDF Grant (RM)'
  ];
  sheet.appendRow(row1);

  // Row 2
  const row2 = [
    '', '', '', '',
    'Training Fees', 'Meal', 'Subsistance Allowance', 'Hotel Fees', 'Mileage Claim', 'Taxi Fees', 'Toll Fees', 'Flight',
    '', ''
  ];
  sheet.appendRow(row2);

  // Merge headers
  sheet.getRange('A1:A2').merge().setValue('Training Title');
  sheet.getRange('B1:B2').merge().setValue('Training Date (From)');
  sheet.getRange('C1:C2').merge().setValue('Training Date (To)');
  sheet.getRange('D1:D2').merge().setValue('Total Participant');
  sheet.getRange('E1:L1').merge().setValue('Training Cost (RM)');
  sheet.getRange('M1:M2').merge().setValue(`Total Cost ${year}`);
  sheet.getRange('N1:N2').merge().setValue('Total HRDF Grant (RM)');

  // Header Styling
  const headerRange = sheet.getRange('A1:N2');
  headerRange.setFontWeight('bold')
    .setBackground('#2563EB')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // Append data rows
  (data.rows || []).forEach(r => {
    sheet.appendRow([
      r.trainingTitle,
      r.dateFrom,
      r.dateTo,
      r.totalParticipant,
      r.trainingFees,
      r.meal,
      r.subsistanceAllowance,
      r.hotelFees,
      r.mileageClaim,
      r.taxiFees,
      r.tollFees,
      r.flight,
      r.totalCost,
      r.totalHrdfGrant
    ]);
  });

  sheet.autoResizeColumns(1, 14);
}

// 3. Training Title Header Formatter:
// Name | Emp. No | Training Title | (EE)/Cost Centre Description | Training Type | Date (From) | Date (Until) | Month | Total Hours | Trainer | Training Provide | Expiry Date (if applicable)
function formatTitleReportSheet(sheet, data) {
  const headers = [
    'Name', 'Emp. No', 'Training Title', '(EE)/Cost Centre Description',
    'Training Type', 'Date (From)', 'Date (Until)', 'Month',
    'Total Hours', 'Trainer', 'Training Provide', 'Expiry Date (if applicable)'
  ];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');

  (data.rows || []).forEach(r => {
    sheet.appendRow([
      r.name, r.empNo, r.trainingTitle, r.costCentreDesc,
      r.trainingType, r.dateFrom, r.dateUntil, r.month,
      r.totalHours, r.trainer, r.trainingProvider, r.expiryDate
    ]);
  });

  sheet.autoResizeColumns(1, headers.length);
}

// 4. Employee Header Formatter
function formatEmployeeReportSheet(sheet, data) {
  const headers = ['No', 'Emp. No', 'Employee Name', 'Cost Centre / Department', 'Position', 'Total Trainings', 'Total Hours', 'Attended Trainings', 'Status'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');

  (data.rows || []).forEach(r => {
    sheet.appendRow([
      r.no, r.empNo, r.name, r.costCentre, r.position, r.totalTrainings, r.totalHours, r.attendedList.join(', '), r.status
    ]);
  });

  sheet.autoResizeColumns(1, headers.length);
}

// 5. ATP Header Formatter:
// No | Training Title | Training Category | TNA Source | Training Mode | Training Duration (hrs) | Trainer | Department | Position / Employee No | Total Pax | Planned Date | Actual Date (From) | Actual Date (To) | Training Status | Remarks
function formatAtpReportSheet(sheet, data) {
  const headers = [
    'No', 'Training Title', 'Training Category', 'TNA Source', 'Training Mode',
    'Training Duration (hrs)', 'Trainer', 'Department', 'Position / Employee No',
    'Total Pax', 'Planned Date', 'Actual Date (From)', 'Actual Date (To)',
    'Training Status', 'Remarks'
  ];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');

  (data.rows || []).forEach(r => {
    sheet.appendRow([
      r.no, r.trainingTitle, r.trainingCategory, r.tnaSource, r.trainingMode,
      r.durationHours, r.trainer, r.department, r.positionEmpNo,
      r.totalPax, r.plannedDate, r.actualDateFrom, r.actualDateTo,
      r.trainingStatus, r.remarks
    ]);
  });

  sheet.autoResizeColumns(1, headers.length);
}

// ─── Export Single Attendance Sheet ───────────────────────────────────────────
function exportAttendanceSheet(trainingId) {
  try {
    const tResult = JSON.parse(getTrainingById(trainingId));
    if (!tResult.success) return err('Training not found.');
    const training = tResult.data;

    const attData = JSON.parse(getAttendance(trainingId)).data || [];

    const ss    = getSpreadsheet();
    const name  = 'Attendance_' + training.Code + '_' + new Date().getTime();
    const sheet = ss.insertSheet(name);

    const headers = ['Employee ID', 'Employee Name', 'Department',
                     'Day', 'Date', 'Check In', 'Check Out', 'Hours', 'Status', 'Remarks'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');

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

// ─── Export Single Training Full Report to Excel (.xlsx / Sheet) ──────────────
function exportReportExcel(trainingId) {
  try {
    const reportRes = JSON.parse(generateReport(trainingId));
    if (!reportRes.success) return err(reportRes.message);
    const rep = reportRes.data;
    const t = rep.training;

    const ssName = `${t.Code}_Full_Report_${Date.now()}`;
    const ss = SpreadsheetApp.create(ssName);

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

    const sAtt = ss.insertSheet('Attendance Records');
    const attHeaders = ['Day', 'Date', 'Employee ID', 'Employee Name', 'Department', 'Check In', 'Check Out', 'Hours', 'Status', 'Remarks'];
    sAtt.appendRow(attHeaders);
    sAtt.getRange(1, 1, 1, attHeaders.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
    (rep.attendance.days || []).forEach(day => {
      (day.records || []).forEach(r => {
        sAtt.appendRow([r.Day, r.Date, r.EmployeeID, r.EmployeeName, r.Department, r.CheckIn, r.CheckOut, r.Hours, r.Status, r.Remarks]);
      });
    });

    const sEval = ss.insertSheet('Training Evaluations');
    const evalHeaders = ['Evaluation ID', 'Employee ID', 'Employee Name', 'Q1 (Obj)', 'Q2 (Rel)', 'Q3 (Mat)', 'Q4 (Trn)', 'Q5 (Eng)', 'Q6 (Dur)', 'Q7 (App)', 'Avg Score', 'Section B1', 'Section B2', 'Section B3', 'Submitted At'];
    sEval.appendRow(evalHeaders);
    sEval.getRange(1, 1, 1, evalHeaders.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
    (rep.evaluation.responses || []).forEach(ev => {
      sEval.appendRow([ev.ID, ev.EmployeeID, ev.EmployeeName, ev.Q1, ev.Q2, ev.Q3, ev.Q4, ev.Q5, ev.Q6, ev.Q7, ev.AvgScore, ev.SectionB1, ev.SectionB2, ev.SectionB3, ev.SubmittedAt]);
    });

    const sPost = ss.insertSheet('3-Month Post Evaluations');
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
              <div class="cell cell-label">3-Month Post Evaluations:</div>
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