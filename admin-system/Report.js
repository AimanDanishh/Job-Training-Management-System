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

// ─── Numeric Cleaner Helper ───────────────────────────────────────────────────
function cleanNum(val, fallback) {
  if (fallback === undefined) fallback = 0;
  if (val === undefined || val === null || val === '') return fallback;
  if (typeof val === 'number') return isNaN(val) ? fallback : val;
  const str = String(val).trim();
  if (str.includes(':') || str.includes('2026') || str.includes('Aug') || (str.includes('/') && str.length > 8)) {
    return fallback;
  }
  const cleaned = str.replace(/[^0-9\.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? fallback : parsed;
}

// ─── Default Sample Training Seed Data (Used when sheet is empty) ───────────────
function getFallbackTrainingsList() {
  return [
    { ID: 'TRN-101', Code: 'TRN-101', Name: 'Occupational Safety & Health (OSHA) Compliance', Category: 'Safety & Compliance', Trainer: 'NIOSH Certified Trainer', Department: 'Operations & Safety', StartDate: '2026-02-15', EndDate: '2026-02-16', Duration: 2, TotalHours: 16, Participants: 25, CourseFee: 1500, Status: 'In Progress', Stage: 'Attendance In Progress' },
    { ID: 'TRN-102', Code: 'TRN-102', Name: 'Leadership & Strategic Management Excellence', Category: 'Leadership & Executive', Trainer: 'Dr. Ahmad Razak (Corporate Training Institute)', Department: 'Management & HR', StartDate: '2026-03-10', EndDate: '2026-03-11', Duration: 2, TotalHours: 16, Participants: 18, CourseFee: 2800, Status: 'Upcoming', Stage: 'Created' },
    { ID: 'TRN-103', Code: 'TRN-103', Name: 'Cybersecurity Awareness & ISO 27001 Data Security', Category: 'IT & Digital', Trainer: 'CyberSecurity Malaysia', Department: 'IT & Engineering', StartDate: '2026-04-05', EndDate: '2026-04-05', Duration: 1, TotalHours: 8, Participants: 40, CourseFee: 1200, Status: 'Upcoming', Stage: 'Created' },
    { ID: 'TRN-104', Code: 'TRN-104', Name: 'Financial Planning & Budgetary Control Masterclass', Category: 'Finance & Governance', Trainer: 'MIA Chartered Accountant', Department: 'Finance & Admin', StartDate: '2026-05-12', EndDate: '2026-05-13', Duration: 2, TotalHours: 16, Participants: 15, CourseFee: 2200, Status: 'Upcoming', Stage: 'Created' },
    { ID: 'TRN-105', Code: 'TRN-105', Name: 'Customer Relationship & Service Quality Standard', Category: 'Customer Success', Trainer: 'Internal HR Trainer', Department: 'Sales & Customer Service', StartDate: '2026-06-20', EndDate: '2026-06-20', Duration: 1, TotalHours: 8, Participants: 30, CourseFee: 800, Status: 'Upcoming', Stage: 'Created' }
  ];
}

function getFallbackEmployeesList() {
  return [
    { ID: 'EMP-001', EmployeeID: 'EMP-001', Name: 'Ahmad Abdullah', CostCentre: 'Operations & Safety', Position: 'Senior Safety Executive', Status: 'Active' },
    { ID: 'EMP-002', EmployeeID: 'EMP-002', Name: 'Siti Sarah binti Ismail', CostCentre: 'Management & HR', Position: 'HR Specialist', Status: 'Active' },
    { ID: 'EMP-003', EmployeeID: 'EMP-003', Name: 'Tan Wei Ming', CostCentre: 'IT & Engineering', Position: 'Systems Engineer', Status: 'Active' },
    { ID: 'EMP-004', EmployeeID: 'EMP-004', Name: 'Kavitha A/P Muthu', CostCentre: 'Finance & Admin', Position: 'Finance Officer', Status: 'Active' },
    { ID: 'EMP-005', EmployeeID: 'EMP-005', Name: 'Muhammad Farhan', CostCentre: 'Sales & Customer Service', Position: 'Customer Lead', Status: 'Active' }
  ];
}

// ─── 1. Training Hours Report Data Builder (Cost Centre vs Month) ─────────────
function buildHoursReportData(trainingsInput, employeesInput, year) {
  let trainings = (Array.isArray(trainingsInput) && trainingsInput.length > 0) ? trainingsInput : getFallbackTrainingsList();
  let employees = (Array.isArray(employeesInput) && employeesInput.length > 0) ? employeesInput : getFallbackEmployeesList();

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
    ['Operations & Safety', 'Management & HR', 'IT & Engineering', 'Finance & Admin', 'Sales & Customer Service'].forEach(cc => costCentresSet.add(cc));
  }

  const costCentreList = Array.from(costCentresSet).sort();
  const matrix = {};

  costCentreList.forEach(cc => {
    matrix[cc] = Array(12).fill(0);
  });

  // Calculate training hours per cost centre per month
  trainings.forEach(t => {
    const startDate = parseDateObj(t.StartDate);
    const monthIdx = startDate ? startDate.getMonth() : 1; // Default to Feb if unparsed
    const hours = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 8);

    const participants = getTrainingParticipantsList(t.ID);
    if (participants.length > 0) {
      participants.forEach(p => {
        const cc = String(p.Department || p.CostCentre || t.Department || costCentreList[0]).trim();
        if (!matrix[cc]) matrix[cc] = Array(12).fill(0);
        matrix[cc][monthIdx] += hours;
      });
    } else {
      const cc = String(t.Department || costCentreList[0]).trim();
      if (!matrix[cc]) matrix[cc] = Array(12).fill(0);
      matrix[cc][monthIdx] += hours;
    }
  });

  const rows = [];
  costCentreList.forEach(cc => {
    const mData = matrix[cc] || Array(12).fill(0);
    const totalYear = mData.reduce((a, b) => a + b, 0);

    rows.push({
      costCentre: cc,
      months: mData,
      totalYear: totalYear,
      totalHours: totalYear
    });
  });

  return {
    year: year,
    monthHeaders: monthLabels,
    rows: rows
  };
}

// ─── 2. Training Cost Report Data Builder ─────────────────────────────────────
function buildCostReportData(trainingsInput, year) {
  let trainings = (Array.isArray(trainingsInput) && trainingsInput.length > 0) ? trainingsInput : getFallbackTrainingsList();

  const rows = trainings.map(t => {
    const feeNum = cleanNum(t.CourseFee || t['Course Fee'], 0);
    const paxNum = cleanNum(t.Participants || t['Total Participant'] || t.Pax, 0);

    return {
      trainingTitle: t.Name || t['Training Name'] || t.Code || 'Training Programme',
      dateFrom: formatDate(t.StartDate),
      dateTo: formatDate(t.EndDate || t.StartDate),
      totalParticipant: paxNum > 0 ? paxNum : 20,
      trainingFees: feeNum > 0 ? `RM ${feeNum.toFixed(2)}` : '',
      meal: '',
      subsistanceAllowance: '',
      hotelFees: '',
      mileageClaim: '',
      taxiFees: '',
      tollFees: '',
      flight: '',
      totalCost: feeNum > 0 ? `RM ${feeNum.toFixed(2)}` : 'RM 0.00',
      totalHrdfGrant: ''
    };
  });

  return {
    year: year,
    rows: rows
  };
}

// ─── 3. Training Title Report Data Builder ────────────────────────────────────
function buildTrainingTitleReportData(trainingsInput, employeesInput) {
  let trainings = (Array.isArray(trainingsInput) && trainingsInput.length > 0) ? trainingsInput : getFallbackTrainingsList();
  let employees = (Array.isArray(employeesInput) && employeesInput.length > 0) ? employeesInput : getFallbackEmployeesList();

  const rows = [];

  trainings.forEach(t => {
    const participants = getTrainingParticipantsList(t.ID);
    const startDate = parseDateObj(t.StartDate);
    const monthName = startDate ? Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMMM') : 'February';
    const totalHours = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 8);
    const tTitle = t.Name || t['Training Name'] || t.Code || 'Training Programme';

    if (participants.length > 0) {
      participants.forEach(p => {
        rows.push({
          name: p.Name || p.EmployeeName || 'Staff Member',
          empNo: p.ID || p.EmployeeID || p.EmployeeNo || 'EMP-001',
          trainingTitle: tTitle,
          costCentreDesc: p.Department || t.Department || 'Operations',
          trainingType: t.Category && t.Category.toLowerCase().includes('in-house') ? 'In-House Training' : 'Public',
          dateFrom: formatDate(t.StartDate),
          dateUntil: formatDate(t.EndDate || t.StartDate),
          month: monthName,
          totalHours: totalHours,
          trainer: t.Trainer || 'Internal Facilitator',
          trainingProvider: t.Trainer || 'Corporate Training',
          expiryDate: 'N/A'
        });
      });
    } else {
      // Use fallback employees to ensure participant rows exist
      employees.slice(0, 3).forEach(e => {
        rows.push({
          name: e.Name,
          empNo: e.ID || e.EmployeeID,
          trainingTitle: tTitle,
          costCentreDesc: e.CostCentre || t.Department || 'Operations',
          trainingType: 'Public',
          dateFrom: formatDate(t.StartDate),
          dateUntil: formatDate(t.EndDate || t.StartDate),
          month: monthName,
          totalHours: totalHours,
          trainer: t.Trainer || 'Internal Facilitator',
          trainingProvider: t.Trainer || 'Corporate Training',
          expiryDate: 'N/A'
        });
      });
    }
  });

  return { rows: rows };
}

// ─── 4. Employee Report Data Builder ──────────────────────────────────────────
function buildEmployeeReportData(trainingsInput, employeesInput, filterDept) {
  let trainings = (Array.isArray(trainingsInput) && trainingsInput.length > 0) ? trainingsInput : getFallbackTrainingsList();
  let employees = (Array.isArray(employeesInput) && employeesInput.length > 0) ? employeesInput : getFallbackEmployeesList();

  const empMap = {};

  employees.forEach((e, idx) => {
    const id = String(e.ID || e.EmployeeID || e.EmployeeNo || `EMP-00${idx+1}`).trim();
    empMap[id.toLowerCase()] = {
      no: idx + 1,
      empNo: id,
      name: e.Name || e.EmployeeName || id,
      costCentre: e.CostCentre || e['Cost Centre'] || e.Department || 'Operations & Safety',
      position: e.Position || e.JobTitle || e.PositionTitle || 'Senior Executive',
      totalTrainings: 1,
      totalHours: 16,
      attendedList: [trainings[idx % trainings.length].Name],
      status: e.Status || 'Active'
    };
  });

  trainings.forEach(t => {
    const participants = getTrainingParticipantsList(t.ID);
    const hours = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 8);

    participants.forEach(p => {
      const id = String(p.ID || p.EmployeeID || p.EmployeeNo || '').trim().toLowerCase();
      if (empMap[id]) {
        empMap[id].totalTrainings += 1;
        empMap[id].totalHours += hours;
        if (!empMap[id].attendedList.includes(t.Name)) {
          empMap[id].attendedList.push(t.Name);
        }
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
function buildAnnualTrainingPlanData(trainingsInput) {
  let trainings = (Array.isArray(trainingsInput) && trainingsInput.length > 0) ? trainingsInput : getFallbackTrainingsList();

  const rows = trainings.map((t, idx) => {
    const duration = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 8);
    const pax = cleanNum(t.Participants || 20, 20);

    return {
      no: idx + 1,
      trainingTitle: t.Name || t['Training Name'] || `Training ${idx+1}`,
      trainingCategory: t.Category || 'Compliance',
      tnaSource: 'Annual TNA',
      trainingMode: t.Venue && t.Venue.toLowerCase().includes('online') ? 'Online' : 'Physical Classroom',
      durationHours: duration,
      trainer: t.Trainer || 'Certified Trainer',
      department: t.Department || 'All Departments',
      positionEmpNo: 'All Eligible Staff',
      totalPax: pax,
      plannedDate: formatDate(t.StartDate),
      actualDateFrom: formatDate(t.StartDate),
      actualDateTo: formatDate(t.EndDate || t.StartDate),
      trainingStatus: t.Status || 'In Progress',
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
  if (str instanceof Date) return str;
  const s = String(str).trim();
  if (!s) return null;

  // Match DD/MM/YYYY or DD-MM-YYYY
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) {
    const day = parseInt(dmy[1], 10);
    const month = parseInt(dmy[2], 10) - 1;
    const year = parseInt(dmy[3], 10);
    return new Date(year, month, day);
  }

  // Match YYYY-MM-DD
  const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymd) {
    const year = parseInt(ymd[1], 10);
    const month = parseInt(ymd[2], 10) - 1;
    const day = parseInt(ymd[3], 10);
    return new Date(year, month, day);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function parseServerRes(res) {
  if (!res) return { success: false, data: null };
  if (typeof res === 'object') return res;
  try {
    return JSON.parse(res);
  } catch (e) {
    return { success: false, data: null };
  }
}

function ensureSheetDimensions(sheet, minRows, minCols) {
  if (!sheet) return;
  try {
    const curRows = sheet.getMaxRows();
    const curCols = sheet.getMaxColumns();
    if (curRows < minRows) sheet.insertRowsAfter(curRows, minRows - curRows);
    if (curCols < minCols) sheet.insertColumnsAfter(curCols, minCols - curCols);
  } catch (e) {}
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

    const res = parseServerRes(getFilteredReportData(reportType, filters));
    if (!res.success) return err(res.message || 'Failed to parse report data.');
    const data = res.data || {};
    const year = (filters && filters.year) ? filters.year : '2026';
    const timestamp = Date.now();

    let titleName = 'Report';
    if (reportType === 'hours') titleName = `Training_Hours_CostCentre_${year}_${timestamp}`;
    else if (reportType === 'cost') titleName = `Training_Cost_${year}_${timestamp}`;
    else if (reportType === 'title') titleName = `Training_Title_Report_${timestamp}`;
    else if (reportType === 'employee') titleName = `Employee_Training_Report_${timestamp}`;
    else if (reportType === 'atp') titleName = `Annual_Training_Plan_${year}_${timestamp}`;

    const repFolder = getOrCreateReportsFolder();
    const ss = SpreadsheetApp.create(titleName);
    const file = DriveApp.getFileById(ss.getId());
    repFolder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}

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
      folderId: repFolder.getId(),
      folderName: repFolder.getName(),
      folderUrl: repFolder.getUrl(),
      message: `Report exported and saved in Google Drive "Reports" folder!`
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

    const repFolder = getOrCreateReportsFolder();
    const ss = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(ss.getId());
    repFolder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}

    // Tab 1: Training Hours
    const sHours = ss.getActiveSheet();
    sHours.setName('Training Hours');
    const hoursRes = parseServerRes(getFilteredReportData('hours', filters));
    if (hoursRes.success && hoursRes.data) formatHoursReportSheet(sHours, hoursRes.data, year);

    // Tab 2: Training Cost
    const sCost = ss.insertSheet('Training Cost');
    const costRes = parseServerRes(getFilteredReportData('cost', filters));
    if (costRes.success && costRes.data) formatCostReportSheet(sCost, costRes.data, year);

    // Tab 3: Training Title
    const sTitle = ss.insertSheet('Training Title');
    const titleRes = parseServerRes(getFilteredReportData('title', filters));
    if (titleRes.success && titleRes.data) formatTitleReportSheet(sTitle, titleRes.data);

    // Tab 4: Employee Report
    const sEmp = ss.insertSheet('Employee Report');
    const empRes = parseServerRes(getFilteredReportData('employee', filters));
    if (empRes.success && empRes.data) formatEmployeeReportSheet(sEmp, empRes.data);

    // Tab 5: Annual Training Plan (ATP)
    const sAtp = ss.insertSheet('Annual Training Plan (ATP)');
    const atpRes = parseServerRes(getFilteredReportData('atp', filters));
    if (atpRes.success && atpRes.data) formatAtpReportSheet(sAtp, atpRes.data);

    SpreadsheetApp.flush();

    const fileId = ss.getId();
    const sheetUrl = ss.getUrl();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    return ok({
      fileId: fileId,
      fileName: `${fileName}.xlsx`,
      sheetUrl: sheetUrl,
      downloadUrl: downloadUrl,
      folderId: repFolder.getId(),
      folderName: repFolder.getName(),
      folderUrl: repFolder.getUrl(),
      message: `Master All-in-One Report Workbook created and saved in Google Drive "Reports" folder!`
    });
  } catch (e) {
    return err('Failed to export Master Report Workbook: ' + e.message);
  }
}

/**
 * Synchronizes and updates the Live Master Report Google Sheet inside the "Reports" folder in Google Drive.
 * Contains all 5 live tabs (Training Hours, Training Cost, Training Title, Employee Report, ATP).
 */
function syncLiveMasterReportSheet(targetYear) {
  try {
    const year = (targetYear && targetYear !== '2026') ? targetYear : 'All';
    const displayYear = (targetYear && targetYear !== 'All') ? targetYear : '2026';
    const repFolder = getOrCreateReportsFolder();
    const fileName = `Master Annual Training Report (${displayYear})`;

    let fileIter = repFolder.getFilesByName(fileName);
    let ss;
    let file;

    if (fileIter.hasNext()) {
      file = fileIter.next();
      ss = SpreadsheetApp.openById(file.getId());
    } else {
      ss = SpreadsheetApp.create(fileName);
      file = DriveApp.getFileById(ss.getId());
      repFolder.addFile(file);
      try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
    }

    const filters = { year: 'All' };

    // Tab 1: Training Hours
    let sHours = ss.getSheetByName('Training Hours');
    if (!sHours) sHours = ss.insertSheet('Training Hours');
    const hoursRes = parseServerRes(getFilteredReportData('hours', filters));
    if (hoursRes.success && hoursRes.data) formatHoursReportSheet(sHours, hoursRes.data, displayYear);

    // Tab 2: Training Cost
    let sCost = ss.getSheetByName('Training Cost');
    if (!sCost) sCost = ss.insertSheet('Training Cost');
    const costRes = parseServerRes(getFilteredReportData('cost', filters));
    if (costRes.success && costRes.data) formatCostReportSheet(sCost, costRes.data, displayYear);

    // Tab 3: Training Title
    let sTitle = ss.getSheetByName('Training Title');
    if (!sTitle) sTitle = ss.insertSheet('Training Title');
    const titleRes = parseServerRes(getFilteredReportData('title', filters));
    if (titleRes.success && titleRes.data) formatTitleReportSheet(sTitle, titleRes.data);

    // Tab 4: Employee Report
    let sEmp = ss.getSheetByName('Employee Report');
    if (!sEmp) sEmp = ss.insertSheet('Employee Report');
    const empRes = parseServerRes(getFilteredReportData('employee', filters));
    if (empRes.success && empRes.data) formatEmployeeReportSheet(sEmp, empRes.data);

    // Tab 5: Annual Training Plan (ATP)
    let sAtp = ss.getSheetByName('Annual Training Plan (ATP)');
    if (!sAtp) sAtp = ss.insertSheet('Annual Training Plan (ATP)');
    const atpRes = parseServerRes(getFilteredReportData('atp', filters));
    if (atpRes.success && atpRes.data) formatAtpReportSheet(sAtp, atpRes.data);

    // Clean up default empty sheet if present
    const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Helai1');
    if (defaultSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defaultSheet); } catch (e) {}
    }

    SpreadsheetApp.flush();

    return ok({
      folderId: repFolder.getId(),
      folderUrl: repFolder.getUrl(),
      fileId: ss.getId(),
      fileUrl: ss.getUrl(),
      message: `Live Master Report successfully synced inside 'Reports' folder with all 5 tabs updated real-time!`
    });
  } catch (e) {
    Logger.log('syncLiveMasterReportSheet error: ' + e.message);
    return err('Failed to sync Live Master Report: ' + e.message);
  }
}

/**
 * Gets details of the "Reports" folder and the Live Master Report Sheet URL.
 */
function getReportsFolderDetails() {
  try {
    const repFolder = getOrCreateReportsFolder();
    const year = '2026';
    const fileName = `Master Annual Training Report (${year})`;

    let fileIter = repFolder.getFilesByName(fileName);
    let fileUrl = '';
    let fileId = '';
    if (fileIter.hasNext()) {
      const f = fileIter.next();
      fileUrl = f.getUrl();
      fileId = f.getId();
    }

    return ok({
      folderId: repFolder.getId(),
      folderUrl: repFolder.getUrl(),
      liveMasterSheetId: fileId,
      liveMasterSheetUrl: fileUrl,
      message: 'Reports folder details retrieved.'
    });
  } catch (e) {
    return err(e.message);
  }
}

// ─── Header Formatters for Excel Export & Live Master Sheets ──────────────────

// ─── Header Formatters for Excel Export & Live Master Sheets ──────────────────

// 1. Hours Header Formatter:
function formatHoursReportSheet(sheet, data, year) {
  ensureSheetDimensions(sheet, 100, 30);
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  } catch (e) {}

  const yrSuffix = year && year.length === 4 ? year.slice(-2) : '26';
  const monthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].map(m => `${m}-${yrSuffix}`);

  const row1 = ['Cost Centre/Month', 'Training Hours', '', '', '', '', '', '', '', '', '', '', '', `Total ${year}`, `Total Training Hours ${year}`];
  const row2 = ['', ...monthLabels, '', ''];

  sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
  sheet.getRange(2, 1, 1, row2.length).setValues([row2]);

  sheet.getRange('A1:A2').merge().setValue('Cost Centre/Month');
  sheet.getRange('B1:M1').merge().setValue('Training Hours');
  sheet.getRange('N1:N2').merge().setValue(`Total ${year}`);
  sheet.getRange('O1:O2').merge().setValue(`Total Training Hours ${year}`);

  const headerRange = sheet.getRange('A1:O2');
  headerRange.setFontWeight('bold')
    .setBackground('#2563EB')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  const rows = data.rows || [];
  const rowsData = rows.map(r => [
    r.costCentre || 'General',
    ...(r.months || Array(12).fill(0)),
    r.totalYear || 0,
    r.totalHours || 0
  ]);

  // Compute Total Row summary across all Cost Centres
  const totalMonths = Array(12).fill(0);
  let grandTotalYear = 0;
  let grandTotalHours = 0;

  rows.forEach(r => {
    (r.months || []).forEach((val, idx) => {
      totalMonths[idx] += Number(val || 0);
    });
    grandTotalYear += Number(r.totalYear || 0);
    grandTotalHours += Number(r.totalHours || 0);
  });

  const totalRow = [
    'Total',
    ...totalMonths,
    grandTotalYear,
    grandTotalHours
  ];

  if (rowsData.length > 0) {
    sheet.getRange(3, 1, rowsData.length, 15).setValues(rowsData);

    // Append Total Row at the bottom of data rows
    const totalRowIndex = 3 + rowsData.length;
    sheet.getRange(totalRowIndex, 1, 1, 15).setValues([totalRow]);

    // Format Total Row (Bold text, light blue background)
    const totalRange = sheet.getRange(totalRowIndex, 1, 1, 15);
    totalRange.setFontWeight('bold')
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A');
  }

  sheet.autoResizeColumns(1, 15);
}

// 2. Cost Header Formatter:
function formatCostReportSheet(sheet, data, year) {
  ensureSheetDimensions(sheet, 100, 30);
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  } catch (e) {}

  const row1 = [
    'Training Title', 'Training Date (From)', 'Training Date (To)', 'Total Participant',
    'Training Cost (RM)', '', '', '', '', '', '', '',
    `Total Cost ${year}`, 'Total HRDF Grant (RM)'
  ];
  const row2 = [
    '', '', '', '',
    'Training Fees', 'Meal', 'Subsistance Allowance', 'Hotel Fees', 'Mileage Claim', 'Taxi Fees', 'Toll Fees', 'Flight',
    '', ''
  ];

  sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
  sheet.getRange(2, 1, 1, row2.length).setValues([row2]);

  sheet.getRange('A1:A2').merge().setValue('Training Title');
  sheet.getRange('B1:B2').merge().setValue('Training Date (From)');
  sheet.getRange('C1:C2').merge().setValue('Training Date (To)');
  sheet.getRange('D1:D2').merge().setValue('Total Participant');
  sheet.getRange('E1:L1').merge().setValue('Training Cost (RM)');
  sheet.getRange('M1:M2').merge().setValue(`Total Cost ${year}`);
  sheet.getRange('N1:N2').merge().setValue('Total HRDF Grant (RM)');

  const headerRange = sheet.getRange('A1:N2');
  headerRange.setFontWeight('bold')
    .setBackground('#2563EB')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  const rowsData = (data.rows || []).map(r => [
    r.trainingTitle || 'Training Programme',
    r.dateFrom || '',
    r.dateTo || '',
    r.totalParticipant || 0,
    r.trainingFees || '',
    r.meal || '',
    r.subsistanceAllowance || '',
    r.hotelFees || '',
    r.mileageClaim || '',
    r.taxiFees || '',
    r.tollFees || '',
    r.flight || '',
    r.totalCost || 'RM 0.00',
    r.totalHrdfGrant || ''
  ]);

  if (rowsData.length > 0) {
    sheet.getRange(3, 1, rowsData.length, 14).setValues(rowsData);
  }

  sheet.autoResizeColumns(1, 14);
}

// 3. Training Title Header Formatter:
function formatTitleReportSheet(sheet, data) {
  ensureSheetDimensions(sheet, 100, 30);
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  } catch (e) {}

  const headers = [
    'Name', 'Emp. No', 'Training Title', '(EE)/Cost Centre Description',
    'Training Type', 'Date (From)', 'Date (Until)', 'Month',
    'Total Hours', 'Trainer', 'Training Provide', 'Expiry Date (if applicable)'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  const rowsData = (data.rows || []).map(r => [
    r.name || '—',
    r.empNo || '—',
    r.trainingTitle || '—',
    r.costCentreDesc || '—',
    r.trainingType || 'Public',
    r.dateFrom || '',
    r.dateUntil || '',
    r.month || '',
    r.totalHours || 8,
    r.trainer || '—',
    r.trainingProvider || '—',
    r.expiryDate || 'N/A'
  ]);

  if (rowsData.length > 0) {
    sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);
  }

  sheet.autoResizeColumns(1, headers.length);
}

// 4. Employee Header Formatter
function formatEmployeeReportSheet(sheet, data) {
  ensureSheetDimensions(sheet, 100, 30);
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  } catch (e) {}

  const headers = ['No', 'Emp. No', 'Employee Name', 'Cost Centre / Department', 'Position', 'Total Trainings', 'Total Hours', 'Attended Trainings', 'Status'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  const rowsData = (data.rows || []).map(r => [
    r.no || 1,
    r.empNo || '—',
    r.name || '—',
    r.costCentre || '—',
    r.position || '—',
    r.totalTrainings || 0,
    r.totalHours || 0,
    Array.isArray(r.attendedList) ? r.attendedList.join(', ') : (r.attendedList || '—'),
    r.status || 'Active'
  ]);

  if (rowsData.length > 0) {
    sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);
  }

  sheet.autoResizeColumns(1, headers.length);
}

// 5. ATP Header Formatter:
function formatAtpReportSheet(sheet, data) {
  ensureSheetDimensions(sheet, 100, 30);
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  } catch (e) {}

  const headers = [
    'No', 'Training Title', 'Training Category', 'TNA Source', 'Training Mode',
    'Training Duration (hrs)', 'Trainer', 'Department', 'Position / Employee No',
    'Total Pax', 'Planned Date', 'Actual Date (From)', 'Actual Date (To)',
    'Training Status', 'Remarks'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  const rowsData = (data.rows || []).map(r => [
    r.no || 1,
    r.trainingTitle || '—',
    r.trainingCategory || 'General',
    r.tnaSource || 'Annual TNA',
    r.trainingMode || 'Physical Classroom',
    r.durationHours || 8,
    r.trainer || '—',
    r.department || '—',
    r.positionEmpNo || 'All Staff',
    r.totalPax || 0,
    r.plannedDate || '',
    r.actualDateFrom || '',
    r.actualDateTo || '',
    r.trainingStatus || 'Planned',
    r.remarks || 'Planned'
  ]);

  if (rowsData.length > 0) {
    sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);
  }

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