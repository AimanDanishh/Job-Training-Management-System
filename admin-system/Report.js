/**
 * Report.gs - Comprehensive Report Generation & Excel Export Engine
 */

function safeParseObj(val) {
  if (!val) return {};
  if (typeof val === 'object') return val;
  try {
    return JSON.parse(val);
  } catch(e) {
    return {};
  }
}

// --- Full Programme Report (Single Training) ----------------------------------
function generateReport(trainingId) {
  try {
    const tResult  = safeParseObj(getTrainingById(trainingId));
    if (!tResult.success) return err('Training not found.');
    const training = tResult.data;

    const attSummary  = safeParseObj(getAttendanceSummary(trainingId)).data  || {};
    const evalSummary = safeParseObj(getEvaluationSummary(trainingId)).data  || {};
    const attData     = safeParseObj(getAttendance(trainingId)).data         || [];
    const evalData    = safeParseObj(getTrainingEvaluations(trainingId)).data || [];
    const postData    = safeParseObj(getPostEvaluations(trainingId)).data    || [];

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

// --- Dashboard Overview Report --------------------------------------------------
function getDashboardReport() {
  try {
    const trainingSummary = safeParseObj(getTrainingSummary()).data || {};
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
    let returnedCount = 0;

    tRows.forEach(r => {
      const appStatus = String(r.ApprovalStatus || '').trim();
      const appStatusLower = appStatus.toLowerCase();
      if (appStatus === 'Pending HOD Approval') pendingHodCount++;
      else if (appStatus === 'Pending C-Suite Approval') pendingCsuiteCount++;
      else if (appStatus === 'Pending HOHR Approval') pendingHohrCount++;
      else if (appStatus === 'Rejected' || appStatusLower.includes('reject')) rejectedCount++;
      else if (appStatusLower.includes('return')) returnedCount++;
      else if (appStatusLower === 'approved' || appStatusLower === 'auto-approved' || appStatusLower === 'fully approved' || (!appStatus && String(r.Status || '').toLowerCase().includes('approved'))) {
        approvedCount++;
      }
    });

    const approvalSummary = {
      pendingHod:    pendingHodCount,
      pendingCsuite: pendingCsuiteCount,
      pendingHohr:   pendingHohrCount,
      approved:      approvedCount,
      rejected:      rejectedCount,
      returned:      returnedCount,
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

// --- Unified Filtered Reports Engine ------------------------------------------
/**
 * Generates structured report data based on report type and filter parameters.
 * Types: 'hours', 'cost', 'title', 'employee', 'atp', 'single'
 */
function getFilteredReportData(reportType, filters) {
  try {
    const opts = filters || {};
    const selectedYear   = String(opts.year || '').trim();
    const selectedMonth  = String(opts.month || '').trim();
    const selectedStatus = String(opts.status || '').trim().toLowerCase();
    const selectedTitle  = String(opts.title || '').trim().toLowerCase();
    const selectedDept   = String(opts.costCentre || opts.department || '').trim().toLowerCase();
    const selectedCat    = String(opts.category || '').trim().toLowerCase();
    const selectedMode   = String(opts.mode || '').trim().toLowerCase();

    const tSheet = getSheet(SHEET_NAMES.trainings);
    const tRows  = tSheet ? sheetToJson(tSheet) : [];
    const empSheet = getSheet(SHEET_NAMES.employees);
    const empRows = empSheet ? sheetToJson(empSheet) : [];

    // Filter trainings by year, month/quarter, status, title, category, mode, department
    const filteredTrainings = tRows.filter(t => {
      const startDateStr = t.StartDate || '';
      let matchYear = true;
      let matchMonth = true;

      if (startDateStr) {
        const d = parseDateObj(startDateStr);
        if (d) {
          const yStr = String(d.getFullYear());
          const mNum = d.getMonth() + 1;
          const mNumStr = String(mNum).padStart(2, '0');
          const mNameShort = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM');

          if (selectedYear && selectedYear !== 'All') {
            matchYear = (yStr === selectedYear);
          }

          if (selectedMonth && selectedMonth !== 'All') {
            if (selectedMonth === 'Q1') matchMonth = (mNum >= 1 && mNum <= 3);
            else if (selectedMonth === 'Q2') matchMonth = (mNum >= 4 && mNum <= 6);
            else if (selectedMonth === 'Q3') matchMonth = (mNum >= 7 && mNum <= 9);
            else if (selectedMonth === 'Q4') matchMonth = (mNum >= 10 && mNum <= 12);
            else matchMonth = (mNumStr === selectedMonth || mNameShort.toLowerCase() === selectedMonth.toLowerCase() || String(mNum) === selectedMonth);
          }
        }
      }

      const matchTitle = !selectedTitle || String(t.Name || t.ID || t.Code || '').toLowerCase().includes(selectedTitle);
      const matchDept  = !selectedDept  || String(t.Department || '').toLowerCase().includes(selectedDept);
      const matchCat   = !selectedCat   || String(t.Category || '').toLowerCase().includes(selectedCat);
      const matchMode  = !selectedMode  || String(t.TrainingMode || '').toLowerCase().includes(selectedMode);

      const canonical = getCanonicalTrainingStatus(t);
      let matchStatus = true;
      if (selectedStatus) {
        const targetQ = selectedStatus.toLowerCase();
        if (targetQ === 'pending') matchStatus = (canonical === 'Pending Approval');
        else if (targetQ === 'approved') matchStatus = (canonical === 'Approved');
        else if (targetQ === 'ongoing') matchStatus = (canonical === 'Ongoing');
        else if (targetQ === 'completed') matchStatus = (canonical === 'Completed');
        else if (targetQ === 'rejected') matchStatus = (canonical === 'Rejected');
        else matchStatus = canonical.toLowerCase().includes(targetQ);
      }

      return matchYear && matchMonth && matchTitle && matchDept && matchCat && matchMode && matchStatus;
    });

    const currentYr = String(new Date().getFullYear());
    if (reportType === 'hours') {
      return ok(buildHoursReportData(filteredTrainings, empRows, selectedYear || currentYr, selectedDept));
    } else if (reportType === 'cost') {
      return ok(buildCostReportData(filteredTrainings, selectedYear || currentYr, selectedDept));
    } else if (reportType === 'title') {
      return ok(buildTrainingTitleReportData(filteredTrainings, empRows, selectedDept));
    } else if (reportType === 'employee') {
      return ok(buildEmployeeReportData(filteredTrainings, empRows, selectedDept));
    } else if (reportType === 'atp') {
      return ok(buildAnnualTrainingPlanData(filteredTrainings, selectedDept));
    } else {
      return err('Invalid report type specified.');
    }
  } catch (e) {
    return err('Failed to generate filtered report: ' + e.message);
  }
}

// --- Numeric Cleaner Helper ---------------------------------------------------
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


// --- 1. Training Hours Report Data Builder (Cost Centre vs Month) -------------
function buildHoursReportData(trainingsInput, employeesInput, year, selectedDept) {
  let trainings = (Array.isArray(trainingsInput)) ? trainingsInput : [];
  let employees = (Array.isArray(employeesInput)) ? employeesInput : [];

  const yrSuffix = year && year.length === 4 ? year.slice(-2) : String(new Date().getFullYear()).slice(-2);
  const monthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].map(m => `${m}-${yrSuffix}`);

  // Fetch official Cost Centres strictly from "Cost Centre" tab under Employee Spreadsheet
  const ccMap = new Map(); // key: lowercased, value: canonical Cost Centre display name
  try {
    const ccRes = parseServerRes(getCostCentres());
    if (ccRes.success && Array.isArray(ccRes.data)) {
      ccRes.data.forEach(ccName => {
        const str = String(ccName || '').trim();
        if (str && str !== 'All Departments / Cost Centres') {
          const key = str.toLowerCase();
          if (!ccMap.has(key)) ccMap.set(key, str);
        }
      });
    }
  } catch (err) {}

  let costCentreList = Array.from(ccMap.values());
  if (selectedDept) {
    const fKey = String(selectedDept).trim().toLowerCase();
    const filteredCC = costCentreList.filter(cc => cc.toLowerCase().includes(fKey) || fKey.includes(cc.toLowerCase()));
    if (filteredCC.length > 0) {
      costCentreList = filteredCC;
    }
  }

  const matrix = {};

  costCentreList.forEach(cc => {
    matrix[cc] = Array(12).fill(0);
  });

  // Helper function to match any incoming raw cost centre to official Cost Centre list (case-insensitive)
  function matchOfficialCostCentre(rawCC) {
    if (!rawCC) return costCentreList[0];
    const key = String(rawCC).trim().toLowerCase();
    if (ccMap.has(key)) return ccMap.get(key);
    for (let [k, canonical] of ccMap.entries()) {
      if (k.includes(key) || key.includes(k)) return canonical;
    }
    return costCentreList[0];
  }

  // Calculate training hours per cost centre per month
  trainings.forEach(t => {
    const startDate = parseDateObj(t.StartDate);
    const monthIdx = startDate ? startDate.getMonth() : 1; // Default to Feb if unparsed
    const hours = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 8);

    const participants = getTrainingParticipantsList(t.ID || t.Code);
    if (participants.length > 0) {
      participants.forEach(p => {
        const rawCC = String(p.Department || p.CostCentre || t.Department || '').trim();
        const cc = matchOfficialCostCentre(rawCC);
        if (!matrix[cc]) matrix[cc] = Array(12).fill(0);
        matrix[cc][monthIdx] += hours;
      });
    } else {
      const rawCC = String(t.Department || '').trim();
      const cc = matchOfficialCostCentre(rawCC);
      if (!matrix[cc]) matrix[cc] = Array(12).fill(0);
      matrix[cc][monthIdx] += hours;
    }
  });

  const rows = [];
  costCentreList.forEach(cc => {
    const mRaw = matrix[cc] || Array(12).fill(0);
    const totalYear = mRaw.reduce((a, b) => a + Number(b || 0), 0);
    const mData = mRaw.map(v => (v === 0 || !v) ? '' : v);

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

// --- 2. Training Cost Report Data Builder -------------------------------------
function buildCostReportData(trainingsInput, year, selectedDept) {
  let trainings = (Array.isArray(trainingsInput)) ? trainingsInput : [];
  if (selectedDept) {
    const fKey = String(selectedDept).trim().toLowerCase();
    trainings = trainings.filter(t => String(t.Department || t.CostCentre || '').toLowerCase().includes(fKey));
  }

  const rows = trainings.map(t => {
    const feeNum = cleanNum(t.CourseFee || t['Course Fee'], 0);
    const paxNum = cleanNum(t.Participants || t['Total Participant'] || t.Pax, 0);

    return {
      trainingTitle: t.Name || t['Training Name'] || t.Code || 'Training Programme',
      dateFrom: formatOrdinalDate(t.StartDate),
      dateTo: formatOrdinalDate(t.EndDate || t.StartDate),
      totalParticipant: paxNum > 0 ? paxNum : 0,
      trainingFees: feeNum > 0 ? feeNum : '',
      meal: '',
      subsistanceAllowance: '',
      hotelFees: '',
      mileageClaim: '',
      taxiFees: '',
      tollFees: '',
      flight: '',
      totalCost: feeNum > 0 ? feeNum : '',
      totalHrdfGrant: ''
    };
  });

  return {
    year: year,
    rows: rows
  };
}

function formatMonthYear(dateStrOrObj) {
  if (!dateStrOrObj) return '-';
  const d = parseDateObj(dateStrOrObj);
  if (!d || isNaN(d.getTime())) return String(dateStrOrObj).toUpperCase();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const mStr = months[d.getMonth()];
  const yStr = d.getFullYear();
  return `${mStr}-${yStr}`;
}

// --- 3. Training Title Report Data Builder ------------------------------------
function buildTrainingTitleReportData(trainingsInput, employeesInput, selectedDept) {
  let trainings = (Array.isArray(trainingsInput)) ? trainingsInput : [];
  let employees = (Array.isArray(employeesInput)) ? employeesInput : [];

  // Build employee lookup dictionary by ID and Name
  const empMap = {};
  employees.forEach(e => {
    const id = String(e.ID || e.EmployeeID || e.EmployeeNo || '').trim();
    const name = String(e.Name || e.EmployeeName || '').trim().toLowerCase();
    if (id) empMap[id.toLowerCase()] = e;
    if (name) empMap[name] = e;
  });

  let rows = [];

  trainings.forEach(t => {
    const participants = getTrainingParticipantsList(t.ID || t.Code);
    const monthFormatted = formatMonthYear(t.StartDate);
    const totalHours = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 0);
    const tTitle = t.Name || t['Training Name'] || t.Code || 'Training Programme';
    const tType = formatTrainingMode(t.TrainingMode || t.Mode || t.Type || t.Category || t.Venue);

    if (participants.length > 0) {
      participants.forEach(p => {
        const pEmpId = String(p.EmployeeID || p.EmployeeNo || p.EmpNo || p.EmpID || '').trim();
        const pName = String(p.EmployeeName || p.Name || '').trim();

        // Cross reference with Employee Directory if available
        let matchedEmp = (pEmpId && empMap[pEmpId.toLowerCase()]) || (pName && empMap[pName.toLowerCase()]);

        const resolvedEmpNo = (matchedEmp ? (matchedEmp.EmployeeID || matchedEmp.ID || matchedEmp.EmployeeNo) : null)
          || pEmpId
          || (p.ID && !String(p.ID).toUpperCase().startsWith('TP-') ? p.ID : '')
          || '-';

        const resolvedName = (matchedEmp ? matchedEmp.Name : null) || pName || 'Staff Member';
        const resolvedDept = (matchedEmp ? (matchedEmp.CostCentre || matchedEmp.Department) : null) || p.Department || t.Department || '-';

        rows.push({
          name: resolvedName,
          empNo: resolvedEmpNo,
          trainingTitle: tTitle,
          costCentreDesc: resolvedDept,
          trainingType: tType,
          dateFrom: formatOrdinalDate(t.StartDate),
          dateUntil: formatOrdinalDate(t.EndDate || t.StartDate),
          month: monthFormatted,
          totalHours: totalHours,
          trainer: t.Trainer || '-',
          trainingProvider: t.Trainer || '-',
          expiryDate: 'N/A'
        });
      });
    }
  });

  if (selectedDept) {
    const fKey = String(selectedDept).trim().toLowerCase();
    rows = rows.filter(r => String(r.costCentreDesc || '').toLowerCase().includes(fKey));
  }

  return { rows: rows };
}

// --- 4. Employee Report Data Builder ------------------------------------------
function buildEmployeeReportData(trainingsInput, employeesInput, filterDept) {
  let trainings = (Array.isArray(trainingsInput)) ? trainingsInput : [];
  let employees = (Array.isArray(employeesInput)) ? employeesInput : [];

  const empMap = {};

  employees.forEach((e, idx) => {
    const id = String(e.ID || e.EmployeeID || e.EmployeeNo || '').trim();
    if (!id) return;
    const cleanId = id.toLowerCase();
    const cleanName = String(e.Name || e.EmployeeName || id).trim().toLowerCase();

    empMap[cleanId] = {
      no: idx + 1,
      empNo: id,
      name: e.Name || e.EmployeeName || id,
      costCentre: e.CostCentre || e['Cost Centre'] || e.Department || '-',
      position: e.Position || e.JobTitle || e.PositionTitle || '-',
      totalTrainings: 0,
      totalHours: 0,
      attendedList: [],
      status: e.Status || 'Active',
      cleanName: cleanName
    };
  });

  trainings.forEach((t, tIdx) => {
    const tTitle = t.Name || t['Training Name'] || t.Code || `Training ${tIdx+1}`;
    const hours = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 0);
    const tDept = String(t.Department || '').trim().toLowerCase();

    const participants = getTrainingParticipantsList(t.ID || t.Code);

    let attRecords = [];
    try {
      const attRes = parseServerRes(getAttendance(t.ID || t.Code));
      if (attRes.success && Array.isArray(attRes.data)) {
        attRes.data.forEach(dayObj => {
          if (Array.isArray(dayObj.records)) {
            attRecords.push(...dayObj.records);
          }
        });
      }
    } catch (err) {}

    const matchedEmpKeys = new Set();
    const allRecords = [...participants, ...attRecords];

    if (allRecords.length > 0) {
      allRecords.forEach(p => {
        const pEmpId = String(p.EmployeeID || p.EmployeeNo || p.EmpNo || p.EmpID || p.ID || '').trim().toLowerCase();
        const pName  = String(p.EmployeeName || p.Name || p.EmpName || '').trim().toLowerCase();

        let targetKey = null;
        if (pEmpId && empMap[pEmpId]) {
          targetKey = pEmpId;
        } else if (pName) {
          Object.keys(empMap).forEach(k => {
            if (empMap[k].cleanName === pName || (pName.length > 3 && empMap[k].cleanName.includes(pName))) {
              targetKey = k;
            }
          });
        }

        if (targetKey && !matchedEmpKeys.has(targetKey)) {
          matchedEmpKeys.add(targetKey);
        }
      });
    }

    matchedEmpKeys.forEach(empKey => {
      const empObj = empMap[empKey];
      if (empObj) {
        if (!empObj.attendedList.includes(tTitle)) {
          empObj.attendedList.push(tTitle);
          empObj.totalTrainings = empObj.attendedList.length;
          empObj.totalHours += hours;
        }
      }
    });
  });

  let list = Object.values(empMap).map(e => {
    delete e.cleanName;
    return e;
  });

  if (filterDept) {
    const fDept = filterDept.toLowerCase();
    list = list.filter(e => e.costCentre.toLowerCase().includes(fDept));
  }

  return { rows: list };
}

// --- ATP Helper Functions --------------------------------------------------
function getOrdinalSuffix(day) {
  const d = parseInt(day, 10);
  if (d >= 11 && d <= 13) return 'th';
  switch (d % 10) {
    case 1:  return 'st';
    case 2:  return 'nd';
    case 3:  return 'rd';
    default: return 'th';
  }
}

function getOrdinalSuffix(day) {
  const d = parseInt(day, 10);
  if (isNaN(d)) return 'th';
  if (d > 3 && d < 21) return 'th';
  switch (d % 10) {
    case 1:  return 'st';
    case 2:  return 'nd';
    case 3:  return 'rd';
    default: return 'th';
  }
}

function formatOrdinalDate(dateStrOrObj) {
  if (!dateStrOrObj) return '-';
  const d = parseDateObj(dateStrOrObj);
  if (!d || isNaN(d.getTime())) return String(dateStrOrObj);

  const dayNum = d.getDate();
  const suffix = getOrdinalSuffix(dayNum);
  const fullMonths = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  const monthStr = fullMonths[d.getMonth()];
  const yearStr = d.getFullYear();

  return `${dayNum}${suffix} ${monthStr} ${yearStr}`;
}

function formatTrainingMode(val) {
  if (!val) return 'In-house';
  const s = String(val).trim().toLowerCase();
  if (s.includes('public') || s.includes('external') || s.includes('vendor')) {
    return 'Public';
  }
  return 'In-house';
}

function getCanonicalTrainingStatus(t) {
  if (!t) return 'Pending Approval';
  if (typeof t === 'string') {
    const s = t.toLowerCase();
    if (s.includes('reject')) return 'Rejected';
    if (s.includes('return')) return 'Returned';
    if (s.includes('pending') || s.includes('review') || s.includes('draft') || s.includes('plan')) return 'Pending Approval';
    if (s.includes('completed') || s.includes('complete') || s.includes('closed')) return 'Completed';
    if (s.includes('progress') || s.includes('ongoing') || s.includes('going') || s.includes('active') || s.includes('attendance')) return 'Ongoing';
    if (s.includes('approved') || s.includes('acknowledged') || s.includes('upcoming')) return 'Approved';
    return 'Pending Approval';
  }

  const appStatus = String(t.ApprovalStatus || t.Status || '').trim().toLowerCase();
  const statusStr = String(t.Status || '').trim().toLowerCase();
  const stageStr = String(t.Stage || '').trim().toLowerCase();

  // 1. Rejected
  if (appStatus.includes('reject') || statusStr.includes('reject')) {
    return 'Rejected';
  }

  // 1b. Returned
  if (appStatus.includes('return') || statusStr.includes('return') || stageStr.includes('return')) {
    return 'Returned';
  }

  // 2. Pending Approval
  if (appStatus.includes('pending') || appStatus.includes('review') || statusStr.includes('draft') || statusStr.includes('planning') || statusStr.includes('plan')) {
    return 'Pending Approval';
  }

  // 3. Completed
  if (statusStr.includes('completed') || statusStr.includes('complete') || stageStr.includes('completed') || stageStr.includes('closed') || stageStr.includes('review')) {
    return 'Completed';
  }

  // 4. Ongoing / Active
  if (statusStr.includes('progress') || statusStr.includes('ongoing') || statusStr.includes('going') || statusStr.includes('active') || stageStr.includes('attendance')) {
    return 'Ongoing';
  }

  // 5. Approved
  if (appStatus.includes('approved') || appStatus.includes('acknowledged') || statusStr.includes('approved') || statusStr.includes('upcoming')) {
    return 'Approved';
  }

  if (stageStr.includes('created') || stageStr.includes('imported')) {
    return 'Pending Approval';
  }

  return 'Pending Approval';
}

function normalizeAtpStatus(statusStr, stageStr, appStatusStr) {
  const canonical = getCanonicalTrainingStatus({ Status: statusStr, Stage: stageStr, ApprovalStatus: appStatusStr });
  if (canonical === 'Completed') return 'Completed';
  if (canonical === 'Ongoing') return 'On going';
  if (canonical === 'Pending Approval') return 'Pending';
  return 'On planning';
}

// --- 5. Annual Training Plan (ATP) Data Builder -------------------------------
function buildAnnualTrainingPlanData(trainingsInput, selectedDept) {
  let trainings = (Array.isArray(trainingsInput)) ? trainingsInput : [];
  if (selectedDept) {
    const fKey = String(selectedDept).trim().toLowerCase();
    trainings = trainings.filter(t => String(t.Department || t.CostCentre || '').toLowerCase().includes(fKey));
  }

  const rows = trainings.map((t, idx) => {
    const duration = cleanNum(t.TotalHours || (cleanNum(t.Duration, 1) * 8), 0);
    const pax = cleanNum(t.Participants, 0);

    let reqUrl = t.requisitionFormUrl || t.RequisitionUrl || t.BrochureURL || '';
    if (!reqUrl && t.RequisitionFormFileID) {
      reqUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(t.RequisitionFormFileID)}/edit`;
    }

    let dataSheetUrl = t.trainingDataSheetUrl || '';
    if (!dataSheetUrl && t.ParticipantsSheetID) {
      dataSheetUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(t.ParticipantsSheetID)}/edit`;
    }

    return {
      no: idx + 1,
      trainingTitle: t.Name || t['Training Name'] || `Training ${idx+1}`,
      trainingCategory: t.Category || 'Compliance',
      tnaSource: t.TnaSource || t['TNA Source'] || 'Training Requisition Form',
      trainingMode: formatTrainingMode(t.TrainingMode || t.Mode || t.Venue),
      durationHours: duration,
      trainer: t.Trainer || 'Certified Trainer',
      department: t.Department || 'All Departments',
      positionEmpNo: 'Name List',
      requisitionFormUrl: reqUrl,
      trainingDataSheetUrl: dataSheetUrl,
      totalPax: pax,
      plannedDate: formatOrdinalDate(t.StartDate),
      actualDateFrom: formatOrdinalDate(t.StartDate),
      actualDateTo: formatOrdinalDate(t.EndDate || t.StartDate),
      trainingStatus: normalizeAtpStatus(t.Status, t.Stage),
      remarks: t.Stage || 'Planned',
      trainingId: t.ID || t.Code || `TRN-${idx+1}`
    };
  });

  return { rows: rows };
}

// --- Helper to fetch participants for a training ------------------------------
function getTrainingParticipantsList(trainingId) {
  if (!trainingId) return [];
  const cleanId = String(trainingId).trim().toLowerCase();

  // Tier 1: Check per-training spreadsheet
  try {
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (ss) {
      const allSheets = ss.getSheets();
      const sheet = allSheets.find(s => {
        const clean = s.getName().toLowerCase().replace(/[^a-z0-9]/g, '');
        return clean === 'participants' || clean === 'trainingparticipants' || clean === 'participant';
      });
      if (sheet) {
        const rows = sheetToJson(sheet);
        if (rows && rows.length > 0) return rows;
      }

      // Check Attendance tab if Participants tab is empty/missing
      const attSheet = allSheets.find(s => {
        const clean = s.getName().toLowerCase().replace(/[^a-z0-9]/g, '');
        return clean === 'attendance' || clean === 'attendances';
      });
      if (attSheet) {
        const attRows = sheetToJson(attSheet);
        if (attRows && attRows.length > 0) {
          const uniqueEmp = new Map();
          attRows.forEach(a => {
            const empId = String(a.EmployeeNo || a.EmployeeID || a.EmpID || a.ID || '').trim();
            if (empId && !uniqueEmp.has(empId.toLowerCase())) {
              uniqueEmp.set(empId.toLowerCase(), {
                ID: empId,
                EmployeeID: empId,
                EmployeeName: a.EmployeeName || a.Name || empId,
                Department: a.Department || '-',
                Position: a.Position || 'Participant',
                SupervisorID: a.SupervisorID || '',
                SupervisorName: a.SupervisorName || '',
                SupervisorEmail: a.SupervisorEmail || ''
              });
            }
          });
          if (uniqueEmp.size > 0) return Array.from(uniqueEmp.values());
        }
      }
    }
  } catch (e) {}

  // Tier 2: Check central database sheet 'Participants' / 'TrainingParticipants'
  try {
    const pSheet = getSheet(SHEET_NAMES.trainingParticipants || 'Participants') || getSheet('TrainingParticipants') || getSheet('Participants');
    if (pSheet) {
      const pRows = sheetToJson(pSheet);
      const matched = pRows.filter(p => {
        const pTid = String(p.TrainingID || p.TrainingCode || p.ID || '').trim().toLowerCase();
        return pTid === cleanId;
      });
      if (matched.length > 0) return matched;
    }
  } catch(e) {}

  // Tier 3: Check central database sheet 'Attendance'
  try {
    const aSheet = getSheet(SHEET_NAMES.attendance || 'Attendance');
    if (aSheet) {
      const aRows = sheetToJson(aSheet);
      const matched = aRows.filter(a => {
        const aTid = String(a.TrainingID || a.TrainingCode || '').trim().toLowerCase();
        return aTid === cleanId;
      });
      if (matched.length > 0) {
        const uniqueEmp = new Map();
        matched.forEach(a => {
          const empId = String(a.EmployeeNo || a.EmployeeID || a.EmpID || a.ID || '').trim();
          if (empId && !uniqueEmp.has(empId.toLowerCase())) {
            uniqueEmp.set(empId.toLowerCase(), {
              ID: empId,
              EmployeeID: empId,
              EmployeeName: a.EmployeeName || a.Name || empId,
              Department: a.Department || '-',
              Position: a.Position || 'Participant',
              SupervisorID: a.SupervisorID || '',
              SupervisorName: a.SupervisorName || '',
              SupervisorEmail: a.SupervisorEmail || ''
            });
          }
        });
        if (uniqueEmp.size > 0) return Array.from(uniqueEmp.values());
      }
    }
  } catch(e) {}

  // Tier 4: Check ParticipantList JSON on training row
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (tSheet) {
      const tRows = sheetToJson(tSheet);
      const t = tRows.find(r => {
        const tid = String(r.ID || '').trim().toLowerCase();
        const tcode = String(r.Code || '').trim().toLowerCase();
        return tid === cleanId || tcode === cleanId;
      });
      if (t) {
        if (t.ParticipantList) {
          const parsed = typeof t.ParticipantList === 'string' ? JSON.parse(t.ParticipantList) : t.ParticipantList;
          if (Array.isArray(parsed) && parsed.length > 0) return parsed;
        }
        if (t.ParticipantsList) {
          const parsed = typeof t.ParticipantsList === 'string' ? JSON.parse(t.ParticipantsList) : t.ParticipantsList;
          if (Array.isArray(parsed) && parsed.length > 0) return parsed;
        }
        if (typeof t.Participants === 'string' && t.Participants.trim().startsWith('[')) {
          try {
            const parsed = JSON.parse(t.Participants.trim());
            if (Array.isArray(parsed) && parsed.length > 0) return parsed;
          } catch(e) {}
        }
      }
    }
  } catch(e) {}

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

function getOrCreateStandingReportSpreadsheet(reportName) {
  const repFolder = getOrCreateReportsFolder();
  let fileIter = repFolder.getFilesByName(reportName);
  let file;
  if (fileIter.hasNext()) {
    file = fileIter.next();
  } else {
    const files = repFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getName().toLowerCase().trim() === reportName.toLowerCase().trim()) {
        file = f;
        break;
      }
    }
  }

  if (file) {
    return SpreadsheetApp.openById(file.getId());
  }

  const ss = SpreadsheetApp.create(reportName);
  file = DriveApp.getFileById(ss.getId());
  file.moveTo(repFolder);
  return ss;
}

function exportReportToSpreadsheet(reportType, year) {
  try {
    let reportName = 'Employee Report';
    if (reportType === 'hours') reportName = 'Training Hours';
    else if (reportType === 'cost') reportName = 'Training Cost';
    else if (reportType === 'title') reportName = 'Training Title';
    else if (reportType === 'employee') reportName = 'Employee Report';
    else if (reportType === 'atp') reportName = 'Annual Training Plan';

    const filters = { year: year || 'All' };
    const res = parseServerRes(getFilteredReportData(reportType, filters));
    if (!res.success) return err(res.message);
    const data = res.data;

    const ss = getOrCreateStandingReportSpreadsheet(reportName);
    const sheet = ss.getActiveSheet();
    sheet.setName(reportName);

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
    const repFolder = getOrCreateReportsFolder();

    return ok({
      fileId: fileId,
      fileName: `${reportName}.xlsx`,
      sheetUrl: sheetUrl,
      downloadUrl: downloadUrl,
      folderId: repFolder.getId(),
      folderName: repFolder.getName(),
      folderUrl: repFolder.getUrl(),
      message: `Report updated in standing spreadsheet '${reportName}'.`
    });
  } catch (e) {
    return err('Failed to export Excel report: ' + e.message);
  }
}

// --- Export Filtered Report to Excel / Google Sheets --------------------------
/**
 * Creates formatted Google Sheet & returns XLSX download link matching user's exact header specs.
 * Supports single report exports as well as 'master' / 'all' multi-tab exports.
 */
function exportFilteredReportExcel(reportType, filters) {
  try {
    if (reportType === 'master' || !reportType) {
      reportType = 'hours';
    }

    const res = parseServerRes(getFilteredReportData(reportType, filters));
    if (!res.success) return err(res.message || 'Failed to parse report data.');
    const data = res.data || {};
    const year = (filters && filters.year) ? filters.year : String(new Date().getFullYear());
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

// --- Non-Destructive Report Synchronization & Admin Edit Protection Engine ------

const REPORT_CONFIG = {
  'Training Hours': {
    sheetName: 'Training Hours',
    headerRowIndex: 2,
    dataStartRow: 3,
    keyColumnIndex: 1,
    keyFieldName: 'costCentre',
    columns: [
      { name: 'Cost Centre/Month', field: 'costCentre', isKey: true },
      { name: 'Jan', field: 'm0' },
      { name: 'Feb', field: 'm1' },
      { name: 'Mar', field: 'm2' },
      { name: 'Apr', field: 'm3' },
      { name: 'May', field: 'm4' },
      { name: 'Jun', field: 'm5' },
      { name: 'Jul', field: 'm6' },
      { name: 'Aug', field: 'm7' },
      { name: 'Sep', field: 'm8' },
      { name: 'Oct', field: 'm9' },
      { name: 'Nov', field: 'm10' },
      { name: 'Dec', field: 'm11' },
      { name: 'Total Year', field: 'totalYear' },
      { name: 'Total Training Hours', field: 'totalHours' }
    ]
  },
  'Training Cost': {
    sheetName: 'Training Cost',
    headerRowIndex: 2,
    dataStartRow: 3,
    keyColumnIndex: 1,
    keyFieldName: 'trainingTitle',
    columns: [
      { name: 'Training Title', field: 'trainingTitle', isKey: true },
      { name: 'Training Date (From)', field: 'dateFrom' },
      { name: 'Training Date (To)', field: 'dateTo' },
      { name: 'Total Participant', field: 'totalParticipant' },
      { name: 'Training Fees', field: 'trainingFees' },
      { name: 'Meal', field: 'meal' },
      { name: 'Subsistance Allowance', field: 'subsistanceAllowance' },
      { name: 'Hotel Fees', field: 'hotelFees' },
      { name: 'Mileage Claim', field: 'mileageClaim' },
      { name: 'Taxi Fees', field: 'taxiFees' },
      { name: 'Toll Fees', field: 'tollFees' },
      { name: 'Flight', field: 'flight' },
      { name: 'Total Cost', field: 'totalCost' },
      { name: 'Total HRDF Grant (RM)', field: 'totalHrdfGrant' }
    ]
  },
  'Training Title': {
    sheetName: 'Training Title',
    headerRowIndex: 1,
    dataStartRow: 2,
    keyColumnIndex: 2,
    keyFieldName: 'empNoKey',
    columns: [
      { name: 'Name', field: 'name' },
      { name: 'Emp. No', field: 'empNo', isKey: true },
      { name: 'Training Title', field: 'trainingTitle' },
      { name: '(EE)/Cost Centre Description', field: 'costCentreDesc' },
      { name: 'Training Type', field: 'trainingType' },
      { name: 'Date (From)', field: 'dateFrom' },
      { name: 'Date (Until)', field: 'dateUntil' },
      { name: 'Month', field: 'month' },
      { name: 'Total Hours', field: 'totalHours' },
      { name: 'Trainer', field: 'trainer' },
      { name: 'Training Provide', field: 'trainingProvider' },
      { name: 'Expiry Date (if applicable)', field: 'expiryDate' }
    ]
  },
  'Employee Report': {
    sheetName: 'Employee Report',
    headerRowIndex: 1,
    dataStartRow: 2,
    keyColumnIndex: 2,
    keyFieldName: 'empNo',
    columns: [
      { name: 'No', field: 'no' },
      { name: 'Emp. No', field: 'empNo', isKey: true },
      { name: 'Employee Name', field: 'name' },
      { name: 'Cost Centre / Department', field: 'costCentre' },
      { name: 'Position', field: 'position' },
      { name: 'Total Trainings', field: 'totalTrainings' },
      { name: 'Total Hours', field: 'totalHours' },
      { name: 'Attended Trainings', field: 'attendedList' },
      { name: 'Status', field: 'status' }
    ]
  },
  'Annual Training Plan (ATP)': {
    sheetName: 'Annual Training Plan (ATP)',
    headerRowIndex: 1,
    dataStartRow: 2,
    keyColumnIndex: 2,
    keyFieldName: 'trainingTitle',
    columns: [
      { name: 'No', field: 'no' },
      { name: 'Training Title', field: 'trainingTitle', isKey: true },
      { name: 'Training Category', field: 'trainingCategory' },
      { name: 'TNA Source', field: 'tnaSource' },
      { name: 'Training Mode', field: 'trainingMode' },
      { name: 'Training Duration (hrs)', field: 'durationHours' },
      { name: 'Trainer', field: 'trainer' },
      { name: 'Department', field: 'department' },
      { name: 'Position / Employee No', field: 'positionFormula' },
      { name: 'Total Pax', field: 'totalPax' },
      { name: 'Planned Date', field: 'plannedDate' },
      { name: 'Actual Date (From)', field: 'actualDateFrom' },
      { name: 'Actual Date (To)', field: 'actualDateTo' },
      { name: 'Training Status', field: 'trainingStatus' },
      { name: 'Remarks', field: 'remarks' }
    ]
  }
};

function getOrCreateSyncMetadataSheet(ss) {
  let metaSheet = ss.getSheetByName('_SYNC_METADATA');
  if (!metaSheet) {
    metaSheet = ss.insertSheet('_SYNC_METADATA');
    metaSheet.appendRow(['Report Sheet', 'Record Key', 'Column Name', 'Last System Value', 'Last Sync Timestamp']);
    metaSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#334155').setFontColor('#FFFFFF');
    try { metaSheet.hideSheet(); } catch (e) {}
  } else {
    try { metaSheet.hideSheet(); } catch (e) {}
  }
  return metaSheet;
}

function loadSyncMetadata(ss) {
  const metaSheet = getOrCreateSyncMetadataSheet(ss);
  const data = metaSheet.getDataRange().getValues();
  const map = {};

  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      const sheetName = String(data[i][0] || '').trim();
      const recKey    = String(data[i][1] || '').trim();
      const colName   = String(data[i][2] || '').trim();
      const lastVal   = data[i][3];

      if (sheetName && recKey && colName) {
        if (!map[sheetName]) map[sheetName] = {};
        if (!map[sheetName][recKey]) map[sheetName][recKey] = {};
        map[sheetName][recKey][colName] = lastVal;
      }
    }
  }
  return map;
}

function saveSyncMetadata(ss, metadataMap) {
  const metaSheet = getOrCreateSyncMetadataSheet(ss);
  metaSheet.clearContents();
  metaSheet.appendRow(['Report Sheet', 'Record Key', 'Column Name', 'Last System Value', 'Last Sync Timestamp']);

  const timeStamp = new Date().toISOString();
  const rows = [];

  Object.keys(metadataMap).forEach(sheetName => {
    Object.keys(metadataMap[sheetName]).forEach(recKey => {
      Object.keys(metadataMap[sheetName][recKey]).forEach(colName => {
        const val = metadataMap[sheetName][recKey][colName];
        rows.push([sheetName, recKey, colName, val, timeStamp]);
      });
    });
  });

  if (rows.length > 0) {
    metaSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  metaSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#334155').setFontColor('#FFFFFF');
  try { metaSheet.hideSheet(); } catch (e) {}
}

function normalizeSyncValue(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return val.toISOString().slice(0, 10);
  }
  let str = String(val).trim();
  if (str === '-' || str === 'N/A' || str === 'null' || str === 'undefined') {
    return '';
  }
  return str;
}

function isCellAdminModified(currentSheetValue, lastSystemValue) {
  const normCurrent = normalizeSyncValue(currentSheetValue);
  const normLast = normalizeSyncValue(lastSystemValue);
  return normCurrent !== normLast;
}

function ensureReportHeaderExists(sheet, reportKey, config, extraParams) {
  if (sheet.getLastRow() > 0) return;

  const year = (extraParams && extraParams.year) ? extraParams.year : '2026';

  if (reportKey === 'Training Hours') {
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
    sheet.getRange('A1:O2').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('center');
  } else if (reportKey === 'Training Cost') {
    const row1 = ['Training Title', 'Training Date (From)', 'Training Date (To)', 'Total Participant', 'Training Cost (RM)', '', '', '', '', '', '', '', `Total Cost ${year}`, 'Total HRDF Grant (RM)'];
    const row2 = ['', '', '', '', 'Training Fees', 'Meal', 'Subsistance Allowance', 'Hotel Fees', 'Mileage Claim', 'Taxi Fees', 'Toll Fees', 'Flight', '', ''];
    sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
    sheet.getRange(2, 1, 1, row2.length).setValues([row2]);
    sheet.getRange('A1:A2').merge().setValue('Training Title');
    sheet.getRange('B1:B2').merge().setValue('Training Date (From)');
    sheet.getRange('C1:C2').merge().setValue('Training Date (To)');
    sheet.getRange('D1:D2').merge().setValue('Total Participant');
    sheet.getRange('E1:L1').merge().setValue('Training Cost (RM)');
    sheet.getRange('M1:M2').merge().setValue(`Total Cost ${year}`);
    sheet.getRange('N1:N2').merge().setValue('Total HRDF Grant (RM)');
    sheet.getRange('A1:N2').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('center');
  } else {
    const headerCols = config.columns.map(c => c.name);
    sheet.getRange(1, 1, 1, headerCols.length).setValues([headerCols]);
    sheet.getRange(1, 1, 1, headerCols.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  }
}

function getRecordValueForColumn(record, colDef, rowNumber) {
  const field = colDef.field;
  if (!field) return '';
  if (field.startsWith('m') && field.length <= 3 && !isNaN(parseInt(field.slice(1)))) {
    const idx = parseInt(field.slice(1), 10);
    const months = Array.isArray(record.months) ? record.months : [];
    return months[idx] !== undefined ? months[idx] : 0;
  }
  if (field === 'positionFormula') {
    const reqUrl = record.requisitionFormUrl || '#';
    return (reqUrl && reqUrl !== '#') ? `=HYPERLINK("${reqUrl}", "Name List")` : 'Name List';
  }
  if (field === 'totalCost' || field === 'totalCostFormula') {
    const rNum = rowNumber || 3;
    return `=SUM(E${rNum}:L${rNum})`;
  }
  if (['trainingFees', 'meal', 'subsistanceAllowance', 'hotelFees', 'mileageClaim', 'taxiFees', 'tollFees', 'flight', 'totalHrdfGrant'].includes(field)) {
    const val = record[field];
    if (val === 0 || val === '0' || val === 'RM 0.00' || val === 'RM 0' || val === null || val === undefined) {
      return '';
    }
    return val;
  }
  if (field === 'attendedList' && Array.isArray(record.attendedList)) {
    return record.attendedList.join(', ');
  }
  return record[field] !== undefined ? record[field] : '';
}

function updateTrainingHoursTotalRow(sheet, rows, dataStartRow) {
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

  const totalRowValues = [
    'Total',
    ...totalMonths,
    grandTotalYear,
    grandTotalHours
  ];

  const lastRow = sheet.getLastRow();
  let totalRowIndex = lastRow;

  if (lastRow >= dataStartRow) {
    const firstCellVal = String(sheet.getRange(lastRow, 1).getValue() || '').trim().toLowerCase();
    if (firstCellVal === 'total') {
      totalRowIndex = lastRow;
    } else {
      totalRowIndex = lastRow + 1;
    }
  } else {
    totalRowIndex = dataStartRow;
  }

  sheet.getRange(totalRowIndex, 1, 1, 15).setValues([totalRowValues]);
  const totalRange = sheet.getRange(totalRowIndex, 1, 1, 15);
  totalRange.setFontWeight('bold').setBackground('#EFF6FF').setFontColor('#1E3A8A');
}

function autoFitSheetColumns(sheet, numCols) {
  if (!sheet) return;
  try {
    const totalCols = numCols || sheet.getLastColumn();
    if (totalCols > 0) {
      sheet.autoResizeColumns(1, totalCols);
      for (let col = 1; col <= totalCols; col++) {
        try { sheet.autoResizeColumn(col); } catch (e) {}
      }
    }
  } catch (e) {
    Logger.log('autoFitSheetColumns error: ' + e.message);
  }
}

function syncReportSheetIncrementally(ss, reportKey, reportData, extraParams) {
  const config = REPORT_CONFIG[reportKey];
  if (!config) {
    Logger.log('[SYNC] No configuration found for reportKey: ' + reportKey);
    return;
  }

  let sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(config.sheetName);
  }

  ensureReportHeaderExists(sheet, reportKey, config, extraParams);

  const metadataMap = loadSyncMetadata(ss);
  if (!metadataMap[config.sheetName]) {
    metadataMap[config.sheetName] = {};
  }
  const sheetMeta = metadataMap[config.sheetName];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  let existingSheetValues = [];
  if (lastRow >= config.dataStartRow && lastCol >= 1) {
    existingSheetValues = sheet.getRange(config.dataStartRow, 1, lastRow - config.dataStartRow + 1, Math.max(lastCol, config.columns.length)).getValues();
  }

  const rowMap = {};
  existingSheetValues.forEach((rowVals, idx) => {
    const sheetRowNumber = config.dataStartRow + idx;
    const rawKeyVal = rowVals[config.keyColumnIndex - 1];
    const keyVal = normalizeSyncValue(rawKeyVal);

    if (keyVal && keyVal.toLowerCase() !== 'total' && keyVal.toLowerCase() !== 'summary') {
      rowMap[keyVal.toLowerCase()] = {
        rowNumber: sheetRowNumber,
        rowValues: rowVals
      };
    }
  });

  const incomingRows = reportData.rows || [];

  incomingRows.forEach(record => {
    let recKey = '';
    if (config.keyFieldName === 'empNoKey') {
      recKey = normalizeSyncValue(`${record.empNo || record.name || ''}_${record.trainingTitle || ''}`);
    } else {
      recKey = normalizeSyncValue(record[config.keyFieldName] || record.trainingTitle || record.costCentre || record.empNo || record.ID || record.no);
    }
    if (!recKey) return;

    if (!sheetMeta[recKey]) {
      sheetMeta[recKey] = {};
    }
    const recordMeta = sheetMeta[recKey];

    const existingRow = rowMap[recKey.toLowerCase()];

    if (existingRow) {
      const rowNum = existingRow.rowNumber;
      const currentVals = existingRow.rowValues;

      config.columns.forEach((colDef, cIdx) => {
        const colNumber = cIdx + 1;
        const colName = colDef.name;
        const incomingVal = getRecordValueForColumn(record, colDef, rowNum);

        const currentSheetVal = currentVals[cIdx] !== undefined ? currentVals[cIdx] : '';
        let lastSystemVal     = recordMeta[colName];

        if (lastSystemVal === undefined) {
          lastSystemVal = currentSheetVal;
          if (normalizeSyncValue(currentSheetVal) !== '') {
            recordMeta[colName] = currentSheetVal;
          }
        }

        if (!isCellAdminModified(currentSheetVal, lastSystemVal)) {
          if (normalizeSyncValue(currentSheetVal) !== normalizeSyncValue(incomingVal)) {
            sheet.getRange(rowNum, colNumber).setValue(incomingVal);
            recordMeta[colName] = incomingVal;
            Logger.log(`[SYNC] ${config.sheetName} - ${recKey} - ${colName} updated to: "${incomingVal}"`);
          }
        } else {
          Logger.log(`[SYNC] ${config.sheetName} - ${recKey} - ${colName} preserved admin edit: "${currentSheetVal}"`);
        }
      });

    } else {
      const targetRowNum = sheet.getLastRow() + 1;
      const newRowVals = config.columns.map(colDef => getRecordValueForColumn(record, colDef, targetRowNum));
      sheet.appendRow(newRowVals);
      const newRowNum = sheet.getLastRow();

      config.columns.forEach((colDef, cIdx) => {
        recordMeta[colDef.name] = newRowVals[cIdx];
      });

      Logger.log(`[SYNC] ${config.sheetName} - ${recKey} - New record added at row ${newRowNum}`);
    }
  });

  if (reportKey === 'Training Hours') {
    updateTrainingHoursTotalRow(sheet, incomingRows, config.dataStartRow);
  }

  saveSyncMetadata(ss, metadataMap);
  autoFitSheetColumns(sheet, config.columns.length);
}

// --- Targeted Event-Based Incremental Synchronization Engine ------------------

function generateSyncId() {
  return 'SYNC-' + String(Date.now()).slice(-6);
}

function getOrCreateSyncHistorySheet(ss) {
  let historySheet = ss.getSheetByName('_SYNC_HISTORY');
  if (!historySheet) {
    historySheet = ss.insertSheet('_SYNC_HISTORY');
    historySheet.appendRow(['Sync ID', 'Timestamp', 'Training ID', 'Action', 'Trigger', 'Reports', 'Status', 'Duration', 'Message', 'Error']);
    historySheet.getRange('A1:J1').setFontWeight('bold').setBackground('#1E293B').setFontColor('#FFFFFF');
    try { historySheet.setFrozenRows(1); } catch (e) {}
  }
  return historySheet;
}

function logSyncHistory(ss, rec) {
  try {
    if (!ss) {
      const repFolder = getOrCreateReportsFolder();
      const fileName = `Master Annual Training Report (2026)`;
      let fileIter = repFolder.getFilesByName(fileName);
      if (fileIter.hasNext()) ss = SpreadsheetApp.openById(fileIter.next().getId());
    }
    if (!ss) return;

    const historySheet = getOrCreateSyncHistorySheet(ss);
    const syncId = rec.syncId || generateSyncId();
    const timeStamp = rec.timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    historySheet.appendRow([
      syncId,
      timeStamp,
      rec.trainingId || '',
      rec.action || 'UPDATE',
      rec.trigger || 'System Action',
      rec.reports || '',
      rec.status || 'SUCCESS',
      rec.duration || '0.00s',
      rec.message || '',
      rec.error || ''
    ]);
  } catch (e) {
    Logger.log('logSyncHistory error: ' + e.message);
  }
}

function buildEmployeeReportDataForTraining(tRecord, empRows) {
  const tSheet = getSheet(SHEET_NAMES.trainings);
  const allTrainings = tSheet ? sheetToJson(tSheet) : [tRecord];
  const participants = getTrainingParticipantsList(tRecord.ID || tRecord.Code);
  const enrolledEmpIds = new Set(participants.map(p => String(p.EmployeeID || p.ID || p.EmployeeNo || '').trim().toLowerCase()));

  let targetEmpRows = empRows;
  if (enrolledEmpIds.size > 0) {
    targetEmpRows = empRows.filter(e => {
      const eId = String(e.ID || e.EmployeeID || e.EmployeeNo || '').trim().toLowerCase();
      return enrolledEmpIds.has(eId);
    });
  }

  return buildEmployeeReportData(allTrainings, targetEmpRows.length > 0 ? targetEmpRows : empRows);
}

function buildHoursReportDataForTraining(tRecord, empRows, year) {
  const tSheet = getSheet(SHEET_NAMES.trainings);
  const allTrainings = tSheet ? sheetToJson(tSheet) : [tRecord];
  return buildHoursReportData(allTrainings, empRows, year);
}

/**
 * Reusable Incremental Sync Function (Targeted Event-Based Synchronization).
 * Synchronizes ONLY the specified Training ID across affected report sheets in real time.
 */
function syncTrainingById(trainingId, triggerName, actionName) {
  const lock = LockService.getScriptLock();
  let hasLock = false;
  try {
    hasLock = lock.tryLock(5000);
    if (!hasLock) {
      Logger.log('[SYNC] Could not acquire script lock within 5s for trainingId: ' + trainingId);
    }
  } catch (lErr) {
    Logger.log('[SYNC] Script lock error: ' + lErr.message);
  }

  const startTime = Date.now();
  const trigger = triggerName || 'System Trigger';
  const action = actionName || 'UPDATE';
  const displayYear = '2026';

  try {
    if (!trainingId || String(trainingId).trim() === '') {
      return err('Invalid Training ID provided for incremental sync.');
    }

    const tId = String(trainingId).trim();

    // 1. Retrieve the latest training record for trainingId
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return err('Trainings sheet unavailable.');
    const tRows = sheetToJson(tSheet);
    const tRecord = tRows.find(r => String(r.ID || '').trim() === tId || String(r.Code || '').trim() === tId);

    if (!tRecord) {
      const durationStr = ((Date.now() - startTime) / 1000).toFixed(2) + 's';
      const failMsg = `Unable to find Training ID: ${tId}`;
      logSyncHistory(null, {
        syncId: generateSyncId(),
        timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
        trainingId: tId,
        action: action,
        trigger: trigger,
        reports: 'None',
        status: 'FAILED',
        duration: durationStr,
        message: failMsg,
        error: `Training ID ${tId} not found in database.`
      });
      return err(failMsg);
    }

    // 2. Open Live Master Report Spreadsheet
    const repFolder = getOrCreateReportsFolder();
    const fileName = `Master Annual Training Report (${displayYear})`;

    let fileIter = repFolder.getFilesByName(fileName);
    let ss;
    if (fileIter.hasNext()) {
      ss = SpreadsheetApp.openById(fileIter.next().getId());
    } else {
      ss = SpreadsheetApp.create(fileName);
      const file = DriveApp.getFileById(ss.getId());
      repFolder.addFile(file);
      try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
    }

    const empSheet = getSheet(SHEET_NAMES.employees);
    const empRows = empSheet ? sheetToJson(empSheet) : [];

    // 3. Build targeted report data for ONLY this training record
    const atpData = buildAnnualTrainingPlanData([tRecord]);
    const costData = buildCostReportData([tRecord], displayYear);
    const titleData = buildTrainingTitleReportData([tRecord], empRows);
    const empReportData = buildEmployeeReportDataForTraining(tRecord, empRows);
    const hoursData = buildHoursReportDataForTraining(tRecord, empRows, displayYear);

    const reportsUpdated = [];

    // 4. Update affected report tabs incrementally while preserving admin-edited cells
    if (atpData && atpData.rows && atpData.rows.length > 0) {
      syncReportSheetIncrementally(ss, 'Annual Training Plan (ATP)', atpData);
      reportsUpdated.push('Annual Training Plan (ATP)');
    }

    if (costData && costData.rows && costData.rows.length > 0) {
      syncReportSheetIncrementally(ss, 'Training Cost', costData, { year: displayYear });
      reportsUpdated.push('Training Cost');
    }

    if (titleData && titleData.rows && titleData.rows.length > 0) {
      syncReportSheetIncrementally(ss, 'Training Title', titleData);
      reportsUpdated.push('Training Title');
    }

    if (empReportData && empReportData.rows && empReportData.rows.length > 0) {
      syncReportSheetIncrementally(ss, 'Employee Report', empReportData);
      reportsUpdated.push('Employee Report');
    }

    if (hoursData && hoursData.rows && hoursData.rows.length > 0) {
      syncReportSheetIncrementally(ss, 'Training Hours', hoursData, { year: displayYear });
      reportsUpdated.push('Training Hours');
    }

    SpreadsheetApp.flush();

    const durationStr = ((Date.now() - startTime) / 1000).toFixed(2) + 's';
    const syncId = generateSyncId();
    const message = `Successfully synchronized Training ID ${tId} across ${reportsUpdated.length} report tab(s).`;
    const timeStampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    const historyRecord = {
      syncId: syncId,
      timestamp: timeStampStr,
      trainingId: tId,
      action: action,
      trigger: trigger,
      reports: reportsUpdated.join(', '),
      status: 'SUCCESS',
      duration: durationStr,
      message: message,
      error: ''
    };

    logSyncHistory(ss, historyRecord);

    return ok({
      syncId: syncId,
      trainingId: tId,
      action: action,
      trigger: trigger,
      reports: reportsUpdated,
      status: 'SUCCESS',
      duration: durationStr,
      message: message,
      fileUrl: ss.getUrl(),
      timestamp: timeStampStr
    });

  } catch (e) {
    const durationStr = ((Date.now() - startTime) / 1000).toFixed(2) + 's';
    const errorMsg = 'Incremental sync failed for Training ID ' + trainingId + ': ' + e.message;
    Logger.log('syncTrainingById error: ' + errorMsg);

    try {
      const repFolder = getOrCreateReportsFolder();
      const fileName = `Master Annual Training Report (${displayYear})`;
      let fileIter = repFolder.getFilesByName(fileName);
      let ss = fileIter.hasNext() ? SpreadsheetApp.openById(fileIter.next().getId()) : null;
      if (ss) {
        logSyncHistory(ss, {
          syncId: generateSyncId(),
          timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
          trainingId: String(trainingId || 'UNKNOWN'),
          action: action,
          trigger: trigger,
          reports: 'None',
          status: 'FAILED',
          duration: durationStr,
          message: 'Synchronization failed',
          error: e.message
        });
      }
    } catch (hErr) {}

    return err(errorMsg);
  } finally {
    if (hasLock) {
      try { lock.releaseLock(); } catch(rErr) {}
    }
  }
}

/**
 * API: Get Sync History logs from _SYNC_HISTORY sheet
 */
function getSyncHistory(limit, filterTrainingId, filterStatus) {
  try {
    const repFolder = getOrCreateReportsFolder();
    const fileName = `Master Annual Training Report (2026)`;
    let fileIter = repFolder.getFilesByName(fileName);
    if (!fileIter.hasNext()) return ok([]);

    const ss = SpreadsheetApp.openById(fileIter.next().getId());
    const historySheet = ss.getSheetByName('_SYNC_HISTORY');
    if (!historySheet) return ok([]);

    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) return ok([]);

    let rows = [];
    for (let i = 1; i < data.length; i++) {
      rows.push({
        syncId:     String(data[i][0] || ''),
        timestamp:  String(data[i][1] || ''),
        trainingId: String(data[i][2] || ''),
        action:     String(data[i][3] || ''),
        trigger:    String(data[i][4] || ''),
        reports:    String(data[i][5] || ''),
        status:     String(data[i][6] || ''),
        duration:   String(data[i][7] || ''),
        message:    String(data[i][8] || ''),
        error:      String(data[i][9] || '')
      });
    }

    rows.reverse(); // Most recent first

    if (filterTrainingId) {
      const cleanId = String(filterTrainingId).toLowerCase().trim();
      rows = rows.filter(r => r.trainingId.toLowerCase().includes(cleanId));
    }
    if (filterStatus && filterStatus !== 'All') {
      const cleanStatus = String(filterStatus).toLowerCase().trim();
      rows = rows.filter(r => r.status.toLowerCase() === cleanStatus);
    }

    const maxLimit = limit ? parseInt(limit, 10) : 100;
    return ok(rows.slice(0, maxLimit));
  } catch (e) {
    return err('Failed to load sync history: ' + e.message);
  }
}

/**
 * API: Get latest Sync Status for a specific Training ID
 */
function getTrainingSyncStatus(trainingId) {
  try {
    if (!trainingId) return err('Training ID is required.');
    const res = parseServerRes(getSyncHistory(50, trainingId));
    if (!res.success) return err(res.message);

    const history = res.data || [];
    if (history.length === 0) {
      return ok({
        trainingId: trainingId,
        status: 'UNSYNCED',
        lastSync: 'Never',
        lastDuration: '-',
        lastTrigger: '-',
        lastAction: '-',
        lastError: '',
        reports: 'None'
      });
    }

    const latest = history[0];
    return ok({
      trainingId: trainingId,
      status: latest.status,
      lastSync: latest.timestamp,
      lastDuration: latest.duration,
      lastTrigger: latest.trigger,
      lastAction: latest.action,
      lastError: latest.error,
      reports: latest.reports
    });
  } catch (e) {
    return err('Failed to get sync status: ' + e.message);
  }
}

/**
 * Manual Recovery Function: Synchronizes all training records across all 5 report tabs.
 * Use for initial setup, data recovery, manual admin reconciliation, or scheduled maintenance.
 */
function syncAllTrainings(targetYear) {
  const startTime = Date.now();
  try {
    const res = syncLiveMasterReportSheet(targetYear);
    const durationStr = ((Date.now() - startTime) / 1000).toFixed(2) + 's';

    const repFolder = getOrCreateReportsFolder();
    const fileName = `Master Annual Training Report (${targetYear || '2026'})`;
    let fileIter = repFolder.getFilesByName(fileName);
    let ss = fileIter.hasNext() ? SpreadsheetApp.openById(fileIter.next().getId()) : null;

    if (ss) {
      logSyncHistory(ss, {
        syncId: generateSyncId(),
        timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
        trainingId: 'ALL',
        action: 'FULL_SYNC',
        trigger: 'Manual Admin Full Sync',
        reports: 'All 5 Reports',
        status: parseServerRes(res).success ? 'SUCCESS' : 'FAILED',
        duration: durationStr,
        message: 'Full master database synchronization executed.',
        error: parseServerRes(res).success ? '' : parseServerRes(res).message
      });
    }
    return res;
  } catch (e) {
    return err('Full sync failed: ' + e.message);
  }
}

/**
 * Synchronizes and updates the Live Master Report Google Sheet inside the "Reports" folder in Google Drive.
 * Contains all 5 live tabs (Training Hours, Training Cost, Training Title, Employee Report, ATP).
 * Uses cell-level admin edit protection with _SYNC_METADATA tracking.
 */
function syncLiveMasterReportSheet(targetYear) {
  try {
    const year = (targetYear && targetYear !== '2026') ? targetYear : 'All';
    const filters = { year: 'All' };

    // Update Standing Report Files
    const hoursRes = parseServerRes(getFilteredReportData('hours', filters));
    if (hoursRes.success && hoursRes.data) {
      const ssHours = getOrCreateStandingReportSpreadsheet('Training Hours');
      syncReportSheetIncrementally(ssHours, 'Training Hours', hoursRes.data, { year: displayYear });
    }

    const costRes = parseServerRes(getFilteredReportData('cost', filters));
    if (costRes.success && costRes.data) {
      const ssCost = getOrCreateStandingReportSpreadsheet('Training Cost');
      syncReportSheetIncrementally(ssCost, 'Training Cost', costRes.data, { year: displayYear });
    }

    const titleRes = parseServerRes(getFilteredReportData('title', filters));
    if (titleRes.success && titleRes.data) {
      const ssTitle = getOrCreateStandingReportSpreadsheet('Training Title');
      syncReportSheetIncrementally(ssTitle, 'Training Title', titleRes.data);
    }

    const empRes = parseServerRes(getFilteredReportData('employee', filters));
    if (empRes.success && empRes.data) {
      const ssEmp = getOrCreateStandingReportSpreadsheet('Employee Report');
      syncReportSheetIncrementally(ssEmp, 'Employee Report', empRes.data);
    }

    const atpRes = parseServerRes(getFilteredReportData('atp', filters));
    if (atpRes.success && atpRes.data) {
      const ssAtp = getOrCreateStandingReportSpreadsheet('Annual Training Plan');
      syncReportSheetIncrementally(ssAtp, 'Annual Training Plan', atpRes.data);
    }

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
      message: `Live Master Report successfully synced with cell-level admin edit protection!`
    });
  } catch (e) {
    Logger.log('syncLiveMasterReportSheet error: ' + e.message);
    return err('Failed to sync Live Master Report: ' + e.message);
  }
}

/**
 * Gets details of the "Reports" folder and the latest Live Master Report Sheet URL.
 * Automatically sorts by getLastUpdated() descending to ensure the link connects
 * to the latest created spreadsheet in case older files were deleted or replaced.
 */
function getReportsFolderDetails(yearInput) {
  try {
    const repFolder = getOrCreateReportsFolder();
    const targetYear = String(yearInput || '2026').trim();

    let fileUrl = '';
    let fileId = '';
    let latestSheetName = '';

    const matchingFiles = [];
    if (repFolder) {
      const sheetIter = repFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (sheetIter.hasNext()) {
        const file = sheetIter.next();
        const fName = file.getName();
        if (fName.toLowerCase().includes('master annual training report') || fName.toLowerCase().includes('report')) {
          matchingFiles.push(file);
        }
      }
    }

    if (matchingFiles.length > 0) {
      // Sort by getLastUpdated() descending so the latest created/updated sheet is always returned
      matchingFiles.sort((a, b) => b.getLastUpdated().getTime() - a.getLastUpdated().getTime());
      const latestFile = matchingFiles[0];
      fileId = latestFile.getId();
      fileUrl = latestFile.getUrl();
      latestSheetName = latestFile.getName();
    } else {
      // Fallback to SPREADSHEET_ID if no file in Reports folder
      const masterSsId = getMasterSpreadsheetId();
      if (masterSsId) {
        fileId = masterSsId;
        fileUrl = `https://docs.google.com/spreadsheets/d/${masterSsId}/edit`;
        latestSheetName = `Master Annual Training Report (${targetYear})`;
      }
    }

    return ok({
      folderId: repFolder ? repFolder.getId() : '',
      folderUrl: repFolder ? repFolder.getUrl() : '',
      liveMasterSheetId: fileId,
      liveMasterSheetUrl: fileUrl,
      liveMasterSheetName: latestSheetName,
      message: 'Latest master report sheet details retrieved.'
    });
  } catch (e) {
    return err(e.message);
  }
}

// --- Header Formatters for Excel Export & Live Master Sheets ------------------

// --- Header Formatters for Excel Export & Live Master Sheets ------------------

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
  const rowsData = rows.map((r, idx) => {
    const rNum = idx + 3;
    const mVals = (r.months || []).map(val => (val === 0 || val === '0' || !val) ? '' : val);
    return [
      r.costCentre || 'General',
      ...mVals,
      `=SUM(B${rNum}:M${rNum})`,
      `=SUM(B${rNum}:M${rNum})`
    ];
  });

  if (rowsData.length > 0) {
    sheet.getRange(3, 1, rowsData.length, 15).setValues(rowsData);

    const totalRowIndex = 3 + rowsData.length;
    const lastDataRow = totalRowIndex - 1;

    const colLetters = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];
    const totalFormulas = colLetters.map(col => `=SUM(${col}3:${col}${lastDataRow})`);

    const totalRow = [
      'Total',
      ...totalFormulas
    ];

    sheet.getRange(totalRowIndex, 1, 1, 15).setValues([totalRow]);
    const totalRange = sheet.getRange(totalRowIndex, 1, 1, 15);
    totalRange.setFontWeight('bold')
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A');
  }

  autoFitSheetColumns(sheet, 15);
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

  const rowsData = (data.rows || []).map((r, idx) => {
    const rNum = idx + 3;
    return [
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
      `=SUM(E${rNum}:L${rNum})`,
      r.totalHrdfGrant || ''
    ];
  });

  if (rowsData.length > 0) {
    sheet.getRange(3, 1, rowsData.length, 14).setValues(rowsData);

    const totalRowIndex = 3 + rowsData.length;
    const lastDataRow = totalRowIndex - 1;

    const totalRow = [
      'Total', '', '',
      `=SUM(D3:D${lastDataRow})`,
      `=SUM(E3:E${lastDataRow})`,
      `=SUM(F3:F${lastDataRow})`,
      `=SUM(G3:G${lastDataRow})`,
      `=SUM(H3:H${lastDataRow})`,
      `=SUM(I3:I${lastDataRow})`,
      `=SUM(J3:J${lastDataRow})`,
      `=SUM(K3:K${lastDataRow})`,
      `=SUM(L3:L${lastDataRow})`,
      `=SUM(M3:M${lastDataRow})`,
      `=SUM(N3:N${lastDataRow})`
    ];

    sheet.getRange(totalRowIndex, 1, 1, 14).setValues([totalRow]);
    const totalRange = sheet.getRange(totalRowIndex, 1, 1, 14);
    totalRange.setFontWeight('bold')
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A');
  }

  autoFitSheetColumns(sheet, 14);
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
    r.name || '-',
    r.empNo || '-',
    r.trainingTitle || '-',
    r.costCentreDesc || '-',
    r.trainingType || 'Public',
    r.dateFrom || '',
    r.dateUntil || '',
    r.month || '',
    r.totalHours || 8,
    r.trainer || '-',
    r.trainingProvider || '-',
    r.expiryDate || 'N/A'
  ]);

  if (rowsData.length > 0) {
    sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);

    const totalRowIndex = 2 + rowsData.length;
    const lastDataRow = totalRowIndex - 1;

    const totalRow = [
      'Total', '', '', '', '', '', '', '',
      `=SUM(I2:I${lastDataRow})`,
      '', '', ''
    ];

    sheet.getRange(totalRowIndex, 1, 1, headers.length).setValues([totalRow]);
    const totalRange = sheet.getRange(totalRowIndex, 1, 1, headers.length);
    totalRange.setFontWeight('bold')
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A');
  }

  autoFitSheetColumns(sheet, headers.length);
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
    r.empNo || '-',
    r.name || '-',
    r.costCentre || '-',
    r.position || '-',
    r.totalTrainings || 0,
    r.totalHours || 0,
    Array.isArray(r.attendedList) ? r.attendedList.join(', ') : (r.attendedList || '-'),
    r.status || 'Active'
  ]);

  if (rowsData.length > 0) {
    sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);

    const totalRowIndex = 2 + rowsData.length;
    const lastDataRow = totalRowIndex - 1;

    const totalRow = [
      'Total', '', '', '', '',
      `=SUM(F2:F${lastDataRow})`,
      `=SUM(G2:G${lastDataRow})`,
      '', ''
    ];

    sheet.getRange(totalRowIndex, 1, 1, headers.length).setValues([totalRow]);
    const totalRange = sheet.getRange(totalRowIndex, 1, 1, headers.length);
    totalRange.setFontWeight('bold')
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A');
  }

  autoFitSheetColumns(sheet, headers.length);
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

  const rows = data.rows || [];
  const rowsData = rows.map((r, i) => {
    const reqUrl = r.requisitionFormUrl || '#';
    const posFormula = (reqUrl && reqUrl !== '#') ? `=HYPERLINK("${reqUrl}", "Name List")` : 'Name List';

    return [
      r.no || (i + 1),
      r.trainingTitle || '-',
      r.trainingCategory || 'General',
      r.tnaSource || 'Training Requisition Form',
      r.trainingMode || 'In-house',
      r.durationHours || 8,
      r.trainer || '-',
      r.department || '-',
      posFormula,
      r.totalPax || 0,
      r.plannedDate || '',
      r.actualDateFrom || '',
      r.actualDateTo || '',
      r.trainingStatus || 'On planning',
      r.remarks || 'Planned'
    ];
  });

  if (rowsData.length > 0) {
    sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);

    const totalRowIndex = 2 + rowsData.length;
    const lastDataRow = totalRowIndex - 1;

    const totalRow = [
      'Total', '', '', '', '',
      `=SUM(F2:F${lastDataRow})`,
      '', '', '',
      `=SUM(J2:J${lastDataRow})`,
      '', '', '', '', ''
    ];

    sheet.getRange(totalRowIndex, 1, 1, headers.length).setValues([totalRow]);
    const totalRange = sheet.getRange(totalRowIndex, 1, 1, headers.length);
    totalRange.setFontWeight('bold')
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A');
  }

  autoFitSheetColumns(sheet, headers.length);
}

// --- Export Single Attendance Sheet -------------------------------------------
function exportAttendanceSheet(trainingId) {
  try {
    const tResult = safeParseObj(getTrainingById(trainingId));
    if (!tResult.success) return err('Training not found.');
    const training = tResult.data;

    const attData = safeParseObj(getAttendance(trainingId)).data || [];

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

    autoFitSheetColumns(sheet, headers.length);
    return ok({ sheetName: name, message: 'Export sheet created: ' + name });
  } catch (e) {
    return err('Export failed: ' + e.message);
  }
}

// --- Export Single Training Full Report to Excel (.xlsx / Sheet) --------------
function exportReportExcel(trainingId) {
  try {
    const reportRes = safeParseObj(generateReport(trainingId));
    if (!reportRes.success) return err(reportRes.message);
    const rep = reportRes.data;
    const t = rep.training;

    const ssName = `${t.Code}_Full_Report`;
    const ss = SpreadsheetApp.create(ssName);
    const file = DriveApp.getFileById(ss.getId());

    try {
      const trainingRoot = getOrCreateTrainingRootFolder();
      const folderName = `${t.Code} ${t.Name}`.trim();
      let folderIter = trainingRoot.getFoldersByName(folderName);
      let targetFolder = folderIter.hasNext() ? folderIter.next() : null;
      if (!targetFolder && t.FolderID) {
        try { targetFolder = DriveApp.getFolderById(t.FolderID); } catch(e) {}
      }
      if (targetFolder) {
        file.moveTo(targetFolder);
      }
    } catch (mErr) {
      Logger.log('Error moving single training report file: ' + mErr.message);
    }

    const sSummary = ss.getActiveSheet();
    sSummary.setName('Programme Overview');
    sSummary.appendRow(['TRAINING PROGRAMME SUMMARY REPORT']);
    sSummary.appendRow(['Generated Date:', rep.generatedAt]);
    sSummary.appendRow([]);
    sSummary.appendRow(['Programme Code', t.Code]);
    sSummary.appendRow(['Programme Name', t.Name]);
    sSummary.appendRow(['Category', t.Category]);
    sSummary.appendRow(['Trainer / Facilitator', t.Trainer]);
    sSummary.appendRow(['Venue', t.Venue || '-']);
    sSummary.appendRow(['Course Fee (RM)', t.CourseFee || '0.00']);
    sSummary.appendRow(['Department', t.Department || 'All Departments']);
    sSummary.appendRow(['Start Date', formatOrdinalDate(t.StartDate)]);
    sSummary.appendRow(['End Date', formatOrdinalDate(t.EndDate || t.StartDate)]);
    sSummary.appendRow(['Duration (Days)', t.Duration]);
    sSummary.appendRow(['Total Hours', t.TotalHours]);
    sSummary.appendRow(['Status', t.Status]);
    sSummary.appendRow(['Lifecycle Stage', t.Stage]);
    sSummary.appendRow(['Enrolled Participants', t.Participants]);
    sSummary.appendRow(['Objectives', t.Objectives || '-']);
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


// --- Helpers --------------------------------------------------------------------
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
