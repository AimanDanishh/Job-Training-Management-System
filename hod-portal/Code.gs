/**
 * Code.gs — HOD Portal Web App Entry Point & Server Engine
 */

/**
 * Code.gs — HOD Portal Web App Entry Point & Server Engine
 */

function doGet(e) {
  const pageParam = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase().trim() : 'review';
  const emailParam = (e && e.parameter && e.parameter.email) ? String(e.parameter.email).trim() : '';

  const pageMap = {
    'review':   'HodReview',
    'posteval': 'HodPostEvaluation',
    'success':  'Success',
    'error':    'Error'
  };

  const templateName = pageMap[pageParam] || 'HodReview';
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — Approval Portal');

  // Determine user email strictly from Session if available, fallback to parameter
  let activeEmail = resolveActiveSessionEmail(emailParam);

  // If active user email is available, validate authorization
  if (activeEmail) {
    const auth = validateHODAccess(activeEmail);
    if (!auth.valid) {
      try {
        const errTemplate = HtmlService.createTemplateFromFile('Error');
        errTemplate.params = { 
          title: 'Access Denied — Non Authorized Account',
          message: auth.message || `Access Denied: The email address (${activeEmail}) is not registered as an authorized approver.`
        };
        return errTemplate.evaluate()
          .setTitle('Access Denied — TrainHub Approval Portal')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      } catch (err) {}
    }
  }

  try {
    const template = HtmlService.createTemplateFromFile(templateName);
    template.params = (e && e.parameter) ? e.parameter : {};
    template.page = pageParam;

    return template.evaluate()
      .setTitle(appTitle)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log('doGet routing error: ' + err.message);
    try {
      const errTemplate = HtmlService.createTemplateFromFile('Error');
      errTemplate.params = { message: 'Failed to load page: ' + err.message };
      return errTemplate.evaluate()
        .setTitle('Error — TrainHub Approval Portal')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (fallbackErr) {
      return HtmlService.createHtmlOutput('<h3 style="font-family:sans-serif;color:#ef4444;padding:20px;">System Error: ' + err.message + '</h3>');
    }
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return '';
  }
}

/**
 * Helper: Resolve active Google Session user email server-side
 */
function resolveActiveSessionEmail(providedEmail) {
  let sessionEmail = '';
  try {
    sessionEmail = Session.getActiveUser().getEmail();
  } catch (e) {}
  if (sessionEmail && sessionEmail.trim() !== '') {
    return sessionEmail.trim();
  }
  return String(providedEmail || '').trim();
}

/**
 * API: Verify if user email is authorized in HOD email directory
 */
function verifyHODEmail(userEmail) {
  try {
    const activeEmail = resolveActiveSessionEmail(userEmail);
    const auth = validateHODAccess(activeEmail);
    if (!auth.valid) {
      return err(auth.message);
    }
    return ok({
      email: auth.email,
      hod: auth.hod
    });
  } catch (e) {
    return err('Verification error: ' + e.message);
  }
}

/**
 * Helpers: Safely retrieve training ID and Code from a training row object
 */
function getTrainingId(t) {
  if (!t) return '';
  return String(t.ID || t.Code || t.TrainingID || t['Training ID'] || t.RequisitionID || t['Requisition ID'] || t.TrainingCode || t['Training Code'] || '').trim();
}

function getTrainingCode(t) {
  if (!t) return '';
  return String(t.Code || t.TrainingCode || t['Training Code'] || t.RequisitionCode || t['Requisition Code'] || t.ID || '').trim();
}

/**
 * API: Fetch complete dashboard data for the authenticated approver.
 * Dynamic summary statistics (Pending, Approved, Rejected, Total),
 * pending requests, request history, and Employee Portal URL.
 */
function getApproverDashboardData(userEmailParam) {
  try {
    const activeEmail = resolveActiveSessionEmail(userEmailParam);
    if (!activeEmail) {
      return err('Unable to identify your Google account. Please ensure you are logged into your authorized company account.');
    }

    const auth = validateHODAccess(activeEmail);
    if (!auth.valid) {
      return err(auth.message || `Access Denied: Account (${activeEmail}) is not registered as an authorized approver.`);
    }

    const hodProfile = auth.hod;
    const supervisedRequisitions = getHODSupervisedRequisitions(hodProfile);

    let pendingCount = 0;
    let approvedCount = 0;
    let rejectedCount = 0;
    let totalCount = supervisedRequisitions.length;

    const pendingRequests = [];
    const historyRequests = [];

    supervisedRequisitions.forEach(r => {
      const statusStr = String(r.ApprovalStatus || r.Status || 'Pending HOD Approval').trim();
      const statusLower = statusStr.toLowerCase();

      const item = {
        ID: getTrainingId(r),
        Code: getTrainingCode(r),
        Name: String(r.Name || r.TrainingName || 'Training Request').trim(),
        Category: String(r.Category || 'General').trim(),
        RequestedBy: String(r.RequestedByName || r.RequestedBy || 'Employee Requester').trim(),
        EmployeeID: String(r.RequestedBy || r.EmployeeID || 'N/A').trim(),
        RequestedByEmail: String(r.RequestedByEmail || '').trim(),
        Department: String(r.Department || r.CostCentre || 'N/A').trim(),
        Trainer: String(r.Trainer || 'TBD').trim(),
        Venue: String(r.Venue || 'TBD').trim(),
        StartDate: String(r.StartDate || '').trim(),
        EndDate: String(r.EndDate || r.StartDate || '').trim(),
        RequestedDate: String(r.RequestedDate || r.CreatedDate || '').trim(),
        Duration: r.Duration || 1,
        TotalHours: r.TotalHours || 8,
        CourseFee: String(r.CourseFee || '0.00').trim(),
        Objectives: String(r.Objectives || '').trim(),
        ApprovalStatus: statusStr,
        ApprovalRemarks: String(r.ApprovalRemarks || '').trim(),
        ApprovedBy: String(r.ApprovedBy || '').trim(),
        ApprovedAt: String(r.ApprovedAt || '').trim(),
        ApprovedCostCentre: String(r.ApprovedCostCentre || '').trim(),
        BrochureURL: String(r.BrochureURL || r.BrochureUrl || '').trim()
      };

      if (statusLower.includes('pending')) {
        pendingCount++;
        pendingRequests.push(item);
      } else if (statusLower.includes('approved')) {
        approvedCount++;
        historyRequests.push(item);
      } else if (statusLower.includes('reject') || statusLower.includes('returned') || statusLower.includes('cancel')) {
        rejectedCount++;
        historyRequests.push(item);
      } else {
        historyRequests.push(item);
      }
    });

    const empPortalUrl = getEmployeePortalUrl();

    return ok({
      approver: hodProfile,
      metrics: {
        pending: pendingCount,
        approved: approvedCount,
        rejected: rejectedCount,
        total: totalCount
      },
      pendingRequests: pendingRequests,
      historyRequests: historyRequests,
      allSupervisedRequests: supervisedRequisitions,
      employeePortalUrl: empPortalUrl
    });
  } catch (e) {
    Logger.log('getApproverDashboardData error: ' + e.message);
    return err('Failed to load dashboard data: ' + e.message);
  }
}

/**
 * API: Fetch complete Training Requisition details and Supervision Queue for HOD Review
 */
function getRequisitionDetails(trainingId, userEmail) {
  try {
    const activeEmail = resolveActiveSessionEmail(userEmail);
    const auth = validateHODAccess(activeEmail);
    if (!auth.valid) return err(auth.message);


    const hodProfile = auth.hod;

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const trainings = sheetToJson(tSheet);
    
    // Get all requisitions strictly under this HOD's supervision / assigned employees
    const supervisedRequisitions = getHODSupervisedRequisitions(hodProfile);

    let training = null;
    if (trainingId) {
      const cleanId = String(trainingId).trim().toLowerCase();
      training = supervisedRequisitions.find(t => {
        const tid = getTrainingId(t).toLowerCase();
        const tcode = getTrainingCode(t).toLowerCase();
        return tid === cleanId || tcode === cleanId;
      });
      if (!training) {
        // Search in master trainings if not directly in supervised queue
        training = trainings.find(t => {
          const tid = getTrainingId(t).toLowerCase();
          const tcode = getTrainingCode(t).toLowerCase();
          return tid === cleanId || tcode === cleanId;
        });
      }
    }

    // Default to first pending training in HOD's supervision queue
    if (!training) {
      if (supervisedRequisitions.length > 0) {
        training = supervisedRequisitions.find(t => String(t.ApprovalStatus || t.Status || '').toLowerCase().includes('pending')) || supervisedRequisitions[0];
      }
    }

    if (!training) {
      return ok({
        training: null,
        participants: [],
        requester: null,
        hod: hodProfile,
        supervisedRequisitions: [],
        message: 'No pending training requisitions found under your supervision.'
      });
    }

    // Fetch Participants for this training using 3-level fallback
    let participants = [];

    // Level 1: Dedicated single Drive spreadsheet (TrainingParticipants tab)
    try {
      const ss = getTrainingDataSpreadsheet(training.ID);
      const partSheet = ss ? ss.getSheetByName('TrainingParticipants') : null;
      if (partSheet) {
        participants = sheetToJson(partSheet);
      }
    } catch(e1) {}

    // Level 2: Master database TrainingParticipants sheet tab
    if (!participants || participants.length === 0) {
      try {
        const masterTpSheet = getSheet('TrainingParticipants');
        if (masterTpSheet) {
          const allTp = sheetToJson(masterTpSheet);
          participants = allTp.filter(r => String(r.TrainingID || r.TrainingId || '').trim() === String(training.ID).trim());
        }
      } catch(e2) {}
    }

    // Level 3: Extract from AP-HRD-F01-01 Training Requisition Form Google Sheet (rows 15-39)
    if (!participants || participants.length === 0) {
      try {
        const formFileId = training.RequisitionFormFileID;
        if (formFileId) {
          const formSs = SpreadsheetApp.openById(formFileId);
          const formSheet = formSs.getSheetByName('Training Form') || formSs.getSheets()[0];
          const rows = formSheet.getRange('A15:I39').getValues();
          rows.forEach(r => {
            const empId = String(r[0] || '').trim();
            const name  = String(r[2] || '').trim();
            const dept  = String(r[3] || '').trim();
            const pos   = String(r[7] || '').trim();
            if (empId || name) {
              participants.push({
                EmployeeID: empId,
                EmployeeName: name,
                Department: dept,
                Position: pos
              });
            }
          });
        }
      } catch(e3) {}
    }

    // Requester employee info lookup
    let requester = {
      ID: training.RequestedBy || training.EmployeeID || 'N/A',
      Name: training.RequestedByName || training.Trainer || 'Employee Requester',
      Department: training.Department || 'N/A',
      Email: training.RequestedByEmail || ''
    };

    const empSheet = getSheet('Employees');
    if (empSheet) {
      const emps = sheetToJson(empSheet);
      const matched = emps.find(e => {
        const eId = String(e['Employee No'] || e.EmployeeNo || e.ID || e.EmployeeID || '').toLowerCase().trim();
        const eEmail = String(e.Email || e.EmailAddress || '').toLowerCase().trim();
        return (eId && eId === String(requester.ID).toLowerCase()) || (eEmail && eEmail === String(requester.Email).toLowerCase());
      });
      if (matched) {
        requester.ID = String(matched['Employee No'] || matched.EmployeeNo || matched.ID || requester.ID).trim();
        requester.Name = String(matched.Name || matched.EmployeeName || matched['HOD'] || requester.Name).trim();
        requester.Department = String(matched['Cost Centre'] || matched.Department || requester.Department).trim();
        requester.Email = String(matched.Email || requester.Email).trim();
      }
    }

    // Cleanse legacy stamp data if training has old dummy text
    if (training.ApprovedBy && (training.ApprovedBy.includes('IT.INTERN') || training.ApprovedBy.includes('It Intern') || training.ApprovedBy.includes('EMP-HOD001') || training.ApprovedBy.includes('Head of Department'))) {
      training.ApprovedBy = `${hodProfile.Name} (${hodProfile.ID})`;
    }
    if (!training.ApprovedCostCentre || training.ApprovedCostCentre.includes('All Departments')) {
      training.ApprovedCostCentre = hodProfile.CostCentre;
    }

    return ok({
      training: training,
      participants: participants,
      requester: requester,
      hod: hodProfile,
      supervisedRequisitions: supervisedRequisitions
    });
  } catch (e) {
    return err('Failed to load requisition details: ' + e.message);
  }
}

/**
 * Helper: Retrieve all requisitions strictly under a specific HOD's supervision.
 */
function getHODSupervisedRequisitions(hodProfile) {
  try {
    const tSheet = getSheet('Trainings');
    if (!tSheet) return [];
    const trainings = sheetToJson(tSheet);
    if (!trainings || trainings.length === 0) return [];

    const hodNameClean = String(hodProfile ? hodProfile.Name : '').toLowerCase().trim();
    const hodIdClean = String(hodProfile ? hodProfile.ID : '').toLowerCase().trim();
    const hodCcClean = String(hodProfile ? hodProfile.CostCentre : '').toLowerCase().trim();
    const hodEmailClean = String(hodProfile ? hodProfile.Email : '').toLowerCase().trim();
    const isCsuite = hodProfile && (hodProfile.isCsuite || (hodProfile.Position || '').toLowerCase().includes('c-suite') || (hodProfile.Position || '').toLowerCase().includes('csuite') || (hodProfile.Position || '').toLowerCase().includes('ceo') || (hodProfile.Position || '').toLowerCase().includes('chief'));
    const isHohr = hodProfile && (hodProfile.isHohr || (hodProfile.Position || '').toLowerCase().includes('hohr') || (hodProfile.Position || '').toLowerCase().includes('head of hr') || (hodProfile.Position || '').toLowerCase().includes('hr head'));

    // 1. Build supervised employee set from "For IT" sheet
    const supervisedEmpSet = new Set();
    const csuiteEmpSet = new Set();

    try {
      const itSheet = getSheet('For IT');
      if (itSheet) {
        const itRows = sheetToJson(itSheet);
        itRows.forEach(r => {
          const empId = String(r['Employee No'] || r.EmployeeNo || r.ID || r.EmployeeID || r['Employee ID'] || '').toLowerCase().trim();
          const empName = String(r.Name || r.EmployeeName || r['Employee Name'] || '').toLowerCase().trim();
          const assignedHod = String(r.HOD || r.HODName || r.HodName || r.Manager || r.ReportTo || r['HOD Name'] || '').toLowerCase().trim();
          const assignedCs = String(r.Csuite || r.CsuiteName || r.CSuiteName || r.CsuiteEmail || r['C-Suite'] || '').toLowerCase().trim();

          if (assignedHod && (
            (hodNameClean && (assignedHod.includes(hodNameClean) || hodNameClean.includes(assignedHod))) ||
            (hodIdClean && assignedHod.includes(hodIdClean)) ||
            (hodEmailClean && assignedHod.includes(hodEmailClean))
          )) {
            if (empId) supervisedEmpSet.add(empId);
            if (empName) supervisedEmpSet.add(empName);
          }

          if (assignedCs && (
            (hodNameClean && (assignedCs.includes(hodNameClean) || hodNameClean.includes(assignedCs))) ||
            (hodIdClean && assignedCs.includes(hodIdClean)) ||
            (hodEmailClean && assignedCs.includes(hodEmailClean))
          )) {
            if (empId) csuiteEmpSet.add(empId);
            if (empName) csuiteEmpSet.add(empName);
          }
        });
      }
    } catch(e) {}

    // 2. Build supervised employee set from "Employees" sheet
    try {
      const empSheet = getSheet('Employees');
      if (empSheet) {
        const empRows = sheetToJson(empSheet);
        empRows.forEach(e => {
          const empId = String(e['Employee No'] || e.EmployeeNo || e.ID || e.EmployeeID || '').toLowerCase().trim();
          const empName = String(e.Name || e.EmployeeName || e['Employee Name'] || '').toLowerCase().trim();
          const assignedHod = String(e.HOD || e.HODName || e.HodName || e.Manager || e.ReportTo || '').toLowerCase().trim();
          const empDept = String(e.Department || e.CostCentre || e['Cost Centre'] || '').toLowerCase().trim();

          if (assignedHod && (
            (hodNameClean && (assignedHod.includes(hodNameClean) || hodNameClean.includes(assignedHod))) ||
            (hodIdClean && assignedHod.includes(hodIdClean)) ||
            (hodEmailClean && assignedHod.includes(hodEmailClean))
          )) {
            if (empId) supervisedEmpSet.add(empId);
            if (empName) supervisedEmpSet.add(empName);
          } else if (hodCcClean && empDept && (empDept.includes(hodCcClean) || hodCcClean.includes(empDept))) {
            if (empId) supervisedEmpSet.add(empId);
            if (empName) supervisedEmpSet.add(empName);
          }
        });
      }
    } catch(e) {}

    // 3. Filter trainings under supervision or multi-stage queue
    return trainings.filter(t => {
      const reqId = String(t.RequestedBy || t.EmployeeID || '').toLowerCase().trim();
      const reqName = String(t.RequestedByName || '').toLowerCase().trim();
      const tDept = String(t.Department || '').toLowerCase().trim();
      const appStatus = String(t.ApprovalStatus || t.Status || '').toLowerCase();

      // C-Suite Role Matching
      if (isCsuite) {
        if (appStatus.includes('c-suite')) return true;
        if (reqId && csuiteEmpSet.has(reqId)) return true;
        if (reqName && csuiteEmpSet.has(reqName)) return true;
      }

      // HOHR Role Matching
      if (isHohr) {
        if (appStatus.includes('hohr') || appStatus.includes('hr')) return true;
      }

      // Direct HOD Supervision via assigned employee match
      if (reqId && supervisedEmpSet.has(reqId)) return true;
      if (reqName && supervisedEmpSet.has(reqName)) return true;

      // Matching Cost Centre / Department
      if (hodCcClean && tDept) {
        const deptCodeMatch = tDept.match(/\d+/);
        const hodCodeMatch = hodCcClean.match(/\d+/);
        if (deptCodeMatch && hodCodeMatch && deptCodeMatch[0] === hodCodeMatch[0]) return true;
        if (tDept.includes(hodCcClean) || hodCcClean.includes(tDept)) return true;
      }

      return false;
    });
  } catch(e) {
    Logger.log('getHODSupervisedRequisitions error: ' + e.message);
    return [];
  }
}


/**
 * API: Submit HOD Decision (Approve / Reject / Postpone / Reschedule)
 */
/**
 * API: Submit HOD Decision (Approve / Reject / Return)
 */
function submitHODDecision(data) {
  try {
    if (!data || !data.trainingId || !data.decision) {
      return err('Training ID and decision are required.');
    }

    const cleanId = String(data.trainingId).trim();
    const activeEmail = resolveActiveSessionEmail(data.userEmail || data.email);
    const auth = validateHODAccess(activeEmail);
    if (!auth.valid) return err(auth.message);

    // Verify server-side authorization: request must be in approver's supervised queue
    const supervisedQueue = getHODSupervisedRequisitions(auth.hod);
    const isAuthorizedRequest = supervisedQueue.some(s => getTrainingId(s).toLowerCase() === cleanId.toLowerCase());
    if (!isAuthorizedRequest) {
      return err(`Unauthorized Access: Training request (${cleanId}) is not assigned to your approval queue.`);
    }

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');


    const row = findRowById(tSheet, cleanId);
    if (row === -1) return err(`Training request (${cleanId}) not found in database.`);

    const headers = tSheet.getDataRange().getValues()[0].map(h => String(h).trim());

    const timestamp = formatDateTime(new Date());

    const updateCol = (colName, val) => {
      const idx = headers.indexOf(colName) + 1;
      if (idx > 0) {
        tSheet.getRange(row, idx).setValue(val);
      }
    };

    let rawDecision = String(data.decision).trim();
    let validDecision = 'Approved';
    if (rawDecision.toLowerCase().includes('reject')) validDecision = 'Rejected';
    else if (rawDecision.toLowerCase().includes('return')) validDecision = 'Returned';
    else validDecision = 'Approved';

    const hodId = data.hodId || (auth.hod ? auth.hod.ID : 'HOD-UNKNOWN');
    const hodName = data.hodName || (auth.hod ? auth.hod.Name : 'HOD Approver');
    const hodCostCentre = data.hodCostCentre || (auth.hod ? auth.hod.CostCentre : 'Cost Centre');
    const remarks = data.remarks || '';

    // Fetch requester and assigned approvers from employee database & directory tabs
    const trainingsList = sheetToJson(tSheet);
    const currentT = trainingsList.find(t => String(t.ID || '').toLowerCase() === cleanId.toLowerCase()) || {};

    const requesterId = currentT.RequestedBy || currentT['Requested By'] || '';
    const requesterName = currentT.RequestedByName || currentT['Requested By Name'] || requesterId || 'Employee Requester';
    const requesterCostCentre = currentT.CostCentre || currentT.Department || hodCostCentre;


    const csProfile = resolveCSuiteProfileForRequester(requesterCostCentre, requesterId || requesterName);
    const csuiteName = csProfile ? csProfile.Name : (auth.hod.CsuiteName || '');
    const csuiteId = csProfile ? csProfile.ID : '';

    const hrProfile = getHOHREmailProfile();
    const hohrName = hrProfile ? hrProfile.Name : (auth.hod.HohrName || '');
    const hohrId = hrProfile ? hrProfile.ID : '';

    // Identity comparison helper (matches by ID or Name)
    const isSamePerson = (n1, id1, n2, id2) => {
      const clean1 = String(n1 || '').toLowerCase().trim();
      const clean2 = String(n2 || '').toLowerCase().trim();
      const cleanId1 = String(id1 || '').toLowerCase().trim();
      const cleanId2 = String(id2 || '').toLowerCase().trim();
      if (!clean1 && !cleanId1) return false;
      if (!clean2 && !cleanId2) return false;
      if (cleanId1 && cleanId2 && cleanId1 === cleanId2) return true;
      if (clean1 && clean2 && (clean1 === clean2 || clean1.includes(clean2) || clean2.includes(clean1))) return true;
      return false;
    };

    const reqIsHod = isSamePerson(requesterName, requesterId, hodName, hodId);
    const hodIsCsuite = isSamePerson(hodName, hodId, csuiteName, csuiteId);
    const csuiteIsHohr = isSamePerson(csuiteName, csuiteId, hohrName, hohrId);
    const hodIsHohr = isSamePerson(hodName, hodId, hohrName, hohrId);

    const currentAppStatus = String(tSheet.getRange(row, headers.indexOf('ApprovalStatus') + 1).getValue() || 'Pending HOD Approval');
    const userIsCsuite = auth.hod && auth.hod.isCsuite;
    const userIsHohr = auth.hod && auth.hod.isHohr;

    let nextApprovalStatus = validDecision;
    let approvalStep = 'HOD';
    if (currentAppStatus.includes('C-Suite')) {
      approvalStep = 'Csuite';
    } else if (currentAppStatus.includes('HOHR')) {
      approvalStep = 'HOHR';
    } else if (currentAppStatus.includes('HOD') && userIsCsuite && !auth.hod.isHod) {
      approvalStep = 'Csuite';
    }

    let autoVerifiedRemarks = '';

    if (validDecision === 'Approved') {
      if (approvalStep === 'HOD') {
        if ((hodIsCsuite || userIsCsuite) && (csuiteIsHohr || hodIsHohr || userIsHohr)) {
          // HOD == C-Suite == HOHR -> Directly verify all 3 levels -> Fully Approved!
          nextApprovalStatus = 'Approved';
          autoVerifiedRemarks = ' (Auto-verified: HOD, C-Suite, and HOHR stage approved)';
        } else if (hodIsCsuite || userIsCsuite) {
          // HOD == C-Suite -> Directly verify both (HOD & C-Suite) -> Pending HOHR Approval!
          nextApprovalStatus = 'Pending HOHR Approval';
          autoVerifiedRemarks = ' (Auto-verified: HOD and C-Suite stage approved)';
        } else {
          // Standard: HOD Approved -> Pending C-Suite Approval
          nextApprovalStatus = 'Pending C-Suite Approval';
        }
      } else if (approvalStep === 'Csuite') {
        if (csuiteIsHohr || hodIsHohr || userIsHohr) {
          // C-Suite == HOHR -> Directly verify both C-Suite & HOHR -> Fully Approved!
          nextApprovalStatus = 'Approved';
          if (csuiteIsHohr || hodIsHohr) autoVerifiedRemarks = ' (Auto-verified: C-Suite and HOHR are the same approver)';
        } else {
          nextApprovalStatus = 'Pending HOHR Approval';
        }
      } else if (approvalStep === 'HOHR') {
        nextApprovalStatus = 'Approved';
      }
    } else if (validDecision === 'Rejected') {
      nextApprovalStatus = 'Rejected';
    } else if (validDecision === 'Returned') {
      nextApprovalStatus = 'Returned';
    }

    const fullRemarks = (remarks || '') + autoVerifiedRemarks;

    updateCol('ApprovalStatus', nextApprovalStatus);
    updateCol('ApprovedBy', `${hodName} (${hodId})`);
    updateCol('ApprovedCostCentre', hodCostCentre);
    updateCol('ApprovedAt', timestamp);
    updateCol('ApprovalRemarks', fullRemarks.trim());

    // Update Status & Stage based on decision
    if (nextApprovalStatus === 'Approved') {
      updateCol('Status', 'Upcoming');
      updateCol('Stage', 'Created');
    } else if (validDecision === 'Rejected') {
      updateCol('Status', 'Cancelled');
      updateCol('Stage', 'Programme Closed');
    } else if (validDecision === 'Returned') {
      updateCol('Status', 'Returned');
      updateCol('Stage', 'Requisition Draft');
    }

    updateCol('UpdatedDate', timestamp);
    SpreadsheetApp.flush();

    // Automatically update AP-HRD-F01-01 Training Requisition Form Google Sheet digital approval signatures
    let hodPosition = auth.hod ? (auth.hod.Position || auth.hod.JobTitle || auth.hod.PositionTitle || auth.hod.RoleTitle || '') : '';
    if (!hodPosition && getSheet('Employees')) {
      const emps = sheetToJson(getSheet('Employees'));
      const empMatch = emps.find(e => String(e.ID).toLowerCase() === String(hodId).toLowerCase());
      if (empMatch) hodPosition = empMatch.Position || empMatch.JobTitle || empMatch.PositionTitle || '';
    }
    if (!hodPosition) {
      hodPosition = approvalStep === 'Csuite' ? 'C-Suite Executive' : (approvalStep === 'HOHR' ? 'Head of HR' : 'Head of Department');
    }

    try {
      // Stamp active approval step signature
      updateTrainingRequisitionSignatures(cleanId, approvalStep, {
        status: (validDecision === 'Approved' ? (approvalStep === 'HOD' ? 'Verified' : 'Approved') : validDecision),
        employeeNo: hodId,
        name: hodName,
        position: hodPosition,
        date: timestamp
      });

      // If C-Suite approved when status was HOD stage, also stamp HOD signature as Verified
      if (validDecision === 'Approved' && approvalStep === 'Csuite' && currentAppStatus.includes('HOD')) {
        updateTrainingRequisitionSignatures(cleanId, 'HOD', {
          status: 'Verified',
          employeeNo: hodId,
          name: hodName,
          position: hodPosition,
          date: timestamp
        });
      }

      // If HOD == C-Suite, also stamp C-Suite signature automatically
      if (validDecision === 'Approved' && (hodIsCsuite || userIsCsuite) && approvalStep === 'HOD') {
        updateTrainingRequisitionSignatures(cleanId, 'Csuite', {
          status: 'Approved',
          employeeNo: csuiteId || hodId,
          name: csuiteName || hodName,
          position: csProfile ? csProfile.Position : 'C-Suite Executive',
          date: timestamp
        });
      }

      // If HOD == C-Suite == HOHR or C-Suite == HOHR, also stamp HOHR signature automatically
      if (validDecision === 'Approved' && (hodIsHohr || csuiteIsHohr || userIsHohr) && nextApprovalStatus === 'Approved') {
        updateTrainingRequisitionSignatures(cleanId, 'HOHR', {
          status: 'Approved',
          employeeNo: hohrId || hodId,
          name: hohrName || hodName,
          position: hrProfile ? hrProfile.Position : 'Head of HR',
          date: timestamp
        });
      }
    } catch(sigErr) {
      Logger.log('Signature update error in submitHODDecision: ' + sigErr.message);
    }


    // Send email notifications via Gmail Drafts (development mode)
    let draftLog = [];
    try {
      const trainings = sheetToJson(tSheet);
      const currentT = trainings.find(t =>
        String(t.ID || '').toLowerCase() === cleanId.toLowerCase() ||
        String(t.Code || '').toLowerCase() === cleanId.toLowerCase()
      ) || {};
      const trainingName = currentT.Name || currentT.TrainingTitle || cleanId;
      const trainingCode = currentT.Code || currentT.ID || cleanId;

      let requesterEmail = currentT.RequestedByEmail || currentT['Requested By Email'] || currentT.Email || currentT['Email Address'] || currentT.UserEmail || '';
      const requesterId = currentT.RequestedBy || currentT['Requested By'] || currentT.EmployeeID || currentT['Employee ID'] || '';
      let requesterName = currentT.RequestedByName || currentT['Requested By Name'] || currentT.Trainer || requesterId || 'Employee Requester';

      if (!requesterEmail && getSheet('Employees')) {
        const emps = sheetToJson(getSheet('Employees'));
        const cleanReqId = String(requesterId).toLowerCase().trim();
        const cleanReqName = String(requesterName).toLowerCase().trim();

        const empMatch = emps.find(e => {
          const eId = String(e['Employee No'] || e.EmployeeNo || e.ID || e.EmployeeID || e['Employee ID'] || '').toLowerCase().trim();
          const eName = String(e.Name || e.EmployeeName || e['Employee Name'] || '').toLowerCase().trim();
          return (cleanReqId && eId === cleanReqId) || (cleanReqName && eName === cleanReqName);
        });

        if (empMatch) {
          requesterEmail = String(empMatch.Email || empMatch.EmailAddress || empMatch['Email Address'] || '').trim();
          if (!requesterName || requesterName === requesterId) {
            requesterName = String(empMatch.Name || empMatch.EmployeeName || empMatch['Employee Name'] || requesterName).trim();
          }
        }
      }

      // Robust Fallback: If requester email is still empty, fallback to HOD's email or active user email
      if (!requesterEmail) {
        requesterEmail = auth.hod.Email || Session.getActiveUser().getEmail() || 'requester@apollofood.com.my';
      }

      const hodPortalUrl = getConfigProperty('HOD_PORTAL_URL', '');
      const reviewUrl = hodPortalUrl ? `${hodPortalUrl}?page=review&id=${cleanId}` : getAppUrl();

      const proposedDateStr = (currentT.EndDate && currentT.EndDate !== currentT.StartDate)
        ? `${formatDateForEmail(currentT.StartDate)} – ${formatDateForEmail(currentT.EndDate)}`
        : formatDateForEmail(currentT.StartDate);
      const durationStr = formatDurationForEmail(currentT.Duration || 1, currentT.TotalHours || 8);
      const feeStr = formatFeeForEmail(currentT.CourseFee);
      const categoryStr = currentT.Category || 'General';

      // IF REJECTED OR RETURNED: Email the requester only
      if (validDecision === 'Rejected' || validDecision === 'Returned') {
        const reqSubject = `Training Requisition — ${trainingName} | ${cleanId}`;
        const reqBody = `Dear ${requesterName},\n\n` +
          `Your Training Requisition Request for "${trainingName}" (${cleanId}) has been ${validDecision.toUpperCase()}.\n\n` +
          `Decision Status: ${validDecision.toUpperCase()}\n` +
          `Reviewed By: ${hodName} (${hodId})\n` +
          `Position: ${hodPosition}\n` +
          `Department / Cost Centre: ${hodCostCentre}\n` +
          `Timestamp: ${timestamp}\n` +
          `Remarks / Reason: ${remarks || 'No remarks provided.'}\n\n` +
          (validDecision === 'Returned'
            ? 'Please review the remarks above, make the necessary amendments, and re-submit your requisition.\n\n'
            : 'This requisition request has been closed.\n\n') +
          `Thank you,\nTrainHub Training Management System`;

        const reqHtml = buildTrainingRequisitionEmailHtml({
          requestId: cleanId,
          trainingTitle: trainingName,
          requesterName: requesterName,
          employeeId: requesterId,
          department: requesterCostCentre,
          category: categoryStr,
          proposedDate: proposedDateStr,
          duration: durationStr,
          estimatedFee: feeStr,
          status: validDecision,
          reviewUrl: validDecision === 'Returned' ? reviewUrl : '',
          badgeText: validDecision.toUpperCase(),
          headlineText: `Training Requisition ${validDecision}`,
          greetingText: `Dear ${requesterName},`,
          introText: `Your Training Requisition Request for "${trainingName}" (${cleanId}) has been ${validDecision.toLowerCase()} by ${hodName}.${remarks ? ' Remarks: ' + remarks : ''}`
        });

        sendOrDraftEmail(requesterEmail, reqSubject, reqBody, 'requester', draftLog, reqHtml);
      }

      // IF APPROVED: Email the next approver (or Requester + Arina if final complete approval)
      else if (validDecision === 'Approved') {
        // Step 1: HOD Approved -> Pending C-Suite Approval (Only notify C-Suite Executive)
        if (nextApprovalStatus === 'Pending C-Suite Approval') {
          const requesterCostCentre = currentT.CostCentre || currentT.Department || currentT['Cost Centre'] || currentT['CostCentre'] || hodCostCentre;
          const csProfile = resolveCSuiteProfileForRequester(requesterCostCentre, requesterId);
          const csuiteEmail = csProfile.Email || auth.hod.CsuiteEmail || getConfigProperty('CSUITE_EMAIL', '') || auth.hod.Email;
          const csuiteName = csProfile.Name || auth.hod.CsuiteName || 'C-Suite Executive';

          const csSubject = `Training Requisition — ${trainingName} | ${cleanId}`;
          const csBody = `Dear ${csuiteName},\n\n` +
            `A Training Requisition Request for "${trainingName}" (${cleanId}) submitted by ${requesterName} (${requesterId}) has been APPROVED by Head of Department (${hodName}) and requires your managerial approval.\n\n` +
            `Training Name: ${trainingName}\n` +
            `Requester: ${requesterName} (${requesterId})\n` +
            `Requester Department / Cost Centre: ${requesterCostCentre}\n` +
            `HOD Approved By: ${hodName} (${hodId})\n` +
            `Date Approved: ${timestamp}\n` +
            `HOD Remarks: ${remarks || 'Approved by HOD.'}\n\n` +
            `Please review and issue your digital approval in the TrainHub HOD Portal:\n${reviewUrl}\n\n` +
            `Thank you,\nTrainHub Training Management System`;

          const csHtml = buildTrainingRequisitionEmailHtml({
            requestId: cleanId,
            trainingTitle: trainingName,
            requesterName: requesterName,
            employeeId: requesterId,
            department: requesterCostCentre,
            category: categoryStr,
            proposedDate: proposedDateStr,
            duration: durationStr,
            estimatedFee: feeStr,
            status: 'Pending C-Suite Approval',
            reviewUrl: reviewUrl,
            badgeText: 'ACTION REQUIRED',
            headlineText: 'Training Requisition Requires Your Review',
            greetingText: `Dear ${csuiteName},`,
            introText: `A Training Requisition Request for "${trainingName}" (${cleanId}) submitted by ${requesterName} (${requesterId}) has been APPROVED by Head of Department (${hodName}) and is currently awaiting your review and approval.`
          });

          sendOrDraftEmail(csuiteEmail, csSubject, csBody, `C-Suite [${csuiteName}]`, draftLog, csHtml);
        }

        // Step 2: C-Suite Approved -> Pending HOHR Approval (Only notify Head of HR)
        else if (nextApprovalStatus === 'Pending HOHR Approval') {
          const hrProfile = getHOHREmailProfile();
          const hohrEmail = hrProfile.Email || auth.hod.HohrEmail || getConfigProperty('HOHR_EMAIL', '') || 'hohr@apollofood.com.my';
          const hohrName = hrProfile.Name || auth.hod.HohrName || 'Head of HR';

          const hrSubject = `Training Requisition — ${trainingName} | ${cleanId}`;
          const hrBody = `Dear ${hohrName},\n\n` +
            `A Training Requisition Request for "${trainingName}" (${cleanId}) has received C-Suite approval and now requires final Head of HR approval.\n\n` +
            `Training Name: ${trainingName}\n` +
            `Requester: ${requesterName} (${requesterId})\n` +
            `C-Suite Approver: ${hodName} (${hodId})\n` +
            `Department / Cost Centre: ${hodCostCentre}\n` +
            `Date Approved: ${timestamp}\n` +
            `Remarks: ${remarks || 'Approved by C-Suite.'}\n\n` +
            `Please review and issue your digital approval in the TrainHub HOD Portal:\n${reviewUrl}\n\n` +
            `Thank you,\nTrainHub Training Management System`;

          const hrHtml = buildTrainingRequisitionEmailHtml({
            requestId: cleanId,
            trainingTitle: trainingName,
            requesterName: requesterName,
            employeeId: requesterId,
            department: requesterCostCentre,
            category: categoryStr,
            proposedDate: proposedDateStr,
            duration: durationStr,
            estimatedFee: feeStr,
            status: 'Pending HOHR Approval',
            reviewUrl: reviewUrl,
            badgeText: 'ACTION REQUIRED',
            headlineText: 'Training Requisition Requires Your Review',
            greetingText: `Dear ${hohrName},`,
            introText: `A Training Requisition Request for "${trainingName}" (${cleanId}) submitted by ${requesterName} (${requesterId}) has received C-Suite approval and is currently awaiting your final Head of HR approval.`
          });

          sendOrDraftEmail(hohrEmail, hrSubject, hrBody, `HOHR [${hohrName}]`, draftLog, hrHtml);
        }

        // Step 3: HOHR Approved -> Fully Approved (Complete Approval: Notify Requester + Arina)
        else if (nextApprovalStatus === 'Approved') {
          // Complete Approval Notification to Requester
          const reqSubject = `Training Requisition — ${trainingName} | ${cleanId}`;
          const reqBody = `Dear ${requesterName},\n\n` +
            `Great news! Your Training Requisition Request for "${trainingName}" (${cleanId}) has received ALL required managerial approvals (HOD, C-Suite, and Head of HR).\n\n` +
            `Decision Status: APPROVED\n` +
            `Final Approver (HOHR): ${hodName} (${hodId})\n` +
            `Department / Cost Centre: ${hodCostCentre}\n` +
            `Timestamp: ${timestamp}\n` +
            `Remarks: ${remarks || 'Fully Approved.'}\n\n` +
            `The HR Training Administrator will now schedule training sessions and generate attendance links.\n\n` +
            `Thank you,\nTrainHub Training Management System`;

          const reqHtml = buildTrainingRequisitionEmailHtml({
            requestId: cleanId,
            trainingTitle: trainingName,
            requesterName: requesterName,
            employeeId: requesterId,
            department: requesterCostCentre,
            category: categoryStr,
            proposedDate: proposedDateStr,
            duration: durationStr,
            estimatedFee: feeStr,
            status: 'Approved',
            reviewUrl: '',
            badgeText: 'APPROVED',
            headlineText: 'Training Requisition Fully Approved!',
            greetingText: `Dear ${requesterName},`,
            introText: `Great news! Your Training Requisition Request for "${trainingName}" (${cleanId}) has received ALL required managerial approvals (HOD, C-Suite, and Head of HR).`
          });

          sendOrDraftEmail(requesterEmail, reqSubject, reqBody, 'requester', draftLog, reqHtml);

          // Notification to Arina (HR Admin)
          const arinaSubject = `Training Requisition — ${trainingName} | ${cleanId}`;
          const arinaBody = `Dear Arina,\n\n` +
            `The following Training Requisition has received all required approvals (HOD, C-Suite, HOHR):\n\n` +
            `Training Name: ${trainingName} (${cleanId})\n` +
            `Requester: ${requesterName} (${requesterId})\n` +
            `HOHR Approved By: ${hodName} (${hodId})\n` +
            `Cost Centre: ${hodCostCentre}\n` +
            `Date Approved: ${timestamp}\n\n` +
            `The Admin System can now proceed with session creation, QR code generation, and participant attendance tracking.\n\n` +
            `Thank you,\nTrainHub Training Management System`;

          const arinaHtml = buildTrainingRequisitionEmailHtml({
            requestId: cleanId,
            trainingTitle: trainingName,
            requesterName: requesterName,
            employeeId: requesterId,
            department: requesterCostCentre,
            category: categoryStr,
            proposedDate: proposedDateStr,
            duration: durationStr,
            estimatedFee: feeStr,
            status: 'Approved',
            reviewUrl: '',
            badgeText: 'ACTION REQUIRED',
            headlineText: 'Training Requisition Fully Approved & Ready for Session Setup',
            greetingText: 'Dear Arina,',
            introText: `The Training Requisition for "${trainingName}" (${cleanId}) submitted by ${requesterName} (${requesterId}) has received all required approvals (HOD, C-Suite, HOHR) and is ready for session setup.`
          });

          sendOrDraftEmail('arina.ismail@apollofood.com.my', arinaSubject, arinaBody, 'Arina (HR Admin)', draftLog, arinaHtml);
        }
      }
    } catch (mailErr) {
      Logger.log('Notification mail error: ' + mailErr.message);
      draftLog.push(`General mail engine error: ${mailErr.message}`);
    }

    try {
      syncTrainingById(cleanId, 'HOD Decision (' + nextApprovalStatus + ')', 'STATUS_CHANGE');
    } catch(syncErr) {
      Logger.log('syncTrainingById error in submitHODDecision: ' + syncErr.message);
    }

    return ok({
      trainingId: cleanId,
      decision: nextApprovalStatus,
      approvedBy: hodName,
      timestamp: timestamp,
      draftLog: draftLog,
      message: `Training requisition request marked as ${nextApprovalStatus.toUpperCase()}.${logSummary}`
    });
  } catch (e) {
    return err('Failed to process HOD decision: ' + e.message);
  }
}

/**
 * Helper function to create a Gmail Draft strictly (Drafts Only)
 */
function sendOrDraftEmail(recipient, subject, body, logPrefix, draftLog, htmlBody) {
  if (!recipient) return;
  try {
    const options = {};
    if (htmlBody) options.htmlBody = htmlBody;
    GmailApp.createDraft(recipient, subject, body, options);
    if (draftLog) draftLog.push(`Draft created for ${logPrefix} (${recipient})`);
  } catch (dErr) {
    Logger.log('GmailApp createDraft error: ' + dErr.message);
    if (draftLog) draftLog.push(`Draft error for ${logPrefix} (${recipient}): ${dErr.message}`);
  }
}

/**
 * Run this function ONCE in Apps Script Editor to grant Gmail Draft permissions
 */
function authorizeGmailDrafts() {
  const user = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'admin@apollofood.com.my';
  const draft = GmailApp.createDraft(user, '[TrainHub Draft Test]', 'Gmail Draft permission verified successfully.');
  Logger.log('Gmail Draft authorized successfully. Draft ID: ' + draft.getId());
  return 'Gmail Draft permissions authorized successfully! Draft ID: ' + draft.getId();
}



/**
 * API: Fetch participants pending 3-Month Post Evaluation for a training under this HOD
 */
function getPendingPostEvalParticipants(trainingId, userEmail) {
  try {
    const auth = validateHODAccess(userEmail);
    if (!auth.valid) return err(auth.message);

    const targetHodCostCentre = auth.hod.CostCentre;

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const trainings = sheetToJson(tSheet);

    const candidateTrainings = trainings.filter(t => {
      const st = String(t.Stage || t.Status || t.ApprovalStatus || '').toLowerCase();
      return st.includes('3-month') || st.includes('6-month') || st.includes('completed') || st.includes('approved') || st.includes('ongoing');
    });

    const activeList = candidateTrainings.length > 0 ? candidateTrainings : trainings;
    let training = null;

    if (trainingId) {
      const cleanId = String(trainingId).trim().toLowerCase();
      training = activeList.find(t => {
        const tid = getTrainingId(t).toLowerCase();
        const tcode = getTrainingCode(t).toLowerCase();
        return tid === cleanId || tcode === cleanId;
      });
      if (!training) {
        training = trainings.find(t => {
          const tid = getTrainingId(t).toLowerCase();
          const tcode = getTrainingCode(t).toLowerCase();
          return tid === cleanId || tcode === cleanId;
        });
      }
    }

    if (!training && activeList.length > 0) {
      training = activeList[0];
    }

    if (!training) return err('No training programme found for post evaluation.');

    // Fetch all participants enrolled in this training
    const targetTId = getTrainingId(training);
    const ss = getTrainingDataSpreadsheet(targetTId);
    const partSheet = ss ? ss.getSheetByName('TrainingParticipants') : null;
    let allParticipants = partSheet ? sheetToJson(partSheet) : [];

    // Filter participants under this HOD's Cost Centre / Department
    let hodParticipants = allParticipants.filter(p => {
      const dept = String(p.Department || p.CostCentre || '').toLowerCase();
      const hodCc = String(targetHodCostCentre).toLowerCase();
      return dept.includes(hodCc) || hodCc.includes(dept) || targetHodCostCentre === 'ALL';
    });

    // Check which participants have already been evaluated in PostEval sheet
    const postSheet = ss ? ss.getSheetByName('PostEval') : null;
    let evaluatedEmpIds = [];
    if (postSheet) {
      const postRows = sheetToJson(postSheet);
      evaluatedEmpIds = postRows.map(pr => String(pr.EmployeeID || '').trim().toLowerCase());
    }

    const pendingList = hodParticipants.filter(p => !evaluatedEmpIds.includes(String(p.EmployeeID || p.ID || '').trim().toLowerCase()));
    const completedList = hodParticipants.filter(p => evaluatedEmpIds.includes(String(p.EmployeeID || p.ID || '').trim().toLowerCase()));

    return ok({
      training: training,
      hod: auth.hod,
      totalUnderHod: hodParticipants.length,
      pendingCount: pendingList.length,
      completedCount: completedList.length,
      pendingParticipants: pendingList,
      completedParticipants: completedList,
      postEvalTrainings: activeList
    });
  } catch (e) {
    return err('Failed to load post evaluation participants: ' + e.message);
  }
}

/**
 * API: Submit HOD Post Evaluation for a participant
 */
function submitHODPostEval(data) {
  try {
    if (!data || !data.trainingId || !data.employeeId) {
      return err('Training ID and Employee ID are required.');
    }

    const userEmail = data.userEmail || data.email || '';
    const auth = validateHODAccess(userEmail);
    if (!auth.valid) return err(auth.message);

    const cleanTId = String(data.trainingId).trim();
    const cleanEmpId = String(data.employeeId).trim();

    const ss = getTrainingDataSpreadsheet(cleanTId);
    if (!ss) return err('Could not open per-training sheet for ID: ' + cleanTId);

    let postSheet = ss.getSheetByName('PostEval');
    if (!postSheet) {
      postSheet = ss.insertSheet('PostEval');
      postSheet.appendRow(['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID', 'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply', 'FurtherTraining', 'Comments', 'SubmittedAt']);
      postSheet.getRange('A1:L1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      postSheet.setFrozenRows(1);
    }

    // Check for duplicate submission
    const existingRows = sheetToJson(postSheet);
    const duplicate = existingRows.find(r =>
      String(r.EmployeeID || '').trim().toLowerCase() === cleanEmpId.toLowerCase()
    );

    if (duplicate) {
      return err(`Post evaluation for employee (${cleanEmpId}) has already been submitted.`);
    }

    const evalId = 'PEVAL-' + Math.floor(100000 + Math.random() * 900000);
    const evaluatorName = auth.hod.Name;
    const evaluatorId = auth.hod.ID;
    const submittedAt = formatDateTime(new Date());

    const rowData = [
      evalId,
      cleanTId,
      cleanEmpId,
      evaluatorName,
      evaluatorId,
      data.competencyBefore || '3 - Proficient',
      data.competencyAfter  || '5 - Expert',
      data.improvement       || 'High',
      data.canApply          || 'Yes',
      data.furtherTraining   || 'None required',
      data.comments          || 'Good progress demonstrated in work output.',
      submittedAt
    ];

    postSheet.appendRow(rowData);
    SpreadsheetApp.flush();

    // Return updated pending list for this HOD
    return getPendingPostEvalParticipants(cleanTId, auth.hod.Email || userEmail);
  } catch (e) {
    return err('Failed to submit post evaluation: ' + e.message);
  }
}

