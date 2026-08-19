/**
 * Code.gs — HOD Portal Web App Entry Point & Server Engine
 */

/**
 * Code.gs — HOD Portal Web App Entry Point & Server Engine
 */

function doGet(e) {
  const pageParam = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase().trim() : 'review';

  const pageMap = {
    'review':   'HodReview',
    'posteval': 'HodPostEvaluation',
    'success':  'Success',
    'error':    'Error'
  };

  const templateName = pageMap[pageParam] || 'HodReview';
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — Approval Portal');

  // Determine user email strictly from authenticated Google Workspace session
  let activeEmail = resolveActiveSessionEmail();

  let systemLogoUrl = '';
  try {
    const rawSysLogo = getSystemLogoUrl() || '';
    systemLogoUrl = rawSysLogo ? (convertDriveLinkToDirectImageUrl(rawSysLogo) || rawSysLogo) : '';
  } catch(sysLogoErr) {}

  // Validate approver authorization
  const auth = validateHODAccess(activeEmail);
  if (!auth.valid) {
    try {
      const errTemplate = HtmlService.createTemplateFromFile('Error');
      errTemplate.systemLogoUrl = String(systemLogoUrl);
      errTemplate.params = { 
        title: 'Access Denied — Unauthorized Account',
        message: auth.message || (activeEmail ? `Access Denied: The email address (${activeEmail}) is not registered as an authorized approver.` : 'No authenticated Google account detected. Please log in with your authorized company email.')
      };
      return errTemplate.evaluate()
        .setTitle('Access Denied — TrainHub Approval Portal')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {}
  }

  try {
    const template = HtmlService.createTemplateFromFile(templateName);
    template.params = (e && e.parameter) ? e.parameter : {};
    template.page = pageParam;
    template.activeEmail = activeEmail || '';
    template.systemLogoUrl = String(systemLogoUrl);

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
function resolveActiveSessionEmail() {
  let sessionEmail = '';
  try {
    sessionEmail = Session.getActiveUser().getEmail();
  } catch (e) {}
  return String(sessionEmail || '').trim();
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
/**
 * API: Fetch complete dashboard data for the authenticated approver.
 * SECURITY ENFORCEMENT: Resolves identity server-side.
 * ENFORCES CURRENT ACTIVE APPROVAL STAGE & ASSIGNED APPROVER MATCHING.
 */
function getApproverDashboardData() {
  try {
    const auth = resolveAuthenticatedUserServerSide();
    if (!auth.valid || !auth.hod) {
      return err(auth.message || 'Unable to identify your Google account. Please ensure you are logged into your authorized company account.');
    }

    const hodProfile = auth.hod;
    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const allRequisitions = sheetToJson(tSheet);

    let pendingHodCount = 0;
    let pendingCsuiteCount = 0;
    let pendingHohrCount = 0;
    let approvedCount = 0;
    let rejectedCount = 0;

    const pendingRequests = [];
    const historyRequests = [];
    const allSupervisedRequests = [];

    allRequisitions.forEach(r => {
      const activeStage = getCurrentActiveApprovalStage(r);
      const assignedApprovers = getAssignedApproversForRequisition(r);

      const statusStr = String(r.ApprovalStatus || r.Status || 'Pending HOD Approval').trim();
      const hodSt = String(r.HODStatus || 'Pending').trim();
      const csSt = String(r.CsuiteStatus || 'N/A').trim();
      const hrSt = String(r.HOHRStatus || 'N/A').trim();

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
        ActiveStage: activeStage,
        DisplayStatus: activeStage !== 'NONE' ? `Pending ${activeStage === 'CSuite' ? 'C-Suite' : activeStage} Approval` : statusStr,
        HODStatus: hodSt,
        CsuiteStatus: csSt,
        HOHRStatus: hrSt,
        HOD: assignedApprovers.HOD ? assignedApprovers.HOD.Name : hodSt,
        Csuite: assignedApprovers.CSuite ? assignedApprovers.CSuite.Name : csSt,
        HOHR: assignedApprovers.HOHR ? assignedApprovers.HOHR.Name : hrSt,
        ExpiryDate: String(r.ExpiryDate || r.CertExpiryDate || r.CertificateExpiryDate || '').trim(),
        CertExpiryDate: String(r.CertExpiryDate || r.ExpiryDate || r.CertificateExpiryDate || '').trim(),
        ApprovalRemarks: String(r.ApprovalRemarks || '').trim(),
        ApprovedBy: String(r.ApprovedBy || '').trim(),
        ApprovedAt: String(r.ApprovedAt || '').trim(),
        ApprovedCostCentre: String(r.ApprovedCostCentre || '').trim(),
        BrochureURL: String(r.BrochureURL || r.BrochureUrl || '').trim()
      };

      const isAssignedHod = isSameApprover(hodProfile, assignedApprovers.HOD);
      const isAssignedCs = isSameApprover(hodProfile, assignedApprovers.CSuite);
      const isAssignedHr = isSameApprover(hodProfile, assignedApprovers.HOHR);
      const isAssignedAnyStage = isAssignedHod || isAssignedCs || isAssignedHr;

      if (activeStage !== 'NONE') {
        const activeApprover = getAssignedApproverForStage(r, activeStage);
        if (isSameApprover(hodProfile, activeApprover)) {
          if (activeStage === 'HOD') pendingHodCount++;
          else if (activeStage === 'CSuite') pendingCsuiteCount++;
          else if (activeStage === 'HOHR') pendingHohrCount++;

          pendingRequests.push(item);
          allSupervisedRequests.push(item);
        } else if (isAssignedAnyStage) {
          // If logged-in user was an assigned approver for an earlier completed stage
          if (activeStage === 'CSuite' && isAssignedHod) {
            approvedCount++;
            historyRequests.push(item);
            allSupervisedRequests.push(item);
          } else if (activeStage === 'HOHR' && (isAssignedHod || isAssignedCs)) {
            approvedCount++;
            historyRequests.push(item);
            allSupervisedRequests.push(item);
          }
        }
      } else {
        // Request is completed/rejected/returned/closed
        if (isAssignedAnyStage) {
          const stLower = statusStr.toLowerCase();
          if (stLower.includes('approved') || stLower === 'upcoming') {
            approvedCount++;
          } else {
            rejectedCount++;
          }
          historyRequests.push(item);
          allSupervisedRequests.push(item);
        }
      }
    });

    const totalPending = pendingHodCount + pendingCsuiteCount + pendingHohrCount;
    const empPortalUrl = getEmployeePortalUrl();

    return ok({
      approver: hodProfile,
      metrics: {
        pending: totalPending,
        pendingHod: pendingHodCount,
        pendingCsuite: pendingCsuiteCount,
        pendingHohr: pendingHohrCount,
        approved: approvedCount,
        rejected: rejectedCount,
        total: totalPending + historyRequests.length
      },
      pendingRequests: pendingRequests,
      historyRequests: historyRequests,
      allSupervisedRequests: allSupervisedRequests,
      employeePortalUrl: empPortalUrl
    });
  } catch (e) {
    Logger.log('getApproverDashboardData error: ' + e.message);
    return err('Failed to load dashboard data: ' + e.message);
  }
}

/**
 * Server scriptlet helper: Pre-fetches initial dashboard data for instant HTML rendering.
 */
function getInitialDashboardDataJson(userEmailParam) {
  try {
    return getApproverDashboardData(userEmailParam || '');
  } catch (e) {
    return JSON.stringify({ success: false, message: e.message });
  }
}

/**
 * API: Fetch complete Training Requisition details and Supervision Queue for HOD Review
 */
/**
 * API: Fetch complete Training Requisition details.
 * SECURITY ENFORCEMENT: Server-side identity resolution & active stage access control.
 */
function getRequisitionDetails(trainingId) {
  try {
    const auth = resolveAuthenticatedUserServerSide();
    if (!auth.valid || !auth.hod) return err(auth.message || 'Unauthorized user identity.');

    const hodProfile = auth.hod;
    const cleanId = String(trainingId || '').trim().toLowerCase();
    if (!cleanId) return err('Training ID is required.');

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const trainings = sheetToJson(tSheet);
    const training = trainings.find(t => {
      const tid = getTrainingId(t).toLowerCase();
      const tcode = getTrainingCode(t).toLowerCase();
      return tid === cleanId || tcode === cleanId;
    });

    if (!training) {
      return err(`Training request (${cleanId}) not found.`);
    }

    const activeStage = getCurrentActiveApprovalStage(training);
    const assignedApprovers = getAssignedApproversForRequisition(training);
    const activeApprover = getAssignedApproverForStage(training, activeStage);

    const isCurrentActiveApprover = activeStage !== 'NONE' && isSameApprover(hodProfile, activeApprover);
    const isAssignedHod = isSameApprover(hodProfile, assignedApprovers.HOD);
    const isAssignedCs = isSameApprover(hodProfile, assignedApprovers.CSuite);
    const isAssignedHr = isSameApprover(hodProfile, assignedApprovers.HOHR);
    const isAssignedAnyStage = isAssignedHod || isAssignedCs || isAssignedHr;

    // Security Check: User must be the CURRENT active stage approver OR an assigned approver for a completed history stage.
    // Future stage approvers or unrelated users are strictly rejected!
    let isAuthorizedView = false;
    if (isCurrentActiveApprover) {
      isAuthorizedView = true;
    } else if (activeStage === 'NONE' && isAssignedAnyStage) {
      isAuthorizedView = true;
    } else if (activeStage === 'CSuite' && isAssignedHod) {
      isAuthorizedView = true; // HOD viewing their already approved request currently at C-Suite
    } else if (activeStage === 'HOHR' && (isAssignedHod || isAssignedCs)) {
      isAuthorizedView = true; // HOD or C-Suite viewing their already approved request currently at HOHR
    }

    if (!isAuthorizedView) {
      return err(`Unauthorized Access: You are not authorized to view training request (${cleanId}).`);
    }

    // Resolve participants via multi-tier fallback (Training Data, Requisition Form, DB, JSON, enriched with Employees)
    const participants = getParticipantsForRequisition(training, cleanId);

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
      activeStage: activeStage,
      isActionableForUser: isCurrentActiveApprover,
      participants: participants,
      requester: requester,
      hod: hodProfile
    });
  } catch (e) {
    return err('Failed to load requisition details: ' + e.message);
  }
}

/**
 * Helper: Retrieve all requisitions strictly under a specific HOD's supervision.
 */
/**
 * Helper: Retrieve all requisitions where the authenticated approver is an assigned approver for at least one stage.
 */
function getHODSupervisedRequisitions(hodProfile) {
  try {
    const tSheet = getSheet('Trainings');
    if (!tSheet) return [];
    const trainings = sheetToJson(tSheet);
    if (!trainings || trainings.length === 0) return [];

    return trainings.filter(t => {
      const assigned = getAssignedApproversForRequisition(t);
      return isSameApprover(hodProfile, assigned.HOD) ||
             isSameApprover(hodProfile, assigned.CSuite) ||
             isSameApprover(hodProfile, assigned.HOHR);
    });
  } catch(e) {
    Logger.log('getHODSupervisedRequisitions error: ' + e.message);
    return [];
  }
}

/**
 * API: Submit HOD / CSuite / HOHR Decision
 * SECURITY ENFORCEMENT: Server-side identity resolution & Active Stage Authorization.
 */
function submitHODDecision(data) {
  try {
    if (!data || !data.trainingId || !data.decision) {
      return err('Training ID and decision are required.');
    }

    const cleanId = String(data.trainingId).trim();
    // STRICT SECURITY: Resolve identity strictly from server-side session
    const auth = resolveAuthenticatedUserServerSide();
    if (!auth.valid || !auth.hod) {
      return err(auth.message || 'Unauthorized user identity.');
    }

    const hodProfile = auth.hod;
    const hodName = hodProfile.Name || 'Approver';
    const hodId = hodProfile.ID || 'N/A';
    const hodCostCentre = hodProfile.CostCentre || 'N/A';

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const row = findRowById(tSheet, cleanId);
    if (row === -1) return err(`Training request (${cleanId}) not found in database.`);

    const trainingsList = sheetToJson(tSheet);
    const currentT = trainingsList.find(t => {
      const id = String(t.ID || '').trim().toLowerCase();
      const code = String(t.Code || '').trim().toLowerCase();
      const tId = String(t.TrainingID || '').trim().toLowerCase();
      const target = cleanId.toLowerCase();
      return id === target || code === target || tId === target;
    });
    if (!currentT) return err(`Training request (${cleanId}) not found.`);

    const currentAppStatus = String(currentT.ApprovalStatus || currentT.Status || 'Pending HOD Approval');

    // 1. Determine Current Active Approval Stage
    const activeStage = getCurrentActiveApprovalStage(currentT);
    if (activeStage === 'NONE') {
      return err(`Invalid Action: Training request (${cleanId}) is not currently awaiting approval.`);
    }

    const approvalStep = activeStage === 'CSuite' ? 'Csuite' : activeStage;

    // 2. Resolve Assigned Approver for CURRENT Active Stage
    const assignedApprover = getAssignedApproverForStage(currentT, activeStage);
    if (!isSameApprover(hodProfile, assignedApprover)) {
      return err(`Unauthorized Access: You are not the assigned approver for the current active approval stage (${activeStage}) of training request (${cleanId}).`);
    }

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

    const remarks = data.remarks || '';
    let nextApprovalStatus = validDecision;

    // Retrieve assigned approver profiles for auto-bypass checks & email notifications
    const assignedAll = getAssignedApproversForRequisition(currentT);
    const hodIsCs = isSameApprover(assignedAll.HOD, assignedAll.CSuite);
    const csIsHr = isSameApprover(assignedAll.CSuite, assignedAll.HOHR);
    const hodIsHr = isSameApprover(assignedAll.HOD, assignedAll.HOHR);

    const hodIsCsuite = hodIsCs;
    const csuiteIsHohr = csIsHr;
    const hodIsHohr = hodIsHr;

    const userIsCsuite = Boolean(hodProfile.isCsuite);
    const userIsHohr = Boolean(hodProfile.isHohr);

    const csProfile = assignedAll.CSuite;
    const csuiteName = csProfile ? csProfile.Name : 'C-Suite Executive';
    const csuiteId = csProfile ? csProfile.ID : '';

    const hrProfile = assignedAll.HOHR;
    const hohrName = hrProfile ? hrProfile.Name : 'Head of HR';
    const hohrId = hrProfile ? hrProfile.ID : '';

    let autoVerifiedRemarks = '';

    if (validDecision === 'Approved') {
      if (activeStage === 'HOD') {
        updateCol('HOD', hodProfile.Name);
        updateCol('HODStatus', 'Approved');

        if (hodIsCs && csIsHr) {
          nextApprovalStatus = 'Approved';
          autoVerifiedRemarks = ' (Auto-verified: HOD, C-Suite, and HOHR are the same approver)';
          updateCol('Csuite', hodProfile.Name);
          updateCol('CsuiteStatus', 'Approved');
          updateCol('HOHR', hodProfile.Name);
          updateCol('HOHRStatus', 'Approved');
        } else if (hodIsCs) {
          nextApprovalStatus = 'Pending HOHR Approval';
          autoVerifiedRemarks = ' (Auto-verified: HOD and C-Suite are the same approver)';
          updateCol('Csuite', hodProfile.Name);
          updateCol('CsuiteStatus', 'Approved');
          updateCol('HOHRStatus', 'Pending');
        } else {
          nextApprovalStatus = 'Pending C-Suite Approval';
          updateCol('CsuiteStatus', 'Pending');
        }
      } else if (activeStage === 'CSuite') {
        updateCol('Csuite', hodProfile.Name);
        updateCol('CsuiteStatus', 'Approved');

        if (csIsHr || hodIsHr) {
          nextApprovalStatus = 'Approved';
          autoVerifiedRemarks = ' (Auto-verified: C-Suite and HOHR are the same approver)';
          updateCol('HOHR', hodProfile.Name);
          updateCol('HOHRStatus', 'Approved');
        } else {
          nextApprovalStatus = 'Pending HOHR Approval';
          updateCol('HOHRStatus', 'Pending');
        }
      } else if (activeStage === 'HOHR') {
        updateCol('HOHR', hodProfile.Name);
        updateCol('HOHRStatus', 'Approved');
        nextApprovalStatus = 'Approved';
      }
    } else {
      // Decision is Rejected or Returned
      if (activeStage === 'HOD') {
        updateCol('HODStatus', validDecision);
      } else if (activeStage === 'CSuite') {
        updateCol('CsuiteStatus', validDecision);
      } else if (activeStage === 'HOHR') {
        updateCol('HOHRStatus', validDecision);
      }
      nextApprovalStatus = validDecision;
    }

    const fullRemarks = (remarks || '') + autoVerifiedRemarks;
    updateCol('ApprovalStatus', nextApprovalStatus);
    updateCol('ApprovedBy', `${hodProfile.Name} (${hodProfile.ID})`);
    updateCol('ApprovedCostCentre', hodProfile.CostCentre);
    updateCol('ApprovedAt', timestamp);
    updateCol('ApprovalRemarks', fullRemarks.trim());

    if (nextApprovalStatus === 'Approved') {
      updateCol('Status', 'Upcoming');
      updateCol('Stage', Number(currentT.Participants || 0) > 0 ? 'Participants Imported' : 'Created');
    } else if (validDecision === 'Rejected') {
      updateCol('Status', 'Cancelled');
      updateCol('Stage', 'Programme Closed');
    } else if (validDecision === 'Returned') {
      updateCol('Status', 'Pending Revision');
      updateCol('Stage', 'Form Returned');
    } else {
      // Intermediate approval tier (Pending C-Suite / Pending HOHR) - clear any previous returned status
      updateCol('Status', 'Draft');
      updateCol('Stage', 'Created');
    }

    updateCol('UpdatedDate', timestamp);
    SpreadsheetApp.flush();

    // Automatically update AP-HRD-F01-01 Training Requisition Form Google Sheet digital approval signatures
    let hodPosition = hodProfile.Position || hodProfile.RoleTitle || '';
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
      let requesterCostCentre = currentT.CostCentre || currentT.Department || currentT['Cost Centre'] || currentT['CostCentre'] || '';

      if ((!requesterEmail || !requesterCostCentre || !requesterName) && getSheet('Employees')) {
        const emps = sheetToJson(getSheet('Employees'));
        const cleanReqId = String(requesterId).toLowerCase().trim();
        const cleanReqName = String(requesterName).toLowerCase().trim();

        const empMatch = emps.find(e => {
          const eId = String(e['Employee No'] || e.EmployeeNo || e.ID || e.EmployeeID || e['Employee ID'] || '').toLowerCase().trim();
          const eName = String(e.Name || e.EmployeeName || e['Employee Name'] || '').toLowerCase().trim();
          return (cleanReqId && eId === cleanReqId) || (cleanReqName && eName === cleanReqName);
        });

        if (empMatch) {
          if (!requesterEmail) {
            requesterEmail = String(empMatch.Email || empMatch.EmailAddress || empMatch['Email Address'] || '').trim();
          }
          if (!requesterName || requesterName === requesterId) {
            requesterName = String(empMatch.Name || empMatch.EmployeeName || empMatch['Employee Name'] || requesterName).trim();
          }
          if (!requesterCostCentre) {
            requesterCostCentre = String(empMatch.Department || empMatch.CostCentre || empMatch['Cost Centre'] || '').trim();
          }
        }
      }

      if (!requesterCostCentre) {
        requesterCostCentre = hodCostCentre || 'N/A';
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

    const logSummary = (draftLog && draftLog.length > 0) ? ` (${draftLog.join('; ')})` : '';

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
 * API: Fetch participants pending 3-Month Post Evaluation for a supervisor / HOD
 * Grouped and separated by training programme.
 */
function getPendingPostEvalParticipants(trainingId, userEmail) {
  try {
    let auth = validateHODAccess(userEmail);
    if (!auth.valid || !auth.hod) {
      return err(auth.message || `Access Denied: The account (${userEmail || 'anonymous'}) is not registered as an authorized approver.`);
    }
    let targetEmail = auth.email || userEmail || '';
    let targetEmpId = auth.hod.ID || auth.hod.EmployeeID || '';
    let targetHodCostCentre = auth.hod.CostCentre || 'ALL';

    const tSheet = getSheet('Trainings');
    if (!tSheet) return err('Trainings sheet unavailable.');

    const trainings = sheetToJson(tSheet);

    const candidateTrainings = trainings.filter(t => {
      const st = String(t.Stage || t.Status || t.ApprovalStatus || '').toLowerCase();
      return st.includes('3-month') || st.includes('6-month') || st.includes('completed') || st.includes('approved') || st.includes('ongoing');
    });

    const activeList = candidateTrainings.length > 0 ? candidateTrainings : trainings;
    const groupedTrainings = [];
    let overallPendingCount = 0;

    activeList.forEach(t => {
      const targetTId = getTrainingId(t);
      const allParticipants = getParticipantsForRequisition(t, targetTId);
      if (!allParticipants || allParticipants.length === 0) return;
      const ss = getTrainingDataSpreadsheet(targetTId);

      // Filter participants assigned to this supervisor/HOD (by SupervisorID, SupervisorEmail, or Cost Centre)
      let supParticipants = allParticipants.filter(p => {
        const supEmail = String(p.SupervisorEmail || '').trim().toLowerCase();
        const supId    = String(p.SupervisorID || '').trim().toLowerCase();
        const uEmail   = String(targetEmail).trim().toLowerCase();
        const uEmpId   = String(targetEmpId).trim().toLowerCase();

        if (supEmail && uEmail && supEmail === uEmail) return true;
        if (supId && uEmpId && isSameEmployeeId(supId, uEmpId)) return true;
        if (supEmail && uEmpId && supEmail.toLowerCase().includes(uEmpId)) return true;

        // Fallback: match by HOD Cost Centre if no direct supervisor set
        const dept = String(p.Department || p.CostCentre || '').toLowerCase();
        const hodCc = String(targetHodCostCentre).toLowerCase();
        return (!p.SupervisorEmail && !p.SupervisorID) && (dept.includes(hodCc) || hodCc.includes(dept) || targetHodCostCentre === 'ALL');
      });

      // Check completed Post Evaluation records
      const postSheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
      let evaluatedEmpIds = [];
      if (postSheet) {
        const postRows = sheetToJson(postSheet);
        evaluatedEmpIds = postRows.map(pr => String(pr.EmployeeID || '').trim().toLowerCase());
      }

      const pendingList = supParticipants.filter(p => !evaluatedEmpIds.some(eid => isSameEmployeeId(eid, p.EmployeeID || p.ID || '')));
      const completedList = supParticipants.filter(p => evaluatedEmpIds.some(eid => isSameEmployeeId(eid, p.EmployeeID || p.ID || '')));

      if (pendingList.length > 0 || completedList.length > 0) {
        groupedTrainings.push({
          training: t,
          totalParticipants: supParticipants.length,
          pendingCount: pendingList.length,
          completedCount: completedList.length,
          pendingParticipants: pendingList,
          completedParticipants: completedList
        });
        overallPendingCount += pendingList.length;
      }
    });

    let primaryGroup = groupedTrainings.find(g => getTrainingId(g.training) === String(trainingId || '').trim()) || groupedTrainings[0];

    return ok({
      training: primaryGroup ? primaryGroup.training : (activeList[0] || null),
      hod: auth.hod || { Email: userEmail, Name: userEmail },
      groupedTrainings: groupedTrainings,
      totalPendingOverall: overallPendingCount,
      pendingCount: primaryGroup ? primaryGroup.pendingCount : 0,
      completedCount: primaryGroup ? primaryGroup.completedCount : 0,
      pendingParticipants: primaryGroup ? primaryGroup.pendingParticipants : [],
      completedParticipants: primaryGroup ? primaryGroup.completedParticipants : [],
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

    let postSheet = ss.getSheetByName('Post Evaluation') || ss.getSheetByName('PostEval');
    if (!postSheet) {
      postSheet = ss.insertSheet('Post Evaluation');
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

