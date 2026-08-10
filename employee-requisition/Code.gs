/**
 * Code.gs — Employee Training Requisition Web App Router & Server Services
 */

function doGet(e) {
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — Employee Training Requisition');

  try {
    const template = HtmlService.createTemplateFromFile('Requisition');
    template.params = (e && e.parameter) ? e.parameter : {};

    return template.evaluate()
      .setTitle(appTitle)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log('doGet error: ' + err.message);
    try {
      const errTemplate = HtmlService.createTemplateFromFile('Error');
      errTemplate.params = { message: 'Failed to load page: ' + err.message };
      return errTemplate.evaluate()
        .setTitle('Error — TrainHub')
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
 * API: Get current logged-in Google account email
 */
function getLoggedInUserEmail() {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const effectiveEmail = Session.getEffectiveUser().getEmail();
    return ok(activeEmail || effectiveEmail || '');
  } catch (e) {
    return ok('');
  }
}

/**
 * API: Get real Cost Centres from EMPLOYEE_SPREADSHEET_ID
 */
function getCostCentres() {
  try {
    const ss = getEmployeeSpreadsheet();
    if (ss) {
      const ccSheet = ss.getSheetByName('Cost Centre') || ss.getSheetByName('CostCentre') || ss.getSheetByName('Cost Centres');
      if (ccSheet && ccSheet.getLastRow() > 1) {
        const data = ccSheet.getDataRange().getValues();
        const list = [];
        for (let i = 1; i < data.length; i++) {
          const val = String(data[i][0] || data[i][1] || '').trim();
          if (val && !list.includes(val)) list.push(val);
        }
        if (list.length > 0) return ok(list);
      }

      const empSheet = getSheet('Employees');
      if (empSheet) {
        const rows = sheetToJson(empSheet);
        const uniqueCC = new Set();
        rows.forEach(r => {
          const dept = String(r.Department || r.CostCentre || '').trim();
          if (dept) uniqueCC.add(dept);
        });
        if (uniqueCC.size > 0) return ok(Array.from(uniqueCC));
      }
    }
  } catch (e) {
    Logger.log('getCostCentres error: ' + e.message);
  }
  return ok(['Cost Centre 101 - Engineering', 'Cost Centre 102 - HR', 'Cost Centre 103 - Finance', 'Cost Centre 104 - Operations', 'Cost Centre 105 - IT', 'Cost Centre 106 - Sales']);
}

/**
 * API: Search and retrieve employee participants list strictly from EMPLOYEE_SPREADSHEET_ID
 */
function getParticipantsList() {
  try {
    const directory = getOfficialEmployeeDirectory();
    const emps = Object.keys(directory.byId).map(key => directory.byId[key]);
    return ok(emps);
  } catch (e) {
    return err('Failed to load participants: ' + e.message);
  }
}

/**
 * API: Lookup Employee Details by Employee ID for Auto-Filling Requisition Form
 */
function getEmployeeDetails(employeeId) {
  try {
    const empCheck = getValidEmployee(employeeId);
    if (!empCheck.valid) return err(empCheck.message);

    const emp = empCheck.employee;
    const requestDate = getFormattedCurrentDate();

    return ok({
      ID: emp.ID || employeeId,
      Name: emp.Name || '',
      Department: emp.Department || emp.CostCentre || 'N/A',
      Position: emp.Position || emp.JobTitle || emp.PositionTitle || 'Staff',
      Email: emp.Email || '',
      RequestDate: requestDate
    });
  } catch (e) {
    return err('Employee lookup error: ' + e.message);
  }
}

/**
 * API: Submit Employee Training Requisition Request Form
 */
function submitEmployeeRequisition(data) {
  try {
    if (!data || !data.EmployeeID || !data.TrainingName || !data.StartDate) {
      return err('Employee ID, Training Name, and Start Date are required.');
    }

    // 1. Employee lookup
    const empCheck = getValidEmployee(data.EmployeeID);
    if (!empCheck.valid) return err(empCheck.message);

    const emp = empCheck.employee;
    const empEmail = data.Email || emp.Email || '';

    // 2. Company email validation check
    const emailCheck = validateCompanyEmail(empEmail);
    if (!emailCheck.valid) return err(emailCheck.message);

    const sheet = getSheet('Trainings');
    if (!sheet) return err('Trainings sheet unavailable in master database.');

    const id = generateId('TRN');
    const code = generateTrainingCode(data.Category || 'General');
    const timeNow = now();

    // 3. Multi-Tier Approval & Auto-Bypass Workflow Logic (2-Step Lookup: "For IT" -> "HOD email")
    let assignedHodName = '';
    let hodName = '', hodEmail = '';
    let csuiteName = '', csuiteEmail = '';
    let hohrName = '', hohrEmail = '';

    const requesterId = String(emp.ID || data.EmployeeID || '').toLowerCase().trim();
    const requesterName = String(emp.Name || '').toLowerCase().trim();
    const empDept = String(emp.Department || emp.CostCentre || data.Department || '').toLowerCase().trim();

    // Step 1: Find assigned HOD in "For IT" tab
    try {
      const itSheet = getSheet('For IT');
      if (itSheet) {
        const itRows = sheetToJson(itSheet);
        const matchedIt = itRows.find(r => {
          const rId = String(r.ID || r.EmployeeID || r.EmployeeNo || r['Employee ID'] || r['Employee No'] || '').toLowerCase().trim();
          const rName = String(r.Name || r.EmployeeName || r['Employee Name'] || '').toLowerCase().trim();
          if ((rId && rId === requesterId) || (rName && rName === requesterName) || (requesterName && rName.includes(requesterName))) return true;
          return false;
        });

        if (matchedIt) {
          assignedHodName = String(matchedIt.HOD || matchedIt.HODName || matchedIt.HodName || matchedIt.Manager || matchedIt.ReportTo || matchedIt['HOD Name'] || matchedIt['HOD'] || '').trim();
          Logger.log(`Found assigned HOD "${assignedHodName}" for employee "${emp.Name}" in For IT tab.`);
        }
      }
    } catch (itErr) {
      Logger.log('For IT lookup error: ' + itErr.message);
    }

    // Step 2: Compare assigned HOD in "HOD email" tab (Header: Employee No | HOD | Cost Centre | Position Title | Email)
    try {
      const hodSheet = getSheet('HOD email');
      if (hodSheet) {
        const hodRows = sheetToJson(hodSheet);
        let matchedHod = null;

        const getVal = (rowObj, nameList) => {
          if (!rowObj) return '';
          const keys = Object.keys(rowObj);
          for (let n of nameList) {
            const matchKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === n.toLowerCase().replace(/[^a-z0-9]/g, ''));
            if (matchKey && rowObj[matchKey] !== undefined && rowObj[matchKey] !== null) return String(rowObj[matchKey]).trim();
          }
          return '';
        };

        if (assignedHodName) {
          const cleanAssigned = assignedHodName.toLowerCase();
          matchedHod = hodRows.find(h => {
            const hName = getVal(h, ['HOD', 'HODName', 'HodName', 'Name']).toLowerCase();
            return hName && (hName.includes(cleanAssigned) || cleanAssigned.includes(hName));
          });
        }

        if (!matchedHod) {
          matchedHod = hodRows.find(h => {
            const hEmpId = getVal(h, ['Employee No', 'EmployeeNo', 'EmployeeID', 'ID']).toLowerCase();
            const hEmpName = getVal(h, ['HOD', 'HODName', 'HodName', 'Name']).toLowerCase();
            return (hEmpId && hEmpId === requesterId) || (hEmpName && hEmpName === requesterName);
          });
        }

        if (!matchedHod && empDept) {
          const deptCodeMatch = empDept.match(/\d+/);
          const deptCode = deptCodeMatch ? deptCodeMatch[0] : '';
          matchedHod = hodRows.find(h => {
            const hDept = getVal(h, ['Cost Centre', 'CostCentre', 'Department']).toLowerCase();
            const hCodeMatch = hDept.match(/\d+/);
            if (deptCode && hCodeMatch && deptCode === hCodeMatch[0]) return true;
            return hDept && (empDept.includes(hDept) || hDept.includes(empDept));
          });
        }

        if (!matchedHod && hodRows.length > 0) {
          matchedHod = hodRows[0];
        }

        if (matchedHod) {
          hodName = getVal(matchedHod, ['HOD', 'HODName', 'HodName', 'Name']) || assignedHodName || '';
          hodEmail = getVal(matchedHod, ['Email', 'EmailAddress', 'HODEmail']) || '';
          csuiteName = getVal(matchedHod, ['CsuiteName', 'CSuiteName', 'Csuite']) || '';
          csuiteEmail = getVal(matchedHod, ['CsuiteEmail', 'CSuiteEmail']) || '';
          hohrName = getVal(matchedHod, ['HohrName', 'HOHRName', 'HOHR']) || '';
          hohrEmail = getVal(matchedHod, ['HohrEmail', 'HOHREmail']) || '';
        }
      }
    } catch(e) {
      Logger.log('HOD email tab lookup error: ' + e.message);
    }

    if (!hodEmail) hodEmail = data.HodEmail || getConfigProperty('ADMIN_EMAILS', '');

    const hName = String(hodName || '').toLowerCase().trim();
    const cName = String(csuiteName || '').toLowerCase().trim();
    const hrName = String(hohrName || '').toLowerCase().trim();

    // Auto-bypass rules
    const isHodBypassed = requesterName && hName && (requesterName === hName);
    const isCsuiteBypassed = isHodBypassed && hName && cName && (hName === cName);
    const isHohrBypassed = isCsuiteBypassed && cName && hrName && (cName === hrName);

    let currentApprovalStatus = 'Pending HOD Approval';
    let approvedByVal = '';
    let approvedCostCentreVal = '';
    let approvedAtVal = '';
    let approvalRemarksVal = '';

    if (isHodBypassed) {
      approvedByVal = `Auto-Verified: ${emp.Name} (${emp.ID})`;
      approvedCostCentreVal = emp.Department || 'N/A';
      approvedAtVal = timeNow;
      approvalRemarksVal = 'Auto-verified: Requester is HOD.';

      if (isCsuiteBypassed && isHohrBypassed) {
        currentApprovalStatus = 'Approved';
      } else if (isCsuiteBypassed) {
        currentApprovalStatus = 'Pending HOHR Approval';
      } else {
        currentApprovalStatus = 'Pending C-Suite Approval';
      }
    }

    // Handle Brochure File upload if provided
    let brochureUrl = data.BrochureUrl || '';
    if (data.BrochureFile && data.BrochureFile.data) {
      try {
        const fileBlob = Utilities.newBlob(Utilities.base64Decode(data.BrochureFile.data), data.BrochureFile.mimeType || 'application/octet-stream', data.BrochureFile.name || 'brochure');
        const driveFile = DriveApp.createFile(fileBlob);
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        brochureUrl = driveFile.getUrl();
      } catch(fErr) {
        Logger.log('Brochure file upload error: ' + fErr.message);
      }
    }

    const objectivesVal = (data.Objectives || data.Reason || '') + (brochureUrl ? `\n[Brochure File: ${brochureUrl}]` : '');

    // Resolve every selection against the Employee sheet before it is stored,
    // written to Training Data, or placed on the requisition form.
    const requestedParticipants = Array.isArray(data.ParticipantList) ? data.ParticipantList : (Array.isArray(data.participants) ? data.participants : []);
    const participantResolution = canonicalizeTrainingParticipants(requestedParticipants);
    if (participantResolution.rejected.length > 0) {
      return err('The following participant(s) do not match an Employee-sheet record: ' + participantResolution.rejected.join(', '));
    }
    const participantsList = participantResolution.participants;

    // Create Google Drive Workspace and Training Requisition Form (AP-HRD-F01-01)
    let workspace = { folderId: '', folderUrl: '', partSheetId: '', sessionSheetId: '', attendanceSheetId: '', evaluationSheetId: '', postSheetId: '' };
    let reqForm = { fileId: '', fileUrl: '' };
    const requesterSigData = {
      employeeNo: emp.ID || data.EmployeeID || data.EmployeeNo || '',
      name: emp.Name || data.EmployeeName || data.Name || 'Requester',
      position: emp.Position || emp.JobTitle || emp.PositionTitle || data.JobPosition || data.Position || data.JobTitle || 'Requester',
      date: getFormattedCurrentDate()
    };

    try {
      workspace = createTrainingWorkspace(code, data.TrainingName);
      reqForm = createTrainingRequisitionForm(code, {
        Name: data.TrainingName,
        CourseFee: data.CourseFee,
        StartDate: data.StartDate,
        EndDate: data.EndDate || data.StartDate,
        Duration: data.Duration || 1,
        TotalHours: data.TotalHours || 8,
        Venue: data.Venue,
        Trainer: data.Trainer,
        TrainingProvider: data.TrainingProvider,
        Objectives: objectivesVal,
        ParticipantList: participantsList
      }, workspace.folderId, requesterSigData);
    } catch(wErr) {
      Logger.log('Workspace / Form creation error: ' + wErr.message);
    }

    const rowData = [
      id,                                // Col 1 (A): ID
      code,                              // Col 2 (B): Code
      data.TrainingName,                 // Col 3 (C): Name
      data.Category || 'General',        // Col 4 (D): Category
      data.Trainer || 'TBD',             // Col 5 (E): Trainer
      data.Venue || 'TBD',               // Col 6 (F): Venue
      data.StartDate,                    // Col 7 (G): StartDate
      data.EndDate || data.StartDate,    // Col 8 (H): EndDate
      data.Duration || 1,                // Col 9 (I): Duration
      data.TotalHours || 8,              // Col 10 (J): TotalHours
      emp.Department || data.Department || 'N/A', // Col 11 (K): Department
      objectivesVal,                     // Col 12 (L): Objectives
      'Draft',                           // Col 13 (M): Status
      'Created',                         // Col 14 (N): Stage
      data.TotalPax || participantsList.length || 0, // Col 15 (O): Participants
      workspace.folderId || '',          // Col 16 (P): FolderID
      workspace.partSheetId || '',       // Col 17 (Q): ParticipantsSheetID
      workspace.sessionSheetId || '',    // Col 18 (R): SessionsSheetID
      workspace.attendanceSheetId || '', // Col 19 (S): AttendanceSheetID
      workspace.evaluationSheetId || '', // Col 20 (T): EvaluationSheetID
      workspace.postSheetId || '',       // Col 21 (U): PostSheetID
      reqForm.fileId || '',              // Col 22 (V): RequisitionFormFileID
      timeNow,                           // Col 23 (W): CreatedDate
      timeNow,                           // Col 24 (X): UpdatedDate
      data.CourseFee || '0.00'           // Col 25 (Y): CourseFee
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    ensureTrainingSheetColumns(sheet);

    const lastRow = sheet.getLastRow();
    const headers = sheet.getDataRange().getValues()[0].map(h => String(h).trim());
    const setCol = (name, val) => {
      const idx = headers.indexOf(name) + 1;
      if (idx > 0) sheet.getRange(lastRow, idx).setValue(val);
    };

    setCol('ApprovalStatus', currentApprovalStatus);
    setCol('ApprovedBy', approvedByVal);
    setCol('ApprovedCostCentre', approvedCostCentreVal);
    setCol('ApprovedAt', approvedAtVal);
    setCol('ApprovalRemarks', approvalRemarksVal);
    setCol('RequestedBy', emp.ID || data.EmployeeID);
    setCol('RequestedByName', emp.Name || data.EmployeeName || '');
    setCol('RequestedByEmail', empEmail);
    setCol('RequestedDate', timeNow);
    setCol('FolderID', workspace.folderId || '');
    setCol('ParticipantsSheetID', workspace.partSheetId || '');
    setCol('SessionsSheetID', workspace.sessionSheetId || '');
    setCol('AttendanceSheetID', workspace.attendanceSheetId || '');
    setCol('EvaluationSheetID', workspace.evaluationSheetId || '');
    setCol('PostSheetID', workspace.postSheetId || '');
    setCol('RequisitionFormFileID', reqForm.fileId || '');

    SpreadsheetApp.flush();

    if (participantsList.length > 0) {
      try {
        const tpMasterSheet = getSheet('TrainingParticipants');
        if (tpMasterSheet) {
          const tpRows = participantsList.map(p => [
            generateId('TP'),
            id,
            p.ID || p.EmployeeID || p.EmployeeNo || '',
            p.Name || p.EmployeeName || '',
            p.Department || p.CostCentre || emp.Department || '',
            p.Position || p.JobTitle || '',
            timeNow
          ]);
          tpMasterSheet.getRange(tpMasterSheet.getLastRow() + 1, 1, tpRows.length, 7).setValues(tpRows);
        }
      } catch(pErr) {
        Logger.log('Error saving participants in submitEmployeeRequisition: ' + pErr.message);
      }
      try {
        syncParticipantsToTrainingDriveSheet(id, participantsList);
      } catch(sErr) {
        Logger.log('Error syncing participants to training drive sheet: ' + sErr.message);
      }
    }

    if (reqForm && reqForm.fileId) {
      try {
        updateTrainingRequisitionSignatures(id, 'request', requesterSigData, reqForm.fileId);
        syncTrainingRequisitionParticipants(id);
      } catch(syncErr) {
        Logger.log('sync/sig error in submitEmployeeRequisition: ' + syncErr.message);
      }

      if (isHodBypassed) {
        try {
          updateTrainingRequisitionSignatures(id, 'HOD', {
            status: currentApprovalStatus,
            employeeNo: emp.ID || data.EmployeeID,
            name: emp.Name || 'HOD',
            position: emp.Position || 'HOD',
            date: timeNow
          }, reqForm.fileId);
        } catch(sigErr) {}
      }

      if (isCsuiteBypassed) {
        try {
          updateTrainingRequisitionSignatures(id, 'Csuite', {
            status: currentApprovalStatus,
            employeeNo: emp.ID || data.EmployeeID,
            name: emp.Name || 'C-Suite',
            position: emp.Position || 'C-Suite',
            date: timeNow
          }, reqForm.fileId);
        } catch(sigErr) {}
      }

      if (isHohrBypassed) {
        try {
          updateTrainingRequisitionSignatures(id, 'HOHR', {
            status: currentApprovalStatus,
            employeeNo: emp.ID || data.EmployeeID,
            name: emp.Name || 'Head of HR',
            position: emp.Position || 'Head of HR',
            date: timeNow
          }, reqForm.fileId);
        } catch(sigErr) {}
      }
    }

    // Create Gmail Draft ONLY (Send function removed as requested)
    let draftStatus = 'Not created';
    try {
      const hodPortalUrl = getConfigProperty('HOD_PORTAL_URL', '');
      const reviewUrl = hodPortalUrl ? `${hodPortalUrl}?page=review&id=${id}` : getAppUrl();

      let recipientEmail = hodEmail;
      if (currentApprovalStatus === 'Pending C-Suite Approval' && csuiteEmail) recipientEmail = csuiteEmail;
      else if (currentApprovalStatus === 'Pending HOHR Approval' && hohrEmail) recipientEmail = hohrEmail;
      else if (currentApprovalStatus === 'Approved') recipientEmail = 'arina.ismail@apollofood.com.my';

      const subject = `[TrainHub DRAFT] Training Requisition ${currentApprovalStatus} - ${data.TrainingName}`;
      const body = `Dear Approver / Manager,\n\nA new Training Requisition Form (AP-HRD-F01-01) has been submitted:\n\n` +
        `Requester: ${emp.Name || data.EmployeeID} (${emp.Department || 'N/A'})\n` +
        `Employee ID: ${emp.ID || data.EmployeeID}\n` +
        `Assigned HOD: ${hodName || 'N/A'} (${hodEmail || recipientEmail || 'N/A'})\n` +
        `Training Name: ${data.TrainingName}\n` +
        `Category: ${data.Category || 'General'}\n` +
        `Proposed Date: ${data.StartDate} to ${data.EndDate || data.StartDate}\n` +
        `Duration: ${data.Duration || 1} days (${data.TotalHours || 8} hrs)\n` +
        `Estimated Fee: RM ${data.CourseFee || '0.00'}\n` +
        `Current Status: ${currentApprovalStatus}\n` +
        (brochureUrl ? `Brochure Attachment/Link: ${brochureUrl}\n` : '') +
        `\nPlease review the request details:\n${reviewUrl}\n\n` +
        `Thank you,\nTrainHub Training Management System`;

      if (recipientEmail) {
        try {
          const draft = GmailApp.createDraft(recipientEmail, subject, body);
          if (draft && draft.getId()) {
            draftStatus = `Draft created successfully (ID: ${draft.getId()}) for ${recipientEmail}`;
            Logger.log(draftStatus);
          } else {
            draftStatus = `Draft creation returned no ID for ${recipientEmail}`;
          }
        } catch(draftErr) {
          draftStatus = `Draft creation error: ${draftErr.message}`;
          Logger.log(draftStatus);
        }
      } else {
        draftStatus = 'No recipient HOD email available to create draft.';
      }

      if (currentApprovalStatus === 'Approved' && recipientEmail !== 'arina.ismail@apollofood.com.my') {
        try {
          GmailApp.createDraft('arina.ismail@apollofood.com.my', `[TrainHub DRAFT] Training Requisition Fully Approved - ${data.TrainingName}`, body);
        } catch(dErr) {}
      }
    } catch (mailErr) {
      draftStatus = `Notification error: ${mailErr.message}`;
      Logger.log(draftStatus);
    }

    return ok({
      message: `Training Requisition Form (AP-HRD-F01-01) submitted! Status: ${currentApprovalStatus}. HOD: ${hodName || 'Assigned HOD'}. ${draftStatus}`,
      trainingId: id,
      trainingCode: code,
      approvalStatus: currentApprovalStatus,
      assignedHod: hodName,
      draftStatus: draftStatus
    });
  } catch (e) {
    Logger.log('submitEmployeeRequisition error: ' + e.message);
    return err('Failed to submit training requisition: ' + e.message);
  }
}

/**
 * Run this test function in the Apps Script online editor to authorize Gmail Draft permissions!
 */
function testCreateDraftPermission() {
  const draft = GmailApp.createDraft(Session.getActiveUser().getEmail() || 'test@example.com', '[TrainHub TEST DRAFT]', 'This is a test draft to verify Gmail draft permissions.');
  Logger.log('Test draft created successfully! ID: ' + draft.getId());
}
