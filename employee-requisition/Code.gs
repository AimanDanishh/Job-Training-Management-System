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
/**
 * API: Retrieve all training requests submitted by a specific employee ID
 */
function getEmployeeSubmittedRequests(employeeId) {
  try {
    if (!employeeId || String(employeeId).trim() === '') {
      return err('Employee ID is required.');
    }

    const cleanId = String(employeeId).trim().toLowerCase();
    const sheet = getSheet('Trainings');
    if (!sheet) return ok([]);

    ensureTrainingSheetColumns(sheet);

    const rows = sheetToJson(sheet);

    const matchedRequests = rows.filter(r => {
      const reqId = String(r.RequestedBy || r.EmployeeID || r.EmployeeNo || '').trim().toLowerCase();
      return reqId === cleanId;
    });

    const list = matchedRequests.map(r => {
      const tId = String(r.ID || r.TrainingID || '').trim();
      const tCode = String(r.Code || '').trim();
      const cleanTId = tId.toLowerCase();
      const cleanTCode = tCode.toLowerCase();

      const trainingData = getTrainingDataSpreadsheet(tId || tCode);
      const participantSheet = trainingData ? (trainingData.getSheetByName('Participants') || trainingData.getSheetByName('TrainingParticipants')) : null;
      const rawParticipants = participantSheet ? sheetToJson(participantSheet) : [];

      let reqParticipants = rawParticipants.map(p => ({
        ID: String(p.EmployeeID || p.EmployeeNo || (p.ID && !String(p.ID).startsWith('TP-') ? p.ID : '') || '').trim(),
        EmployeeID: String(p.EmployeeID || p.EmployeeNo || (p.ID && !String(p.ID).startsWith('TP-') ? p.ID : '') || '').trim(),
        Name: String(p.EmployeeName || p.Name || '').trim(),
        EmployeeName: String(p.EmployeeName || p.Name || '').trim(),
        Department: String(p.Department || p.CostCentre || '').trim(),
        Position: String(p.Position || p.JobTitle || '').trim()
      }));

      // Level 3 Fallback: Read from AP-HRD-F01-01 Requisition Form Google Sheet if Master list is empty
      if (reqParticipants.length === 0 && r.RequisitionFormFileID) {
        try {
          const formSs = SpreadsheetApp.openById(r.RequisitionFormFileID);
          const formSheet = formSs.getSheetByName('Training Form') || formSs.getSheets()[0];
          const formRows = formSheet.getRange('A15:I38').getValues();
          formRows.forEach(fr => {
            const empId = String(fr[0] || '').trim();
            const name = String(fr[2] || '').trim();
            const dept = String(fr[3] || '').trim();
            const pos = String(fr[7] || '').trim();
            if (empId || name) {
              reqParticipants.push({
                ID: empId,
                EmployeeID: empId,
                Name: name,
                EmployeeName: name,
                Department: dept,
                Position: pos
              });
            }
          });
        } catch(eForm) {}
      }

      const dbPax = parseInt(r.Participants || r.TotalPax || r['Total Pax'] || r['TotalPax'] || r['Total Participants'] || r['TotalParticipant'] || 0, 10) || 0;
      const totalPaxVal = Math.max(dbPax, reqParticipants.length);

      return {
        ID: tId,
        Code: tCode,
        Name: String(r.Name || r.TrainingName || '').trim(),
        Category: String(r.Category || 'General').trim(),
        Trainer: String(r.Trainer || 'TBD').trim(),
        TrainingProvider: String(r.TrainingProvider || r.Provider || r.Vendor || '').trim(),
        Venue: String(r.Venue || 'TBD').trim(),
        StartDate: String(r.StartDate || '').trim(),
        EndDate: String(r.EndDate || r.StartDate || '').trim(),
        Duration: r.Duration || 1,
        TotalHours: r.TotalHours || 8,
        CourseFee: String(r.CourseFee || '0.00').trim(),
        Department: String(r.Department || 'N/A').trim(),
        ExpiryDate: String(r.ExpiryDate || r.CertExpiryDate || r.CertificateExpiryDate || '').trim(),
        CertExpiryDate: String(r.CertExpiryDate || r.ExpiryDate || r.CertificateExpiryDate || '').trim(),
        Objectives: String(r.Objectives || '').trim(),
        ApprovalStatus: String(r.ApprovalStatus || r.Status || 'Pending HOD Approval').trim(),
        ApprovalRemarks: String(r.ApprovalRemarks || '').trim(),
        ApprovedBy: String(r.ApprovedBy || '').trim(),
        RequestedDate: String(r.RequestedDate || r.CreatedDate || '').trim(),
        BrochureURL: String(r.BrochureURL || r.BrochureUrl || '').trim(),
        RequisitionFormFileID: String(r.RequisitionFormFileID || '').trim(),
        FolderID: String(r.FolderID || '').trim(),
        Participants: totalPaxVal,
        TotalPax: totalPaxVal,
        participants: reqParticipants
      };
    }).reverse();

    return ok(list);
  } catch(e) {
    Logger.log('getEmployeeSubmittedRequests error: ' + e.message);
    return err('Failed to load submitted requests: ' + e.message);
  }
}

/**
 * API: Submit or Edit/Resubmit Employee Training Requisition Request Form
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

    ensureTrainingSheetColumns(sheet);

    const editingTrainingId = String(data.editingTrainingId || data.ID || '').trim();
    const isEditing = Boolean(editingTrainingId);
    let targetRow = isEditing ? findRowById(sheet, editingTrainingId) : -1;
    if (isEditing && targetRow <= 1) {
      return err(`Cannot find existing training request with ID: ${editingTrainingId}`);
    }

    const headers = sheet.getDataRange().getValues()[0].map(h => String(h).trim());
    const getValFromRow = (rIdx, name) => {
      const idx = headers.indexOf(name);
      if (rIdx > 1 && idx >= 0) {
        return sheet.getRange(rIdx, idx + 1).getValue();
      }
      return '';
    };

    const id = isEditing ? editingTrainingId : generateId('TRN');
    const code = isEditing ? (String(getValFromRow(targetRow, 'Code')) || generateTrainingCode(data.Category || 'General')) : generateTrainingCode(data.Category || 'General');
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

    // Step 2: Compare assigned HOD in "HOD email" tab
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
    let approvalRemarksVal = isEditing ? `Resubmitted by employee on ${timeNow}.` : '';

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

    const objectivesVal = String(data.Objectives || data.Reason || '').trim();

    const requestedParticipants = Array.isArray(data.ParticipantList) ? data.ParticipantList : (Array.isArray(data.participants) ? data.participants : []);
    const participantResolution = canonicalizeTrainingParticipants(requestedParticipants);
    if (participantResolution.rejected.length > 0) {
      return err('The following participant(s) do not match an Employee-sheet record: ' + participantResolution.rejected.join(', '));
    }
    const participantsList = participantResolution.participants;

    let workspace = { folderId: '', folderUrl: '', partSheetId: '', sessionSheetId: '', attendanceSheetId: '', evaluationSheetId: '', postSheetId: '' };
    let reqForm = { fileId: '', fileUrl: '' };

    if (isEditing) {
      workspace.folderId = String(getValFromRow(targetRow, 'FolderID') || '');
      workspace.partSheetId = String(getValFromRow(targetRow, 'ParticipantsSheetID') || '');
      workspace.sessionSheetId = String(getValFromRow(targetRow, 'SessionsSheetID') || '');
      workspace.attendanceSheetId = String(getValFromRow(targetRow, 'AttendanceSheetID') || '');
      workspace.evaluationSheetId = String(getValFromRow(targetRow, 'EvaluationSheetID') || '');
      workspace.postSheetId = String(getValFromRow(targetRow, 'PostSheetID') || '');
      reqForm.fileId = String(getValFromRow(targetRow, 'RequisitionFormFileID') || '');
    }

    if (!workspace.folderId) {
      try {
        workspace = createTrainingWorkspace(code, data.TrainingName);
      } catch(wErr) {
        Logger.log('Workspace creation error: ' + wErr.message);
      }
    }

    const requesterSigData = {
      employeeNo: emp.ID || data.EmployeeID || data.EmployeeNo || '',
      name: emp.Name || data.EmployeeName || data.Name || 'Requester',
      position: emp.Position || emp.JobTitle || emp.PositionTitle || data.JobPosition || data.Position || data.JobTitle || 'Requester',
      date: getFormattedCurrentDate()
    };

    if (!reqForm.fileId) {
      try {
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
      } catch(fErr) {
        Logger.log('Form creation error: ' + fErr.message);
      }
    } else if (isEditing) {
      // Re-populate existing form sheet cells
      try {
        const ssForm = SpreadsheetApp.openById(reqForm.fileId);
        const shForm = ssForm.getSheetByName('Training Form') || ssForm.getSheets()[0];

        const setTemplateValue = (area, value) => shForm.getRange(area.split(':')[0]).setValue(value == null ? '' : value);
        shForm.getRange('A5').setValue('Training Title:');
        setTemplateValue('C5:I5', data.TrainingName || '');
        shForm.getRange('A6').setValue('Course Fee (RM):');
        setTemplateValue('C6:E6', data.CourseFee !== undefined ? data.CourseFee : '0.00');
        shForm.getRange('F6').setValue('Date:');
        setTemplateValue('G6:I6', data.EndDate && data.EndDate !== data.StartDate ? `${formatDate(data.StartDate)} - ${formatDate(data.EndDate)}` : (formatDate(data.StartDate) || ''));
        shForm.getRange('A7').setValue('Duration:');
        setTemplateValue('C7:E7', `${data.Duration || 1} day(s) (${data.TotalHours || 8} hours)`);
        shForm.getRange('F7').setValue('Venue:');
        setTemplateValue('G7:I7', data.Venue || '');
        shForm.getRange('A8').setValue('Training Provider:');
        setTemplateValue('C8:I8', data.TrainingProvider || data.Provider || data.Trainer || '');
        shForm.getRange('A10').setValue('Reasons for Training:');
        setTemplateValue('A11:I12', objectivesVal || '');

        if (participantsList.length > 0) {
          shForm.getRange('A15:I38').clearContent();
          participantsList.slice(0, 24).forEach((p, index) => {
            const r = 15 + index;
            setTemplateValue(`A${r}:B${r}`, p.EmployeeID || p.ID || p.EmployeeNo || '');
            setTemplateValue(`C${r}`, p.EmployeeName || p.Name || '');
            setTemplateValue(`D${r}:G${r}`, p.Department || p.CostCentre || '');
            setTemplateValue(`H${r}:I${r}`, p.Position || p.JobTitle || '');
          });
        }
        resetRequisitionFormApprovals(reqForm.fileId);
      } catch(reFormErr) {
        Logger.log('Error re-populating existing requisition form: ' + reFormErr.message);
      }
    }

    // Handle Brochure / Attachment File upload
    let brochureUrl = data.BrochureUrl || '';
    if (isEditing && !brochureUrl) {
      brochureUrl = String(getValFromRow(targetRow, 'BrochureURL') || '');
    }
    if (data.BrochureFile && data.BrochureFile.data) {
      try {
        const fileBlob = Utilities.newBlob(
          Utilities.base64Decode(data.BrochureFile.data),
          data.BrochureFile.mimeType || 'application/octet-stream',
          data.BrochureFile.name || `${code}_brochure`
        );
        let targetFolder = null;
        if (workspace.folderId) {
          try { targetFolder = DriveApp.getFolderById(workspace.folderId); } catch(fldErr) {}
        }
        if (!targetFolder) {
          targetFolder = getSystemRootFolder();
        }
        const driveFile = targetFolder.createFile(fileBlob);
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        brochureUrl = driveFile.getUrl();
      } catch(fErr) {
        Logger.log('Brochure file upload error: ' + fErr.message);
      }
    }

    let activeRowIndex = -1;
    if (isEditing) {
      activeRowIndex = targetRow;
    } else {
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
      activeRowIndex = sheet.getLastRow();
    }

    const setCol = (name, val) => {
      const idx = headers.indexOf(name) + 1;
      if (idx > 0) sheet.getRange(activeRowIndex, idx).setValue(val);
    };

    setCol('Name', data.TrainingName);
    setCol('Category', data.Category || 'General');
    setCol('Trainer', data.Trainer || 'TBD');
    setCol('TrainingProvider', data.TrainingProvider || data.Provider || '');
    setCol('Venue', data.Venue || 'TBD');
    setCol('StartDate', data.StartDate);
    setCol('EndDate', data.EndDate || data.StartDate);
    setCol('Duration', data.Duration || 1);
    setCol('TotalHours', data.TotalHours || 8);
    setCol('Department', data.Department || emp.Department || 'N/A');
    setCol('ExpiryDate', data.CertExpiryDate || data.ExpiryDate || '');
    setCol('CertExpiryDate', data.CertExpiryDate || data.ExpiryDate || '');
    setCol('Objectives', objectivesVal);
    setCol('Participants', data.TotalPax || participantsList.length || 0);
    setCol('CourseFee', data.CourseFee || '0.00');
    setCol('UpdatedDate', timeNow);
    setCol('ApprovalStatus', currentApprovalStatus);
    setCol('HOD', hodName || 'Pending');
    setCol('HODStatus', 'Pending');
    setCol('Csuite', csuiteName || 'N/A');
    setCol('CsuiteStatus', 'N/A');
    setCol('HOHR', hohrName || 'N/A');
    setCol('HOHRStatus', 'N/A');
    setCol('ApprovedBy', approvedByVal);
    setCol('ApprovedCostCentre', approvedCostCentreVal);
    setCol('ApprovedAt', approvedAtVal);
    setCol('ApprovalRemarks', approvalRemarksVal);
    setCol('RequestedBy', emp.ID || data.EmployeeID);
    setCol('RequestedByName', emp.Name || data.EmployeeName || '');
    setCol('RequestedByEmail', empEmail);
    if (!isEditing) setCol('RequestedDate', timeNow);
    setCol('FolderID', workspace.folderId || '');
    setCol('ParticipantsSheetID', workspace.partSheetId || '');
    setCol('SessionsSheetID', workspace.sessionSheetId || '');
    setCol('AttendanceSheetID', workspace.attendanceSheetId || '');
    setCol('EvaluationSheetID', workspace.evaluationSheetId || '');
    setCol('PostSheetID', workspace.postSheetId || '');
    setCol('RequisitionFormFileID', reqForm.fileId || '');
    setCol('BrochureURL', brochureUrl);
    setCol('ParticipantList', JSON.stringify(participantsList));

    SpreadsheetApp.flush();

    if (participantsList.length > 0) {
      try {
        syncParticipantsToTrainingDriveSheet(id, participantsList);
      } catch(sErr) {
        Logger.log('Error syncing participants to training drive sheet: ' + sErr.message);
      }
    }

    if (reqForm && reqForm.fileId) {
      try {
        updateTrainingRequisitionSignatures(id, 'request', requesterSigData, reqForm.fileId);
        syncTrainingRequisitionParticipants(id, participantsList);
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

    // Create Gmail Draft for HOD / Approver
    let draftStatus = 'Not created';
    try {
      const hodPortalUrl = getConfigProperty('HOD_PORTAL_URL', '');
      const reviewUrl = hodPortalUrl ? `${hodPortalUrl}?page=review&id=${id}` : getAppUrl();

      let recipientEmail = hodEmail;
      if (currentApprovalStatus === 'Pending C-Suite Approval' && csuiteEmail) recipientEmail = csuiteEmail;
      else if (currentApprovalStatus === 'Pending HOHR Approval' && hohrEmail) recipientEmail = hohrEmail;
      else if (currentApprovalStatus === 'Approved') recipientEmail = 'arina.ismail@apollofood.com.my';

      const subject = `Training Requisition — ${data.TrainingName} | ${id}`;
      const body = `Dear Approver / Manager,\n\nA Training Requisition Form (AP-HRD-F01-01) has been ${isEditing ? 'RESUBMITTED following updates by the employee' : 'submitted'}:\n\n` +
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

      const proposedDateStr = (data.EndDate && data.EndDate !== data.StartDate)
        ? `${formatDateForEmail(data.StartDate)} – ${formatDateForEmail(data.EndDate)}`
        : formatDateForEmail(data.StartDate);

      const htmlBody = buildTrainingRequisitionEmailHtml({
        requestId: id,
        trainingTitle: data.TrainingName,
        requesterName: emp.Name || data.EmployeeID,
        employeeId: emp.ID || data.EmployeeID,
        department: emp.Department || data.Department || 'N/A',
        category: data.Category || 'General',
        proposedDate: proposedDateStr,
        duration: formatDurationForEmail(data.Duration || 1, data.TotalHours || 8),
        estimatedFee: formatFeeForEmail(data.CourseFee),
        status: currentApprovalStatus,
        reviewUrl: reviewUrl,
        badgeText: 'ACTION REQUIRED',
        headlineText: 'Training Requisition Requires Your Review',
        greetingText: 'Dear Approver / Manager,',
        introText: `A new training requisition has been ${isEditing ? 'resubmitted following updates by the employee' : 'submitted'} and is currently awaiting your review and approval.`
      });

      if (recipientEmail) {
        try {
          const draft = GmailApp.createDraft(recipientEmail, subject, body, { htmlBody: htmlBody });
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
    } catch (mailErr) {
      draftStatus = `Notification error: ${mailErr.message}`;
      Logger.log(draftStatus);
    }

    try {
      syncTrainingById(id, 'Employee Requisition Submission', isEditing ? 'UPDATE' : 'CREATE');
    } catch(syncErr) {
      Logger.log('syncTrainingById error in submitEmployeeRequisition: ' + syncErr.message);
    }

    return ok({
      message: `Training Requisition Form (AP-HRD-F01-01) ${isEditing ? 'resubmitted' : 'submitted'}! Status: ${currentApprovalStatus}. HOD: ${hodName || 'Assigned HOD'}. ${draftStatus}`,
      trainingId: id,
      trainingCode: code,
      approvalStatus: currentApprovalStatus,
      assignedHod: hodName,
      draftStatus: draftStatus,
      isEditing: isEditing
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
