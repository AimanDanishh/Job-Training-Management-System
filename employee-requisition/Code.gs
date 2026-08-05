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
 * API: Lookup Employee Details by Employee ID for Auto-Filling Requisition Form
 */
function getEmployeeDetails(employeeId) {
  try {
    const empCheck = getValidEmployee(employeeId);
    if (!empCheck.valid) return err(empCheck.message);

    const emp = empCheck.employee;
    const requestDate = formatDate(new Date());

    return ok({
      ID: emp.ID || employeeId,
      Name: emp.Name || '',
      Department: emp.Department || emp.CostCentre || 'N/A',
      Position: emp.Position || 'Staff',
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

    const rowData = [
      id,
      code,
      data.TrainingName,
      data.Category || 'General',
      data.Trainer || 'TBD',
      data.Venue || 'TBD',
      data.StartDate,
      data.EndDate || data.StartDate,
      data.Duration || 1,
      data.TotalHours || 8,
      emp.Department || data.Department || 'N/A',
      data.Objectives || '',
      'Draft',
      'Created',
      0,
      '', '', '', '', '', '', '',
      timeNow,
      timeNow,
      data.CourseFee || '0.00'
    ];

    sheet.appendRow(rowData);
    const row = sheet.getLastRow();
    const headers = sheet.getDataRange().getValues()[0].map(h => String(h).trim());

    const setCol = (colName, val) => {
      const idx = headers.indexOf(colName) + 1;
      if (idx > 0) sheet.getRange(row, idx).setValue(val);
    };

    setCol('ApprovalStatus', 'Pending Approval');
    setCol('RequestedBy', emp.ID || data.EmployeeID);
    setCol('RequestedDate', timeNow);

    SpreadsheetApp.flush();

    // 3. Trigger email notification to HOD/Manager with link to hod-portal review page
    try {
      const hodPortalUrl = getConfigProperty('HOD_PORTAL_URL', '');
      const reviewUrl = hodPortalUrl ? `${hodPortalUrl}?page=review&id=${id}` : getAppUrl();

      const hodEmail = data.HodEmail || getConfigProperty('ADMIN_EMAILS', '');
      if (hodEmail) {
        const subject = `[TrainHub] New Training Requisition Approval Required - ${data.TrainingName}`;
        const body = `Dear HOD / Manager,\n\nAn employee under your department has submitted a new Training Requisition Form (AP-HRD-F01-00):\n\n` +
          `Requester: ${emp.Name || data.EmployeeID} (${emp.Department || 'N/A'})\n` +
          `Employee ID: ${emp.ID || data.EmployeeID}\n` +
          `Training Name: ${data.TrainingName}\n` +
          `Category: ${data.Category || 'General'}\n` +
          `Proposed Date: ${data.StartDate} to ${data.EndDate || data.StartDate}\n` +
          `Estimated Fee: RM ${data.CourseFee || '0.00'}\n\n` +
          `Please click the link below to review the request details and digitally approve/reject/postpone:\n${reviewUrl}\n\n` +
          `Thank you,\nTrainHub Training Management System`;

        MailApp.sendEmail(hodEmail, subject, body);
      }
    } catch (mailErr) {
      Logger.log('HOD notification mail error: ' + mailErr.message);
    }

    return ok({
      message: 'Training Requisition Form (AP-HRD-F01-00) successfully submitted! An email notification has been sent to your HOD/Manager for review.',
      trainingId: id,
      trainingCode: code
    });
  } catch (e) {
    Logger.log('submitEmployeeRequisition error: ' + e.message);
    return err('Failed to submit training requisition: ' + e.message);
  }
}
