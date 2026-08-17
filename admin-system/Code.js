/**
 * Code.gs — Main entry point, page router, and Google Workspace Auth
 * Place this file in your Apps Script project root.
 */

// ─── Web App Entry Point ───────────────────────────────────────────────────────
function doGet(e) {
  // Normalise the query value before routing.  Apps Script passes query
  // parameters as strings, but links copied from email can contain mixed case
  // or surrounding whitespace.  Without this normalisation those links were
  // silently sent back to the sign-in page.
  const requestedPage = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  const page = String(requestedPage || 'index').toLowerCase().trim();
  const allowedPages = [
    'index', 'dashboard', 'training', 'attendance', 'evaluation',
    'report', 'session', 'settings', 'employee'
  ];

  const safePage = allowedPages.includes(page) ? page : 'index';
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — Training Management System');

  try {
    const template = HtmlService.createTemplateFromFile(safePage);
    // Pass page parameter, session parameters, and public portal URL into the template
    template.currentPage = safePage;
    template.sessionParam = (e && e.parameter && e.parameter.session) ? e.parameter.session : '';
    template.queryParams = (e && e.parameter) ? JSON.stringify(e.parameter) : '{}';
    const rawPublicUrl = getPublicPortalUrl() || '';
    template.publicPortalUrl = String(rawPublicUrl).replace(/"/g, '');
    try {
      const rawLogo = getCompanyLogoUrl() || '';
      const directLogo = convertDriveLinkToDirectImageUrl(rawLogo) || '';
      template.companyLogoUrl = String(directLogo).replace(/"/g, '');
    } catch(logoErr) {
      template.companyLogoUrl = '';
    }

    return template.evaluate()
      .setTitle(appTitle)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log('doGet error: ' + err.message);
    return HtmlService.createHtmlOutput('<p>Error loading page: ' + err.message + '</p>');
  }
}

// ─── Template Include Helper ────────────────────────────────────────────────────
/**
 * Include another HTML file's content (used inside templates with <?!= include('file') ?>)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── Google Workspace / Admin Auth Handler ──────────────────────────────────────
/**
 * Verifies a Google Account against company domain and administrator requirements.
 */
function verifyGoogleUser(emailInput) {
  let email = emailInput;
  if (!email) {
    try {
      email = Session.getActiveUser().getEmail();
    } catch (e) {
      Logger.log('Could not fetch active user email: ' + e.message);
    }
  }

  const allowedDomain = getConfigProperty('ALLOWED_DOMAIN', '').toLowerCase().trim();
  const adminEmailsStr = getConfigProperty('ADMIN_EMAILS', '').toLowerCase().trim();
  const adminEmails = adminEmailsStr ? adminEmailsStr.split(',').map(e => e.trim()).filter(Boolean) : [];

  const hrRecords = getHrEmailRecords();
  const hrEmails = hrRecords.map(r => r.email).filter(Boolean);

  if (!email) {
    return {
      success: false,
      authenticated: false,
      message: 'No Google Account detected. Please sign in with your Google Account.'
    };
  }

  const userEmail = email.toLowerCase().trim();
  const domain = userEmail.split('@')[1] || '';

  // 1. Check Company Domain requirement (if ALLOWED_DOMAIN is configured)
  if (allowedDomain && domain !== allowedDomain) {
    return {
      success: false,
      authenticated: false,
      email: userEmail,
      message: `Access Denied: Only @${allowedDomain} company accounts are allowed to sign in. (Detected: ${userEmail})`
    };
  }

  // 2. Check Administrator / HR requirement (if ADMIN_EMAILS or HR email tab are configured)
  const isAuthorized = adminEmails.includes(userEmail) || hrEmails.includes(userEmail);
  const hasRestrictions = adminEmails.length > 0 || hrEmails.length > 0;

  if (hasRestrictions && !isAuthorized) {
    return {
      success: false,
      authenticated: false,
      email: userEmail,
      message: `Access Denied: Account (${userEmail}) is not authorized as an Administrator or HR.`
    };
  }

  // Successfully authorized admin/HR profile lookup
  const hrProfile = getHrProfileByEmail(userEmail);

  return {
    success: true,
    authenticated: true,
    email: userEmail,
    name: hrProfile ? hrProfile.name : (userEmail.split('@')[0].replace(/[\._\-]/g, ' ').replace(/\b\w/g, c => c.toUpperCase())),
    employeeNo: hrProfile ? hrProfile.employeeNo : 'ADMIN',
    position: hrProfile ? hrProfile.position : 'Administrator',
    domain: domain,
    role: 'admin'
  };
}

function getCurrentUser() {
  let email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch(e) {}
  return verifyGoogleUser(email);
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ─── Settings API Services ─────────────────────────────────────────────────────
/**
 * Returns structured configuration and system validation for the Settings UI.
 */
function getSettingsData() {
  const validation = validateSystemConfiguration();
  const diagnostics = getSystemConfigurationDiagnostics();
  return ok({
    validation: validation,
    diagnostics: diagnostics,
    databaseId: getConfigProperty('SPREADSHEET_ID', ''),
    employeeSpreadsheetId: getConfigProperty('EMPLOYEE_SPREADSHEET_ID', ''),
    rootFolderId: getConfigProperty('ROOT_FOLDER_ID', ''),
    trainingFolder: getConfigProperty('TRAINING_FOLDER', '') || getConfigProperty('TRAINING_FOLDER_ID', ''),
    requisitionTemplateId: getConfigProperty('TRAINING_REQUISITION_TEMPLATE_ID', ''),
    publicPortalUrl: getConfigProperty('PUBLIC_PORTAL_URL', ''),
    hodPortalUrl: getConfigProperty('HOD_PORTAL_URL', ''),
    allowedDomain: getConfigProperty('ALLOWED_DOMAIN', ''),
    adminEmails: getConfigProperty('ADMIN_EMAILS', ''),
    companyLogoUrl: getConfigProperty('COMPANY_LOGO_URL', ''),
    appTitle: getConfigProperty('APP_TITLE', '')
  });
}

/**
 * Saves all system settings and configuration properties to Script Properties.
 */
function saveSettings(data) {
  try {
    if (!data || typeof data !== 'object') {
      return err('Invalid configuration data.');
    }

    if (data.databaseId !== undefined) {
      setConfigProperty('SPREADSHEET_ID', String(data.databaseId).trim());
    }
    if (data.employeeSpreadsheetId !== undefined) {
      setConfigProperty('EMPLOYEE_SPREADSHEET_ID', String(data.employeeSpreadsheetId).trim());
    }
    if (data.rootFolderId !== undefined) {
      setConfigProperty('ROOT_FOLDER_ID', String(data.rootFolderId).trim());
    }
    if (data.trainingFolder !== undefined) {
      setConfigProperty('TRAINING_FOLDER', String(data.trainingFolder).trim());
    }
    if (data.requisitionTemplateId !== undefined) {
      setConfigProperty('TRAINING_REQUISITION_TEMPLATE_ID', String(data.requisitionTemplateId).trim());
    }
    if (data.publicPortalUrl !== undefined) {
      setConfigProperty('PUBLIC_PORTAL_URL', String(data.publicPortalUrl).trim());
    }
    if (data.hodPortalUrl !== undefined) {
      setConfigProperty('HOD_PORTAL_URL', String(data.hodPortalUrl).trim());
    }
    if (data.allowedDomain !== undefined) {
      setConfigProperty('ALLOWED_DOMAIN', String(data.allowedDomain).trim());
    }
    if (data.adminEmails !== undefined) {
      setConfigProperty('ADMIN_EMAILS', String(data.adminEmails).trim());
    }
    if (data.companyLogoUrl !== undefined) {
      setConfigProperty('COMPANY_LOGO_URL', String(data.companyLogoUrl).trim());
    }
    if (data.appTitle !== undefined) {
      setConfigProperty('APP_TITLE', String(data.appTitle).trim());
    }

    // Invalidate cached spreadsheet objects so new IDs take effect immediately
    if (typeof _cachedSpreadsheet !== 'undefined') _cachedSpreadsheet = null;
    if (typeof _cachedEmployeeSpreadsheet !== 'undefined') _cachedEmployeeSpreadsheet = null;

    Logger.log('System settings successfully saved to Script Properties.');
    return getSettingsData();
  } catch (e) {
    Logger.log('saveSettings error: ' + e.message);
    return err('Failed to save settings: ' + e.message);
  }
}

/**
 * Runs structural auto-creation setup from Settings page if requested.
 */
function runDatabaseSetup(spreadsheetId, employeeSpreadsheetId) {
  const result = setupDatabase(spreadsheetId, employeeSpreadsheetId);
  if (result.success) {
    return ok(result);
  }
  return err(result.message);
}
