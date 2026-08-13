/**
 * Code.gs — Main entry point, page router, and Google Workspace Auth
 * Place this file in your Apps Script project root.
 */

// ─── Web App Entry Point ───────────────────────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  const allowedPages = ['index', 'dashboard', 'training', 'attendance', 'evaluation', 'report', 'session', 'settings'];

  const safePage = allowedPages.includes(page) ? page : 'index';
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — Training Management System');

  try {
    const template = HtmlService.createTemplateFromFile(safePage);
    // Pass page parameter, session parameters, and public portal URL into the template
    template.currentPage = safePage;
    template.sessionParam = (e && e.parameter && e.parameter.session) ? e.parameter.session : '';
    template.queryParams = (e && e.parameter) ? JSON.stringify(e.parameter) : '{}';
    template.publicPortalUrl = getPublicPortalUrl();

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

  // 2. Check Administrator / HR requirement (ADMIN_EMAILS + HR email tab)
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
    name: hrProfile.name,
    employeeNo: hrProfile.employeeNo,
    position: hrProfile.position,
    domain: domain,
    role: 'admin'
  };
}

/**
 * Gets the active Google user's profile and authentication status on page load.
 */
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
    publicPortalUrl: getConfigProperty('PUBLIC_PORTAL_URL', ''),
    hodPortalUrl: getConfigProperty('HOD_PORTAL_URL', ''),
    allowedDomain: getConfigProperty('ALLOWED_DOMAIN', ''),
    adminEmails: getConfigProperty('ADMIN_EMAILS', ''),
    appTitle: getConfigProperty('APP_TITLE', '')
  });
}

/**
 * Runs database setup safely from frontend Settings page.
 */
function runDatabaseSetup(spreadsheetId, employeeSpreadsheetId) {
  const result = setupDatabase(spreadsheetId, employeeSpreadsheetId);
  if (result.success) {
    return ok(result);
  }
  return err(result.message);
}

