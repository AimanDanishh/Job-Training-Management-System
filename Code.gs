/**
 * Code.gs — Main entry point, page router, and Google Workspace Auth
 * Place this file in your Apps Script project root.
 */

// ─── Web App Entry Point ───────────────────────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  const allowedPages = ['index', 'dashboard', 'training', 'attendance', 'evaluation', 'report', 'session'];

  const safePage = allowedPages.includes(page) ? page : 'index';
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — Training Management System');

  try {
    const template = HtmlService.createTemplateFromFile(safePage);
    // Pass page parameter and session parameters into the template
    template.currentPage = safePage;
    template.sessionParam = (e && e.parameter && e.parameter.session) ? e.parameter.session : '';
    template.queryParams = (e && e.parameter) ? JSON.stringify(e.parameter) : '{}';

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

  // 2. Check Administrator requirement (if ADMIN_EMAILS is configured)
  if (adminEmails.length > 0 && !adminEmails.includes(userEmail)) {
    return {
      success: false,
      authenticated: false,
      email: userEmail,
      message: `Access Denied: Account (${userEmail}) is not authorized as an Administrator.`
    };
  }

  // Successfully authorized admin
  const formattedName = userEmail.split('@')[0].replace(/[\._]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());

  return {
    success: true,
    authenticated: true,
    email: userEmail,
    name: formattedName,
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

// ─── Legacy / Fallback Auth Handler ─────────────────────────────────────────────
function verifyLogin(username, password) {
  if (username && username.includes('@')) {
    return verifyGoogleUser(username);
  }

  const adminUser = getConfigProperty('ADMIN_USER', 'admin');
  const adminPass = getConfigProperty('ADMIN_PASS', 'admin123');

  if (username === adminUser && password === adminPass) {
    return { success: true, name: 'Administrator', role: 'admin', email: username };
  }
  return { success: false, message: 'Invalid username or password.' };
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}
