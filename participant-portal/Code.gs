/**
 * Code.gs — Public Participant Portal Router & Web App Entry Point
 * 
 * Target Deployment Settings:
 * - Execute as: Me
 * - Anyone with the link (No login required)
 */

function doGet(e) {
  const modeParam = (e && e.parameter && (e.parameter.mode || e.parameter.page)) ? String(e.parameter.mode || e.parameter.page).toLowerCase().trim() : 'attendance';
  
  // Context Page Mapping
  const pageMap = {
    'attendance': 'Attendance',
    'evaluation': 'TrainingEvaluation',
    'post':       'PostEvaluation',
    'success':    'Success',
    'error':      'Error'
  };

  const templateName = pageMap[modeParam] || 'Attendance';
  const appTitle = getConfigProperty('APP_TITLE', 'TrainHub — Participant Portal');

  try {
    const template = HtmlService.createTemplateFromFile(templateName);
    
    // Pass raw query parameters into client template context
    template.params = (e && e.parameter) ? e.parameter : {};
    template.mode = modeParam;
    template.page = modeParam;
    try {
      const rawLogo = getCompanyLogoUrl();
      const directLogo = convertDriveLinkToDirectImageUrl(rawLogo);
      template.companyLogoUrl = directLogo ? Utilities.base64Encode(directLogo, Utilities.Charset.UTF_8) : '';
    } catch(logoErr) {
      template.companyLogoUrl = '';
    }
    try {
      const rawSysLogo = getSystemLogoUrl() || '';
      const directSysLogo = rawSysLogo ? (convertDriveLinkToDirectImageUrl(rawSysLogo) || rawSysLogo) : '';
      template.systemLogoUrl = String(directSysLogo);
    } catch(sysLogoErr) {
      template.systemLogoUrl = '';
    }

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
        .setTitle('Error — TrainHub')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (fallbackErr) {
      return HtmlService.createHtmlOutput('<h3 style="font-family:sans-serif;color:#ef4444;padding:20px;">System Error: ' + err.message + '</h3>');
    }
  }
}

/**
 * Include helper for template partials (style, script, etc.)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

