/**
 * ValidationService.gs — HOD Authentication & Security Validation
 */

function getActiveHODProfile() {
  let userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    userEmail = '';
  }

  const allowedDomain = getConfigProperty('ALLOWED_DOMAIN', '');
  if (allowedDomain && userEmail && userEmail.includes('@')) {
    const domain = userEmail.split('@')[1] || '';
    if (domain.toLowerCase() !== allowedDomain.toLowerCase()) {
      return {
        valid: false,
        message: `Access denied. Only company email accounts (@${allowedDomain}) can access the HOD Portal.`
      };
    }
  }

  const cleanUserEmail = userEmail.trim().toLowerCase();

  // Strict Check: Match against 'HOD email' tab in EMPLOYEE_SPREADSHEET_ID
  const hodSheet = getSheet('HOD email');
  if (hodSheet) {
    const hodRows = sheetToJson(hodSheet);
    
    // Find matching row where user email matches HODEmail, CsuiteEmail, or HohrEmail
    const matchedHodRecord = hodRows.find(h => {
      const hEmail = String(h.HODEmail || h.HodEmail || h.Email || h.HOD || '').trim().toLowerCase();
      const cEmail = String(h.CsuiteEmail || h.CSuiteEmail || '').trim().toLowerCase();
      const hrEmail = String(h.HohrEmail || h.HOHREmail || '').trim().toLowerCase();
      return (hEmail && hEmail === cleanUserEmail) || (cEmail && cEmail === cleanUserEmail) || (hrEmail && hrEmail === cleanUserEmail);
    });

    if (matchedHodRecord) {
      return {
        valid: true,
        email: userEmail,
        hod: {
          ID: matchedHodRecord.HODID || matchedHodRecord.ID || 'HOD-1001',
          Name: matchedHodRecord.HODName || matchedHodRecord.HodName || matchedHodRecord.HOD || 'HOD Approver',
          CostCentre: matchedHodRecord.Department || matchedHodRecord.CostCentre || 'All Departments',
          Position: 'Head of Department / Approver',
          Email: userEmail
        }
      };
    }

    // If active user email is present but NOT in the HOD email tab, deny access strictly
    if (cleanUserEmail) {
      return {
        valid: false,
        message: `Access Denied: Your email (${userEmail}) is not authorized in the "HOD email" tab in EMPLOYEE_SPREADSHEET_ID.`
      };
    }

    // Fallback for editor testing when user email is blank
    if (hodRows.length > 0) {
      const defaultHod = hodRows[0];
      return {
        valid: true,
        email: defaultHod.HODEmail || 'hod@company.com',
        hod: {
          ID: defaultHod.HODID || 'HOD-1001',
          Name: defaultHod.HODName || defaultHod.HodName || 'HOD Approver',
          CostCentre: defaultHod.Department || defaultHod.CostCentre || 'Management',
          Position: 'Head of Department',
          Email: defaultHod.HODEmail || 'hod@company.com'
        }
      };
    }
  }

  return {
    valid: false,
    message: 'Access Denied: HOD email registry sheet ("HOD email") is unavailable or empty.'
  };
}

function validateHODAccess() {
  const profile = getActiveHODProfile();
  if (!profile.valid) {
    return profile;
  }
  return { valid: true, hod: profile.hod };
}
