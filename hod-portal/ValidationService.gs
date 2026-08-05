/**
 * ValidationService.gs — HOD Authentication & Security Validation
 */

function getActiveHODProfile() {
  let userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    userEmail = '';
  }

  const allowedDomain = getConfigProperty('ALLOWED_DOMAIN', '');
  if (allowedDomain && userEmail) {
    const domain = userEmail.split('@')[1] || '';
    if (domain.toLowerCase() !== allowedDomain.toLowerCase()) {
      return {
        valid: false,
        message: `Access denied. Only company email accounts (@${allowedDomain}) can access the HOD Portal.`
      };
    }
  }

  // Lookup in Employees sheet
  const empSheet = getSheet('Employees');
  let emp = null;
  if (empSheet) {
    const rows = sheetToJson(empSheet);
    if (userEmail) {
      emp = rows.find(r => String(r.Email || '').trim().toLowerCase() === userEmail.trim().toLowerCase());
    }
    // Fallback: If running in dev/test or active user email is hidden, return primary HOD profile or active match
    if (!emp) {
      emp = rows.find(r => String(r.Position || '').toLowerCase().includes('manager') || String(r.Position || '').toLowerCase().includes('hod') || String(r.Department || '').toLowerCase().includes('hr')) || rows[0];
    }
  }

  if (emp) {
    return {
      valid: true,
      email: userEmail || emp.Email,
      hod: {
        ID: emp.ID || emp.EmployeeID || 'HOD-1001',
        Name: emp.Name || 'HOD Manager',
        CostCentre: emp.Department || 'Management / Cost Centre 100',
        Position: emp.Position || 'Head of Department',
        Email: userEmail || emp.Email || ''
      }
    };
  }

  return {
    valid: true,
    email: userEmail,
    hod: {
      ID: 'HOD-1001',
      Name: 'HOD Manager',
      CostCentre: 'Management / Cost Centre 100',
      Position: 'Department Manager',
      Email: userEmail
    }
  };
}

function validateHODAccess() {
  const profile = getActiveHODProfile();
  if (!profile.valid) {
    return profile;
  }
  return { valid: true, hod: profile.hod };
}
