/**
 * ValidationService.gs — Server Validation & Company Email Domain Check
 */

function getValidEmployee(employeeId) {
  if (!employeeId || String(employeeId).trim() === '') {
    return { valid: false, message: 'Employee ID is required.' };
  }

  const emp = getEmployeeById(employeeId);
  if (emp) {
    return {
      valid: true,
      employee: emp
    };
  }

  return { valid: false, message: 'Employee record not found. Please contact HR/Administrator.' };
}

function validateCompanyEmail(email) {
  const allowedDomain = getConfigProperty('ALLOWED_DOMAIN', '');
  let cleanEmail = String(email || '').trim();

  if (!cleanEmail) {
    try {
      cleanEmail = Session.getActiveUser().getEmail() || '';
    } catch (e) {}
  }

  if (allowedDomain && cleanEmail && cleanEmail.includes('@')) {
    const domain = cleanEmail.split('@')[1] || '';
    if (domain.toLowerCase() !== allowedDomain.toLowerCase()) {
      return { valid: false, message: `Access restricted. Only company email accounts (@${allowedDomain}) can fill & submit the Training Requisition Form.` };
    }
  }

  return { valid: true, email: cleanEmail || 'employee@company.com' };
}
