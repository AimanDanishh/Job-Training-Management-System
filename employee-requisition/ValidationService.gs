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
  const allowedDomainStr = getConfigProperty('ALLOWED_DOMAIN', '').toLowerCase().trim();
  const allowedDomains = allowedDomainStr
    ? allowedDomainStr.split(',').map(d => d.trim().replace(/^@/, '')).filter(Boolean)
    : [];

  let cleanEmail = String(email || '').trim();

  if (!cleanEmail) {
    try {
      cleanEmail = Session.getActiveUser().getEmail() || '';
    } catch (e) {}
  }

  if (allowedDomains.length > 0 && cleanEmail && cleanEmail.includes('@')) {
    const domain = (cleanEmail.split('@')[1] || '').toLowerCase().trim();
    if (!allowedDomains.includes(domain)) {
      const domainLabels = allowedDomains.map(d => '@' + d).join(' or ');
      return { valid: false, message: `Access restricted. Only company email accounts (${domainLabels}) can fill & submit the Training Requisition Form.` };
    }
  }

  return { valid: true, email: cleanEmail || 'employee@company.com' };
}
