/**
 * ValidationService.gs — Server Validation & Company Email Domain Check
 */

function getValidEmployee(employeeId) {
  if (!employeeId || String(employeeId).trim() === '') {
    return { valid: false, message: 'Employee ID is required.' };
  }
  const cleanEmpId = String(employeeId).trim().toLowerCase();

  // Primary Lookup: Employees sheet
  const empSheet = getSheet('Employees');
  if (empSheet) {
    const rows = sheetToJson(empSheet);
    const emp = rows.find(r => String(r.ID || r.EmployeeID || r.EmployeeNo || '').trim().toLowerCase() === cleanEmpId);
    if (emp) {
      return { valid: true, employee: emp };
    }
  }

  // Fallback: TrainingParticipants sheet
  try {
    const tpSheet = getSheet('TrainingParticipants');
    if (tpSheet) {
      const tpRows = sheetToJson(tpSheet);
      const tpEmp = tpRows.find(r => String(r.EmployeeID || r.EmployeeNo || r.ID || '').trim().toLowerCase() === cleanEmpId);
      if (tpEmp) {
        return {
          valid: true,
          employee: {
            ID: tpEmp.EmployeeID || String(employeeId).trim(),
            Name: tpEmp.EmployeeName || String(employeeId).trim(),
            Department: tpEmp.Department || tpEmp.CostCentre || 'N/A',
            Position: tpEmp.Position || 'Staff',
            Email: tpEmp.Email || ''
          }
        };
      }
    }
  } catch (e) {
    Logger.log('Fallback employee lookup error: ' + e.message);
  }

  return { valid: false, message: `Employee ID (${String(employeeId).trim()}) is not registered in the system.` };
}

function validateCompanyEmail(email) {
  const allowedDomain = getConfigProperty('ALLOWED_DOMAIN', '');
  if (!email || String(email).trim() === '') {
    return { valid: false, message: 'Official company email is required to submit a training requisition.' };
  }

  const cleanEmail = String(email).trim();
  if (allowedDomain) {
    const domain = cleanEmail.split('@')[1] || '';
    if (domain.toLowerCase() !== allowedDomain.toLowerCase()) {
      return { valid: false, message: `Access restricted. Only company email accounts (@${allowedDomain}) can fill & submit the Training Requisition Form.` };
    }
  }

  return { valid: true, email: cleanEmail };
}
