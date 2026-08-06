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
    const emp = rows.find(r => {
      const idVal = String(r.ID || r.EmployeeID || r.EmployeeNo || r.EmpID || r.StaffID || r['Employee ID'] || r['Employee No'] || r['Staff ID'] || '').trim().toLowerCase();
      return idVal === cleanEmpId;
    });
    if (emp) {
      return {
        valid: true,
        employee: {
          ID: emp.ID || emp.EmployeeID || emp.EmployeeNo || String(employeeId).trim(),
          Name: emp.Name || emp.EmployeeName || emp['Employee Name'] || emp['Staff Name'] || '',
          Department: emp.Department || emp.CostCentre || emp['Cost Centre'] || 'N/A',
          Position: emp.Position || emp.JobTitle || emp.PositionTitle || emp['Position Title'] || emp['Job Title'] || 'Staff',
          Email: emp.Email || emp['Email Address'] || ''
        }
      };
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
            Position: tpEmp.Position || tpEmp.JobTitle || tpEmp.PositionTitle || 'Staff',
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
