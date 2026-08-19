/**
 * Employee.gs - Employee / Participant management
 */

// --- Read Master Employee List --------------------------------------------------
function getEmployees() {
  try {
    const cached = getCachedData(CACHE_NAMESPACES.EMPLOYEES_ALL);
    if (cached) return ok(cached);

    const sheet = getSheet(SHEET_NAMES.employees);
    const rows = sheetToJson(sheet);
    if (Array.isArray(rows)) {
      setCachedData(CACHE_NAMESPACES.EMPLOYEES_ALL, rows, 3600); // 1 hour TTL
    }
    return ok(rows);
  } catch (e) {
    return err('Failed to load employees: ' + e.message);
  }
}

/**
 * Get list of Cost Centres / Departments from EMPLOYEE_SPREADSHEET_ID or Employees sheet
 */
function getCostCentres() {
  try {
    const cached = getCachedData(CACHE_NAMESPACES.EMPLOYEES_CC);
    if (cached) return ok(cached);

    const ccMap = new Map();

    const addIfValid = (val) => {
      if (!val) return;
      const str = String(val).trim();
      if (!str || str === 'All Departments / Cost Centres' || str.toLowerCase() === 'all departments' || str.toLowerCase() === 'all') return;
      const key = str.toLowerCase();
      if (!ccMap.has(key)) {
        ccMap.set(key, str);
      }
    };

    try {
      const ss = getEmployeeSpreadsheet();
      if (ss) {
        const ccSheet = ss.getSheetByName('Cost Centre') || ss.getSheetByName('CostCentre') || ss.getSheetByName('Cost Centres');
        if (ccSheet && ccSheet.getLastRow() > 1) {
          const data = ccSheet.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
            addIfValid(data[i][0] || data[i][1]);
          }
        }

        const empSheet = ss.getSheetByName(SHEET_NAMES.employees) || ss.getSheetByName('Employees');
        if (empSheet && empSheet.getLastRow() > 1) {
          const rows = sheetToJson(empSheet);
          rows.forEach(r => {
            addIfValid(r.Department || r.CostCentre || r['Cost Centre'] || r.Dept);
          });
        }
      }
    } catch (e) {
      Logger.log('getCostCentres external spreadsheet error: ' + e.message);
    }

    try {
      const localEmpSheet = getSheet(SHEET_NAMES.employees);
      if (localEmpSheet && localEmpSheet.getLastRow() > 1) {
        const rows = sheetToJson(localEmpSheet);
        rows.forEach(r => {
          addIfValid(r.Department || r.CostCentre || r['Cost Centre'] || r.Dept);
        });
      }
    } catch (e2) {
      Logger.log('getCostCentres local employee error: ' + e2.message);
    }

    try {
      const trainSheet = getSheet(SHEET_NAMES.trainings);
      if (trainSheet && trainSheet.getLastRow() > 1) {
        const rows = sheetToJson(trainSheet);
        rows.forEach(r => {
          addIfValid(r.Department || r.CostCentre);
        });
      }
    } catch (e3) {}

    const result = Array.from(ccMap.values()).sort();
    if (result.length > 0) {
      setCachedData(CACHE_NAMESPACES.EMPLOYEES_CC, result, 7200); // 2 hours TTL
    }
    return ok(result);
  } catch (e) {
    return err('Failed to get cost centres: ' + e.message);
  }
}

/**
 * Add new employee record
 */
function addEmployee(data) {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    if (!sheet) return err('Employees sheet not found.');
    const empId = data.ID || generateId('EMP');
    sheet.appendRow([
      empId,
      data.Name || '',
      data.Department || '',
      data.Position || '',
      data.Status || 'Active',
      data.Email || '',
      data.Phone || '',
      now()
    ]);
    invalidateEmployeeCaches();
    return ok({ ID: empId, message: 'Employee added successfully.' });
  } catch(e) {
    return err('Failed to add employee: ' + e.message);
  }
}

/**
 * Update existing employee record
 */
function updateEmployee(data) {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    if (!sheet) return err('Employees sheet not found.');
    const rows = sheetToJson(sheet);
    const emp = rows.find(r => String(r.ID) === String(data.ID));
    if (!emp || !emp._row) return err('Employee not found.');
    
    if (data.Name !== undefined)       sheet.getRange(emp._row, 2).setValue(data.Name);
    if (data.Department !== undefined) sheet.getRange(emp._row, 3).setValue(data.Department);
    if (data.Position !== undefined)   sheet.getRange(emp._row, 4).setValue(data.Position);
    if (data.Status !== undefined)     sheet.getRange(emp._row, 5).setValue(data.Status);
    if (data.Email !== undefined)      sheet.getRange(emp._row, 6).setValue(data.Email);
    if (data.Phone !== undefined)      sheet.getRange(emp._row, 7).setValue(data.Phone);
    sheet.getRange(emp._row, 8).setValue(now());
    SpreadsheetApp.flush();
    invalidateEmployeeCaches();
    return ok({ ID: data.ID, message: 'Employee updated successfully.' });
  } catch(e) {
    return err('Failed to update employee: ' + e.message);
  }
}

/**
 * Delete employee record
 */
function deleteEmployee(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    if (!sheet) return err('Employees sheet not found.');
    const rows = sheetToJson(sheet);
    const emp = rows.find(r => String(r.ID) === String(id));
    if (!emp || !emp._row) return err('Employee not found.');
    sheet.deleteRow(emp._row);
    SpreadsheetApp.flush();
    invalidateEmployeeCaches();
    return ok({ ID: id, message: 'Employee deleted successfully.' });
  } catch(e) {
    return err('Failed to delete employee: ' + e.message);
  }
}
