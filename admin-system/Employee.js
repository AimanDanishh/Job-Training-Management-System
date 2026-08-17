/**
 * Employee.gs — Employee / Participant management
 */

// ─── Read Master Employee List ──────────────────────────────────────────────────
function getEmployees() {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    const rows = sheetToJson(sheet);
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
    const ss = getEmployeeSpreadsheet();
    if (ss) {
      const ccSheet = ss.getSheetByName('Cost Centre') || ss.getSheetByName('CostCentre') || ss.getSheetByName('Cost Centres');
      if (ccSheet && ccSheet.getLastRow() > 1) {
        const data = ccSheet.getDataRange().getValues();
        const ccMap = new Map(); // key: lowercased, val: canonical name
        for (let i = 1; i < data.length; i++) {
          const val = String(data[i][0] || data[i][1] || '').trim();
          if (val && val !== 'All Departments / Cost Centres') {
            const key = val.toLowerCase();
            if (!ccMap.has(key)) {
              ccMap.set(key, val);
            }
          }
        }
        if (ccMap.size > 0) return ok(Array.from(ccMap.values()));
      }
    }
  } catch (e) {
    Logger.log('getCostCentres error: ' + e.message);
  }
  return ok([]);
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
    return ok({ ID: id, message: 'Employee deleted successfully.' });
  } catch(e) {
    return err('Failed to delete employee: ' + e.message);
  }
}