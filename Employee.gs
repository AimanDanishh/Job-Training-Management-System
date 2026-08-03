/**
 * Employee.gs — Employee / Participant management
 */

// ─── Read ───────────────────────────────────────────────────────────────────────
function getEmployees() {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    const rows = sheetToJson(sheet);
    return ok(rows);
  } catch (e) {
    return err('Failed to load employees: ' + e.message);
  }
}

function getEmployeeById(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    const rows = sheetToJson(sheet);
    const emp = rows.find(r => r.ID === id);
    if (!emp) return err('Employee not found.');
    return ok(emp);
  } catch (e) {
    return err(e.message);
  }
}

// ─── Create ─────────────────────────────────────────────────────────────────────
function addEmployee(data) {
  try {
    if (!data.Name || !data.Department)
      return err('Name and Department are required.');

    const sheet = getSheet(SHEET_NAMES.employees);
    const id = generateId('EMP');

    sheet.appendRow([
      id,
      data.Name,
      data.Department,
      data.Position   || '',
      data.Email      || '',
      data.Phone      || '',
      data.Status     || 'Active'
    ]);
    return ok({ id: id, message: 'Employee added successfully.' });
  } catch (e) {
    return err('Failed to add employee: ' + e.message);
  }
}

// ─── Update ─────────────────────────────────────────────────────────────────────
function updateEmployee(data) {
  try {
    if (!data.ID) return err('Employee ID is required.');

    const sheet = getSheet(SHEET_NAMES.employees);
    const row = findRowById(sheet, data.ID);
    if (row === -1) return err('Employee not found.');

    sheet.getRange(row, 1, 1, 7).setValues([[
      data.ID,
      data.Name        || '',
      data.Department  || '',
      data.Position    || '',
      data.Email       || '',
      data.Phone       || '',
      data.Status      || 'Active'
    ]]);
    return ok({ message: 'Employee updated successfully.' });
  } catch (e) {
    return err('Failed to update employee: ' + e.message);
  }
}

// ─── Delete ─────────────────────────────────────────────────────────────────────
function deleteEmployee(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    const row = findRowById(sheet, id);
    if (row === -1) return err('Employee not found.');
    sheet.deleteRow(row);
    return ok({ message: 'Employee deleted.' });
  } catch (e) {
    return err('Failed to delete employee: ' + e.message);
  }
}

// ─── Search ─────────────────────────────────────────────────────────────────────
function searchEmployees(query) {
  try {
    const sheet = getSheet(SHEET_NAMES.employees);
    const rows = sheetToJson(sheet);
    const q = query.toLowerCase();
    const filtered = rows.filter(r =>
      r.Name.toLowerCase().includes(q) ||
      r.Department.toLowerCase().includes(q) ||
      r.ID.toLowerCase().includes(q)
    );
    return ok(filtered);
  } catch (e) {
    return err(e.message);
  }
}