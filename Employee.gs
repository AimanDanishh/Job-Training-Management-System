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

// ─── Create, Update & Delete (Disabled: Master Employee Registry is Read-Only) ───
function addEmployee(data) {
  return err('Adding new employees is disabled.');
}

function updateEmployee(data) {
  return err('Updating employee records is disabled.');
}

function deleteEmployee(id) {
  return err('Deleting employee records is disabled.');
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