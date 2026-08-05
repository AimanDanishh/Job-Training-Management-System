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