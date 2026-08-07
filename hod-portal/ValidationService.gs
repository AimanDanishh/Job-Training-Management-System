/**
 * ValidationService.gs — HOD Authentication & Security Validation
 */

/**
 * Resolves the real HOD Employee ID, Name, Cost Centre, Position, and Email 
 * by cross-referencing "For IT", "Employees", and "HOD email" tabs in EMPLOYEE_SPREADSHEET_ID.
 */
function resolveRealHODProfile(userEmail, requesterIdOrName) {
  let userEmailClean = String(userEmail || '').trim().toLowerCase();
  let userPrefix = userEmailClean ? userEmailClean.split('@')[0] : '';

  let realHodName = '';
  let realHodId = '';
  let realCostCentre = '';
  let realPosition = '';

  const getVal = (rowObj, nameList) => {
    if (!rowObj) return '';
    const keys = Object.keys(rowObj);
    for (let n of nameList) {
      const matchKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === n.toLowerCase().replace(/[^a-z0-9]/g, ''));
      if (matchKey && rowObj[matchKey]) return String(rowObj[matchKey]).trim();
    }
    return '';
  };

  // 1. Search in "For IT" tab using requester info or logged-in user email
  try {
    const itSheet = getSheet('For IT');
    if (itSheet) {
      const itRows = sheetToJson(itSheet);
      
      let matchedIt = null;
      if (requesterIdOrName) {
        const cleanReq = String(requesterIdOrName).trim().toLowerCase();
        matchedIt = itRows.find(r => {
          const rId = getVal(r, ['ID', 'EmployeeID', 'EmployeeNo', 'Employee ID', 'Employee No']).toLowerCase();
          const rName = getVal(r, ['Name', 'EmployeeName', 'Employee Name']).toLowerCase();
          return (rId && rId === cleanReq) || (rName && (rName.includes(cleanReq) || cleanReq.includes(rName)));
        });
      }

      if (!matchedIt && userEmailClean) {
        matchedIt = itRows.find(r => {
          const keys = Object.keys(r);
          for (let k of keys) {
            const val = String(r[k] || '').trim().toLowerCase();
            if (val && (val === userEmailClean || val === userPrefix)) return true;
          }
          return false;
        });
      }

      if (matchedIt) {
        realHodName = getVal(matchedIt, ['HOD', 'HODName', 'HodName', 'Manager', 'ReportTo', 'HOD Name']);
        realHodId = getVal(matchedIt, ['HODID', 'HOD ID', 'HODNo', 'HOD No', 'ID']);
        realCostCentre = getVal(matchedIt, ['Department', 'CostCentre', 'Cost Centre', 'CostCenter']);
      }
    }
  } catch (e) {
    Logger.log('For IT lookup error in resolveRealHODProfile: ' + e.message);
  }

  // 2. Search / Enrich in "Employees" tab
  try {
    const empSheet = getSheet('Employees');
    if (empSheet) {
      const empRows = sheetToJson(empSheet);
      
      const matchedEmp = empRows.find(e => {
        const eEmail = getVal(e, ['Email', 'EmailAddress', 'Email Address']).toLowerCase();
        const eId = getVal(e, ['ID', 'EmployeeID', 'EmployeeNo']).toLowerCase();
        const eName = getVal(e, ['Name', 'EmployeeName']).toLowerCase();

        if (userEmailClean && eEmail && eEmail === userEmailClean) return true;
        if (realHodId && eId && eId === realHodId.toLowerCase()) return true;
        if (realHodName && eName && (eName.includes(realHodName.toLowerCase()) || realHodName.toLowerCase().includes(eName))) return true;
        if (userPrefix && eName && userPrefix.length > 2 && (eName.includes(userPrefix) || userPrefix.includes(eName))) return true;
        return false;
      });

      if (matchedEmp) {
        if (!realHodId) realHodId = getVal(matchedEmp, ['ID', 'EmployeeID', 'EmployeeNo']);
        if (!realHodName) realHodName = getVal(matchedEmp, ['Name', 'EmployeeName']);
        if (!realCostCentre) realCostCentre = getVal(matchedEmp, ['Department', 'CostCentre', 'Cost Centre']);
        realPosition = getVal(matchedEmp, ['Position', 'JobTitle', 'Position Title', 'Title']);
      }
    }
  } catch (e) {
    Logger.log('Employees lookup error in resolveRealHODProfile: ' + e.message);
  }

  // 3. Search / Enrich in "HOD email" tab
  try {
    const hodSheet = getSheet('HOD email');
    if (hodSheet) {
      const hodRows = sheetToJson(hodSheet);
      const matchedHod = hodRows.find(h => {
        const keys = Object.keys(h);
        for (let k of keys) {
          const val = String(h[k] || '').trim().toLowerCase();
          if (val && (val === userEmailClean || (realHodName && val === realHodName.toLowerCase()))) return true;
        }
        return false;
      });

      if (matchedHod) {
        if (!realHodName) realHodName = getVal(matchedHod, ['HOD', 'HODName', 'HodName', 'Name']);
        if (!realCostCentre) realCostCentre = getVal(matchedHod, ['Department', 'CostCentre', 'Cost Centre']);
        if (!realHodId) realHodId = getVal(matchedHod, ['HODID', 'ID', 'EmployeeID']);
      }
    }
  } catch (e) {
    Logger.log('HOD email lookup error in resolveRealHODProfile: ' + e.message);
  }

  // Fallbacks if any fields are missing
  if (!userEmailClean) userEmailClean = 'hod@apollofood.com.my';
  if (!userPrefix) userPrefix = userEmailClean.split('@')[0];

  if (!realHodId) realHodId = `EMP-${userPrefix.toUpperCase()}`;
  if (!realHodName) {
    realHodName = userPrefix
      .split(/[\._\-]/)
      .filter(Boolean)
      .map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
      .join(' ') || userEmailClean;
  }
  if (!realCostCentre) realCostCentre = 'All Departments / Cost Centres';
  if (!realPosition) realPosition = 'Head of Department / Manager';

  return {
    ID: realHodId,
    Name: realHodName,
    CostCentre: realCostCentre,
    Position: realPosition,
    Email: userEmailClean
  };
}

function getActiveHODProfile(requesterIdOrName) {
  let userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    userEmail = '';
  }

  const allowedDomain = getConfigProperty('ALLOWED_DOMAIN', '');
  if (allowedDomain && userEmail && userEmail.includes('@')) {
    const domain = userEmail.split('@')[1] || '';
    if (domain.toLowerCase() !== allowedDomain.toLowerCase()) {
      return {
        valid: false,
        message: `Access denied. Only company email accounts (@${allowedDomain}) can access the HOD Portal.`
      };
    }
  }

  // Resolve real HOD Employee ID, Name, Cost Centre, Position via "For IT" -> "Employees" -> "HOD email"
  const realHod = resolveRealHODProfile(userEmail, requesterIdOrName);

  return {
    valid: true,
    email: userEmail,
    hod: realHod
  };
}

function validateHODAccess(requesterIdOrName) {
  const profile = getActiveHODProfile(requesterIdOrName);
  if (!profile.valid) {
    return profile;
  }
  return { valid: true, hod: profile.hod };
}
