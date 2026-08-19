/**
 * Resolves the real HOD Employee ID, Name, Cost Centre, Position, and Email 
 * using the "HOD email" tab schema:
 * Header: Employee No | HOD | Cost Centre | Position Title | Email
 * 
 * Returns null if userEmail is not matched in the "HOD email" directory.
 */
function resolveRealHODProfile(userEmail, requesterIdOrName) {
  let userEmailClean = String(userEmail || '').trim().toLowerCase();
  if (!userEmailClean) return null;

  let realHodName = '';
  let realHodId = '';
  let realCostCentre = '';
  let realPosition = '';
  let csuiteName = '';
  let csuiteEmail = '';
  let hohrName = '';
  let hohrEmail = '';
  let sourceTabName = '';

  const getVal = (rowObj, nameList) => {
    if (!rowObj) return '';
    const keys = Object.keys(rowObj);
    for (let n of nameList) {
      const matchKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === n.toLowerCase().replace(/[^a-z0-9]/g, ''));
      if (matchKey && rowObj[matchKey] !== undefined && rowObj[matchKey] !== null) {
        return String(rowObj[matchKey]).trim();
      }
    }
    return '';
  };

  const emailAliases = ['Email', 'EmailAddress', 'HODEmail', 'HOD Email', 'Email Address', 'HOD Email Address', 'E-Mail', 'Mail'];
  const idAliases = ['Employee No', 'EmployeeNo', 'EmployeeID', 'ID', 'HODID', 'Employee ID', 'Staff ID'];
  const nameAliases = ['HOD', 'HODName', 'HodName', 'Name', 'HOD Name', 'Employee Name'];

  // STEP 1: Direct lookup in "HOD email", "Csuite email", "HOHR email", and "HR email" directory tabs
  try {
    const targetSheets = ['HOD email', 'Csuite email', 'C-Suite email', 'HOHR email', 'HOHR Email', 'HR email'];
    let matchedHod = null;

    for (let tabName of targetSheets) {
      const sheet = getSheet(tabName);
      if (sheet) {
        const rows = sheetToJson(sheet);
        matchedHod = rows.find(h => {
          const emailVal = getVal(h, emailAliases).toLowerCase().trim();
          const idVal = getVal(h, idAliases).toLowerCase().trim();
          const nameVal = getVal(h, nameAliases).toLowerCase().trim();

          if (emailVal) {
            if (emailVal === userEmailClean) return true;
            if (emailVal.includes(userEmailClean) || userEmailClean.includes(emailVal)) return true;
          }
          if (idVal && idVal === userEmailClean) return true;
          if (nameVal && nameVal === userEmailClean) return true;
          return false;
        });

        if (matchedHod) {
          sourceTabName = tabName;
          break;
        }
      }
    }

    if (!matchedHod) {
      // Fallback: Check if email matches any employee in Employees tab
      const empSheet = getSheet('Employees');
      if (empSheet) {
        const empRows = sheetToJson(empSheet);
        const matchedEmp = empRows.find(e => {
          const eEmail = getVal(e, emailAliases).toLowerCase().trim();
          const eId = getVal(e, idAliases).toLowerCase().trim();
          return (eEmail && eEmail === userEmailClean) || (eId && eId === userEmailClean);
        });
        if (matchedEmp) {
          const empName = getVal(matchedEmp, nameAliases).toLowerCase().trim();
          for (let tabName of targetSheets) {
            const sheet = getSheet(tabName);
            if (sheet) {
              const rows = sheetToJson(sheet);
              matchedHod = rows.find(h => {
                const hName = getVal(h, nameAliases).toLowerCase().trim();
                return hName && (hName === empName || hName.includes(empName) || empName.includes(hName));
              });
              if (matchedHod) {
                sourceTabName = tabName;
                break;
              }
            }
          }
        }
      }
    }

    if (!matchedHod) {
      return null; // Match failed: Not an authorized HOD/C-Suite/HOHR credential
    }

    realHodId = getVal(matchedHod, idAliases);
    realHodName = getVal(matchedHod, nameAliases);
    realCostCentre = getVal(matchedHod, ['Cost Centre', 'CostCentre', 'Department']);
    realPosition = getVal(matchedHod, ['Position Title', 'PositionTitle', 'Position', 'JobTitle']);

    // Attempt resolving Csuite and HOHR emails from dedicated tabs if missing
    const csProfile = getCSuiteEmailProfile();
    csuiteName = getVal(matchedHod, ['CsuiteName', 'CSuiteName']) || csProfile.Name;
    csuiteEmail = getVal(matchedHod, ['CsuiteEmail', 'CSuiteEmail']) || csProfile.Email;

    const hrProfile = getHOHREmailProfile();
    hohrName = getVal(matchedHod, ['HohrName', 'HOHRName']) || hrProfile.Name;
    hohrEmail = getVal(matchedHod, ['HohrEmail', 'HOHREmail']) || hrProfile.Email;

  } catch (e) {
    Logger.log('Error in resolveRealHODProfile: ' + e.message);
    return null;
  }

  // STEP 2: Secondary enrichment from "For IT" tab if any field is missing
  if (!realHodId || !realCostCentre) {
    try {
      const itSheet = getSheet('For IT');
      if (itSheet && realHodName) {
        const itRows = sheetToJson(itSheet);
        const cleanHodName = realHodName.toLowerCase().trim();

        const matchedIt = itRows.find(r => {
          const itHodName = getVal(r, ['HOD', 'HODName', 'HodName', 'Name']).toLowerCase().trim();
          return itHodName && (itHodName === cleanHodName || itHodName.includes(cleanHodName) || cleanHodName.includes(itHodName));
        });

        if (matchedIt) {
          if (!realHodId) realHodId = getVal(matchedIt, ['Employee No', 'EmployeeNo', 'EmployeeID', 'ID']);
          if (!realCostCentre) realCostCentre = getVal(matchedIt, ['Cost Centre', 'CostCentre', 'Department']);
          if (!realPosition) realPosition = getVal(matchedIt, ['Position Title', 'PositionTitle', 'Position']);
        }
      }
    } catch (e) {
      Logger.log('Step 2 (For IT fallback) error: ' + e.message);
    }
  }

  // STEP 3: Tertiary enrichment from "Employees" tab
  if (!realHodId || !realPosition) {
    try {
      const empSheet = getSheet('Employees');
      if (empSheet) {
        const empRows = sheetToJson(empSheet);
        const cleanHodName = (realHodName || '').toLowerCase().trim();
        const cleanHodId = (realHodId || '').toLowerCase().trim();

        const matchedEmp = empRows.find(e => {
          const eId = getVal(e, ['Employee No', 'ID', 'EmployeeID', 'EmployeeNo']).toLowerCase().trim();
          const eName = getVal(e, ['Name', 'EmployeeName']).toLowerCase().trim();
          const eEmail = getVal(e, ['Email', 'EmailAddress']).toLowerCase().trim();

          if (userEmailClean && eEmail && eEmail === userEmailClean) return true;
          if (cleanHodId && eId && eId === cleanHodId) return true;
          if (cleanHodName && eName && (eName === cleanHodName || eName.includes(cleanHodName))) return true;
          return false;
        });

        if (matchedEmp) {
          if (!realHodId) realHodId = getVal(matchedEmp, ['Employee No', 'ID', 'EmployeeID', 'EmployeeNo']);
          if (!realHodName) realHodName = getVal(matchedEmp, ['Name', 'EmployeeName']);
          if (!realCostCentre) realCostCentre = getVal(matchedEmp, ['Cost Centre', 'Department', 'CostCentre']);
          if (!realPosition) realPosition = getVal(matchedEmp, ['Position Title', 'Position', 'JobTitle']);
        }
      }
    } catch (e) {
      Logger.log('Step 3 (Employees enrichment) error: ' + e.message);
    }
  }

  // STEP 4: Determine Role Flags & Role Title (HOD, C-Suite, HOHR)
  const tabLower = sourceTabName.toLowerCase();
  const posLower = realPosition.toLowerCase();

  let isCsuite = tabLower.includes('csuite') || tabLower.includes('c-suite') || 
                 posLower.includes('c-suite') || posLower.includes('csuite') || 
                 posLower.includes('ceo') || posLower.includes('chief') || posLower.includes('managing director');
  
  let isHohr = tabLower.includes('hohr') || 
               posLower.includes('hohr') || posLower.includes('head of hr') || 
               posLower.includes('hr head');

  let isHod = tabLower.includes('hod email') || (!isCsuite && !isHohr) || posLower.includes('hod') || posLower.includes('head of department') || posLower.includes('manager');

  // Also check if user email matches C-Suite or HOHR profiles in directory
  if (csuiteEmail && csuiteEmail.toLowerCase().trim() === userEmailClean) isCsuite = true;
  if (hohrEmail && hohrEmail.toLowerCase().trim() === userEmailClean) isHohr = true;

  let roleTitle = 'Head of Department';
  if (isCsuite && isHohr && isHod) roleTitle = 'HOD / C-Suite / HOHR';
  else if (isCsuite && isHohr) roleTitle = 'C-Suite / Head of HR';
  else if (isCsuite && isHod) roleTitle = 'HOD / C-Suite Executive';
  else if (isHohr && isHod) roleTitle = 'HOD / Head of HR';
  else if (isCsuite) roleTitle = 'C-Suite Executive';
  else if (isHohr) roleTitle = 'Head of HR';
  else if (tabLower.includes('hr email')) roleTitle = 'HR Admin';

  if (!realPosition || realPosition === 'Head of Department') {
    if (isCsuite && !isHod) realPosition = 'C-Suite Executive';
    else if (isHohr && !isHod) realPosition = 'Head of HR';
    else if (tabLower.includes('hr email')) realPosition = 'HR Executive';
    else if (!realPosition) realPosition = 'Head of Department';
  }

  return {
    ID: realHodId || 'N/A',
    Name: realHodName || '',
    CostCentre: realCostCentre || 'N/A',
    Position: realPosition || '',
    RoleTitle: roleTitle,
    isHod: isHod,
    isCsuite: isCsuite,
    isHohr: isHohr,
    Email: userEmailClean || '',
    CsuiteName: csuiteName,
    CsuiteEmail: csuiteEmail,
    HohrName: hohrName,
    HohrEmail: hohrEmail
  };

}


/**
 * Resolves the specific C-Suite Executive assigned to a requester according to their row in "For IT" tab.
 * Table Header in "For IT": Employee No | Name | Cost Centre | Position Title | HOD | Csuite | HOHR | Email
 */
function resolveCSuiteProfileForRequester(requesterCostCentre, requesterIdOrName) {
  let cleanDept = String(requesterCostCentre || '').toLowerCase().trim();
  let cleanReqId = String(requesterIdOrName || '').toLowerCase().trim();
  
  const emailAliases = ['Email', 'EmailAddress', 'HODEmail', 'HOD Email', 'Email Address'];
  const nameAliases = ['HOD', 'HODName', 'HodName', 'Name', 'HOD Name', 'Csuite', 'CsuiteName', 'C-Suite'];
  const csAliases = ['Csuite', 'C-Suite', 'C Suite', 'CsuiteName', 'CSuiteName', 'CSuite', 'CsuiteExecutive'];
  const csEmailAliases = ['CsuiteEmail', 'CSuiteEmail', 'Csuite Email', 'C-Suite Email', 'C Suite Email'];
  const deptAliases = ['Cost Centre', 'CostCentre', 'Department', 'Dept'];
  const idAliases = ['Employee No', 'EmployeeNo', 'EmployeeID', 'ID'];

  const getVal = (rowObj, nameList) => {
    if (!rowObj) return '';
    const keys = Object.keys(rowObj);
    for (let n of nameList) {
      const matchKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === n.toLowerCase().replace(/[^a-z0-9]/g, ''));
      if (matchKey && rowObj[matchKey] !== undefined && rowObj[matchKey] !== null) {
        return String(rowObj[matchKey]).trim();
      }
    }
    return '';
  };

  let targetCsuiteName = '';
  let targetCsuiteEmail = '';

  // STEP 1: Refer to the "For IT" tab according to the employee's row
  try {
    const itSheet = getSheet('For IT');
    if (itSheet && cleanReqId) {
      const itRows = sheetToJson(itSheet);
      const empRow = itRows.find(r => {
        const rId = getVal(r, idAliases).toLowerCase().trim();
        const rName = getVal(r, ['Name', 'EmployeeName']).toLowerCase().trim();
        const rEmail = getVal(r, emailAliases).toLowerCase().trim();
        return (rId && rId === cleanReqId) || (rName && (rName === cleanReqId || rName.includes(cleanReqId))) || (rEmail && rEmail === cleanReqId);
      });

      if (empRow) {
        targetCsuiteName = getVal(empRow, csAliases);
        targetCsuiteEmail = getVal(empRow, csEmailAliases);
        if (!cleanDept) cleanDept = getVal(empRow, deptAliases).toLowerCase().trim();
      }
    }
  } catch (e) {
    Logger.log('resolveCSuiteProfileForRequester IT tab lookup error: ' + e.message);
  }

  // STEP 2: Match resolved C-Suite name or department against "Csuite email" directory tab
  try {
    const csSheet = getSheet('Csuite email') || getSheet('C-Suite email');
    if (csSheet) {
      const csRows = sheetToJson(csSheet);
      if (csRows.length > 0) {
        let matchedCs = null;

        // Match by C-Suite name / email resolved from "For IT" tab
        if (targetCsuiteName || targetCsuiteEmail) {
          const cleanCsName = targetCsuiteName.toLowerCase().trim();
          const cleanCsEmail = targetCsuiteEmail.toLowerCase().trim();
          matchedCs = csRows.find(r => {
            const rowEmail = getVal(r, emailAliases).toLowerCase().trim();
            const rowName = getVal(r, nameAliases).toLowerCase().trim();
            return (cleanCsEmail && rowEmail && rowEmail === cleanCsEmail) ||
                   (cleanCsName && rowName && (rowName === cleanCsName || rowName.includes(cleanCsName) || cleanCsName.includes(rowName)));
          });
        }

        // Match by Department / Cost Centre if name match failed
        if (!matchedCs && cleanDept) {
          matchedCs = csRows.find(r => {
            const rowDept = getVal(r, deptAliases).toLowerCase().trim();
            return rowDept && (rowDept === cleanDept || rowDept.includes(cleanDept) || cleanDept.includes(rowDept));
          });
        }

        // Fallback to first row
        if (!matchedCs) matchedCs = csRows[0];

        return {
          ID: getVal(matchedCs, idAliases) || 'CSUITE',
          Name: getVal(matchedCs, nameAliases) || targetCsuiteName || 'C-Suite Executive',
          CostCentre: getVal(matchedCs, deptAliases) || requesterCostCentre,
          Position: getVal(matchedCs, ['Position Title', 'PositionTitle', 'Position']) || 'C-Suite Executive',
          Email: getVal(matchedCs, emailAliases) || targetCsuiteEmail
        };
      }
    }
  } catch(e) {
    Logger.log('resolveCSuiteProfileForRequester directory lookup error: ' + e.message);
  }

  return {
    ID: 'CSUITE',
    Name: targetCsuiteName || 'C-Suite Executive',
    CostCentre: requesterCostCentre || '',
    Position: 'C-Suite Executive',
    Email: targetCsuiteEmail || ''
  };
}


/**
 * Reads C-Suite profile from "Csuite email" tab (default fallback)
 */
function getCSuiteEmailProfile(costCentre) {
  return resolveCSuiteProfileForRequester(costCentre);
}


/**
 * Reads HOHR profile from "HOHR email" tab
 */
function getHOHREmailProfile() {
  const emailAliases = ['Email', 'EmailAddress', 'HODEmail', 'HOD Email', 'Email Address'];
  const nameAliases = ['HOD', 'HODName', 'HodName', 'Name', 'HOD Name'];

  const getVal = (rowObj, nameList) => {
    if (!rowObj) return '';
    const keys = Object.keys(rowObj);
    for (let n of nameList) {
      const matchKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === n.toLowerCase().replace(/[^a-z0-9]/g, ''));
      if (matchKey && rowObj[matchKey] !== undefined && rowObj[matchKey] !== null) {
        return String(rowObj[matchKey]).trim();
      }
    }
    return '';
  };

  try {
    // HOHR email tab is Head of HR; HR email tab is HR Admin
    const hrSheet = getSheet('HOHR email') || getSheet('HOHR Email') || getSheet('HOHR') || getSheet('HR email') || getSheet('HR Email');
    if (hrSheet) {
      const rows = sheetToJson(hrSheet);
      if (rows.length > 0) {
        const row = rows[0];
        return {
          ID: getVal(row, ['Employee No', 'EmployeeNo', 'EmployeeID', 'ID']),
          Name: getVal(row, nameAliases) || 'Head of HR',
          CostCentre: getVal(row, ['Cost Centre', 'CostCentre', 'Department']),
          Position: getVal(row, ['Position Title', 'PositionTitle', 'Position']),
          Email: getVal(row, emailAliases)
        };
      }
    }
  } catch(e) {
    Logger.log('getHOHREmailProfile error: ' + e.message);
  }
  return { ID: '', Name: 'Head of HR', CostCentre: '', Position: 'Head of HR', Email: 'hohr@apollofood.com.my' };
}

/**
 * Returns all real HOD, C-Suite, and HOHR profiles from directory tabs
 */
function getAllRealHODProfiles() {
  const hods = [];
  const getVal = (rowObj, nameList) => {
    if (!rowObj) return '';
    const keys = Object.keys(rowObj);
    for (let n of nameList) {
      const matchKey = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === n.toLowerCase().replace(/[^a-z0-9]/g, ''));
      if (matchKey && rowObj[matchKey] !== undefined && rowObj[matchKey] !== null) {
        return String(rowObj[matchKey]).trim();
      }
    }
    return '';
  };

  const targetTabs = ['HOD email', 'Csuite email', 'C-Suite email', 'HOHR email', 'HOHR Email', 'HR email'];

  try {
    targetTabs.forEach(tabName => {
      const sheet = getSheet(tabName);
      if (sheet) {
        const rows = sheetToJson(sheet);
        const tabLower = tabName.toLowerCase();
        let defaultRole = 'HOD';
        let defaultPos = 'Head of Department';
        if (tabLower.includes('csuite') || tabLower.includes('c-suite')) {
          defaultRole = 'C-Suite';
          defaultPos = 'C-Suite Executive';
        } else if (tabLower.includes('hohr')) {
          defaultRole = 'HOHR';
          defaultPos = 'Head of HR';
        } else if (tabLower.includes('hr')) {
          defaultRole = 'HR Admin';
          defaultPos = 'HR Executive';
        }

        rows.forEach(h => {
          const id = getVal(h, ['Employee No', 'EmployeeNo', 'EmployeeID', 'ID']);
          const name = getVal(h, ['HOD', 'HODName', 'HodName', 'Name']);
          const dept = getVal(h, ['Cost Centre', 'CostCentre', 'Department']);
          const pos = getVal(h, ['Position Title', 'PositionTitle', 'Position', 'JobTitle']);
          const email = getVal(h, ['Email', 'EmailAddress', 'HODEmail']);

          if (name && !hods.some(x => x.Name.toLowerCase() === name.toLowerCase() && x.Email.toLowerCase() === email.toLowerCase())) {
            hods.push({
              ID: id || 'N/A',
              Name: name,
              CostCentre: dept || 'Cost Centre',
              Position: pos || defaultPos,
              RoleTag: defaultRole,
              Email: email || '',
              CsuiteName: getVal(h, ['CsuiteName', 'CSuiteName']),
              CsuiteEmail: getVal(h, ['CsuiteEmail', 'CSuiteEmail']),
              HohrName: getVal(h, ['HohrName', 'HOHRName']),
              HohrEmail: getVal(h, ['HohrEmail', 'HOHREmail'])
            });
          }
        });
      }
    });
  } catch (e) {
    Logger.log('getAllRealHODProfiles error: ' + e.message);
  }

  return hods;
}




function getActiveHODProfile(providedEmail, requesterIdOrName) {
  let userEmail = String(providedEmail || '').trim();
  
  if (!userEmail) {
    try {
      userEmail = Session.getActiveUser().getEmail() || '';
    } catch (e) {
      userEmail = '';
    }
  }

  if (!userEmail) {
    return {
      valid: false,
      email: '',
      message: 'No email address detected. Please sign in with your registered HOD email.'
    };
  }

  const allowedDomainStr = getConfigProperty('ALLOWED_DOMAIN', '').toLowerCase().trim();
  const allowedDomains = allowedDomainStr
    ? allowedDomainStr.split(',').map(d => d.trim().replace(/^@/, '')).filter(Boolean)
    : [];

  if (allowedDomains.length > 0 && userEmail && userEmail.includes('@')) {
    const domain = (userEmail.split('@')[1] || '').toLowerCase().trim();
    if (!allowedDomains.includes(domain)) {
      const domainLabels = allowedDomains.map(d => '@' + d).join(' or ');
      return {
        valid: false,
        email: userEmail,
        message: `Access denied. Only company email accounts (${domainLabels}) can access the HOD Portal.`
      };
    }
  }

  // Resolve real HOD profile strictly via HOD email schema
  const realHod = resolveRealHODProfile(userEmail, requesterIdOrName);

  if (!realHod) {
    return {
      valid: false,
      email: userEmail,
      hod: null,
      message: `Access Denied: The email address (${userEmail}) is not registered as an authorized approver in the system directory.`
    };
  }


  return {
    valid: true,
    email: userEmail,
    hod: realHod
  };
}



function validateHODAccess(userEmail, requesterIdOrName) {
  const profile = getActiveHODProfile(userEmail, requesterIdOrName);
  if (!profile.valid) {
    return profile;
  }
  return { valid: true, email: profile.email, hod: profile.hod };
}


/**
 * Server-side identity resolution for authorization.
 * STRICT SECURITY REQUIREMENT: Never trust client-supplied email/ID parameters.
 * Uses Session.getActiveUser().getEmail() as authoritative source of identity.
 */
function resolveAuthenticatedUserServerSide() {
  let sessionEmail = '';
  try {
    sessionEmail = Session.getActiveUser().getEmail();
  } catch (e) {}

  let targetEmail = String(sessionEmail || '').trim();

  if (!targetEmail) {
    return { valid: false, email: '', hod: null, message: 'No authenticated user identity detected. Please sign in with your authorized company account.' };
  }

  const auth = validateHODAccess(targetEmail);
  if (!auth.valid || !auth.hod) {
    return { valid: false, email: targetEmail, hod: null, message: auth.message || `Access Denied: (${targetEmail}) is not an authorized approver.` };
  }

  return { valid: true, email: targetEmail, hod: auth.hod };
}

/**
 * Compares two approver identities for exact matching.
 * Primary key: Unique Employee No / ID.
 * Secondary keys: Email, Name.
 */
function isSameApprover(approverA, approverB) {
  if (!approverA || !approverB) return false;

  const idA = String(approverA.ID || approverA.EmployeeNo || approverA.EmployeeID || approverA['Employee No'] || '').toLowerCase().trim();
  const idB = String(approverB.ID || approverB.EmployeeNo || approverB.EmployeeID || approverB['Employee No'] || '').toLowerCase().trim();

  const isInvalid = (val) => !val || val === 'n/a' || val === 'pending' || val === 'none' || val === 'approved' || val === 'csuite' || val === 'hohr' || val === 'hod';

  if (!isInvalid(idA) && !isInvalid(idB)) {
    return idA === idB;
  }

  const emailA = String(approverA.Email || approverA.EmailAddress || '').toLowerCase().trim();
  const emailB = String(approverB.Email || approverB.EmailAddress || '').toLowerCase().trim();

  if (emailA && emailB && !emailA.includes('pending') && !emailB.includes('pending')) {
    return emailA === emailB;
  }

  const nameA = String(approverA.Name || approverA.HOD || approverA.Csuite || approverA.HOHR || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const nameB = String(approverB.Name || approverB.HOD || approverB.Csuite || approverB.HOHR || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  if (nameA && nameB && !isInvalid(nameA) && !isInvalid(nameB)) {
    return nameA === nameB || nameA.includes(nameB) || nameB.includes(nameA);
  }

  return false;
}

/**
 * Evaluates the CURRENT ACTIVE APPROVAL STAGE for a training requisition.
 * Sequential workflow: HOD -> CSuite -> HOHR -> NONE
 * Returns: 'HOD' | 'CSuite' | 'HOHR' | 'NONE'
 */
function getCurrentActiveApprovalStage(requisition) {
  if (!requisition) return 'NONE';

  const appStatus = String(requisition.ApprovalStatus || '').trim().toLowerCase();

  // If request is overall complete, approved, rejected, returned, cancelled, or closed -> NONE
  if (appStatus === 'approved' || appStatus === 'auto-approved' || appStatus === 'fully approved' || appStatus.includes('reject') || appStatus === 'returned' || appStatus.includes('cancel') || appStatus.includes('closed') || appStatus.includes('completed')) {
    return 'NONE';
  }

  const hodSt = String(requisition.HODStatus || '').trim().toLowerCase();
  const csSt = String(requisition.CsuiteStatus || '').trim().toLowerCase();
  const hrSt = String(requisition.HOHRStatus || '').trim().toLowerCase();

  // STAGE 1: HOD Stage
  const isHodCompleted = hodSt === 'approved' || hodSt === 'n/a' || appStatus.includes('c-suite') || appStatus.includes('csuite') || appStatus.includes('hohr');
  
  if (!isHodCompleted) {
    if (hodSt === 'pending' || appStatus.includes('pending hod approval') || appStatus === 'pending' || appStatus === 'submitted' || appStatus === 'draft' || !hodSt) {
      if (hodSt !== 'rejected' && hodSt !== 'returned') {
        return 'HOD';
      }
    }
    if (appStatus.includes('pending hod approval') || appStatus === 'pending') {
      return 'HOD';
    }
  }

  // STAGE 2: CSuite Stage
  const isCsCompleted = csSt === 'approved' || csSt === 'n/a' || appStatus.includes('hohr');

  if (isHodCompleted && !isCsCompleted) {
    if (csSt === 'pending' || appStatus.includes('pending c-suite approval') || appStatus.includes('pending csuite approval')) {
      if (csSt !== 'rejected' && csSt !== 'returned') {
        return 'CSuite';
      }
    }
    if (appStatus.includes('c-suite') || appStatus.includes('csuite')) {
      return 'CSuite';
    }
  }

  // STAGE 3: HOHR Stage
  const isHrCompleted = hrSt === 'approved' || hrSt === 'n/a';

  if (isHodCompleted && isCsCompleted && !isHrCompleted) {
    if (hrSt === 'pending' || appStatus.includes('pending hohr approval') || appStatus.includes('pending hr approval')) {
      if (hrSt !== 'rejected' && hrSt !== 'returned') {
        return 'HOHR';
      }
    }
    if (appStatus.includes('hohr') || appStatus.includes('hr')) {
      return 'HOHR';
    }
  }

  return 'NONE';
}

/**
 * Resolves the specific assigned approver profiles (HOD, CSuite, HOHR) for a given training requisition.
 * Checks stored requisition fields and looks up the requester in the "For IT" (Approval Assignment) table.
 */
function getAssignedApproversForRequisition(t) {
  if (!t) return { HOD: null, CSuite: null, HOHR: null };

  let rawHod = String(t.HODApprover || t.HOD || '').trim();
  let rawCs = String(t.CsuiteApprover || t.Csuite || t.CSuite || '').trim();
  let rawHohr = String(t.HOHRApprover || t.HOHR || '').trim();

  const isStatusValue = (val) => {
    const v = String(val || '').toLowerCase().trim();
    return !v || v === 'pending' || v === 'approved' || v === 'rejected' || v === 'returned' || v === 'n/a';
  };

  const reqId = String(t.RequestedBy || t.EmployeeID || '').trim().toLowerCase();
  const reqName = String(t.RequestedByName || '').trim().toLowerCase();
  const reqEmail = String(t.RequestedByEmail || '').trim().toLowerCase();

  let itRow = null;
  try {
    const itSheet = getSheet('For IT');
    if (itSheet) {
      const itRows = sheetToJson(itSheet);
      itRow = itRows.find(r => {
        const empId = String(r['Employee No'] || r.EmployeeNo || r.ID || r.EmployeeID || '').trim().toLowerCase();
        const empName = String(r.Name || r.EmployeeName || '').trim().toLowerCase();
        const empEmail = String(r.Email || r.EmailAddress || '').trim().toLowerCase();

        return (reqId && empId && empId === reqId) || (reqName && empName && (empName === reqName || empName.includes(reqName))) || (reqEmail && empEmail && empEmail === reqEmail);
      });
    }
  } catch(e) {}

  if (itRow) {
    if (isStatusValue(rawHod)) {
      rawHod = String(itRow.HOD || itRow.HODName || itRow.HodName || itRow.Manager || itRow.ReportTo || '').trim();
    }
    if (isStatusValue(rawCs)) {
      rawCs = String(itRow.Csuite || itRow.CsuiteName || itRow.CSuiteName || '').trim();
    }
    if (isStatusValue(rawHohr)) {
      rawHohr = String(itRow.HOHR || itRow.HohrName || itRow.HOHRName || '').trim();
    }
  }

  // Resolve rawHod to full profile
  let hodProfile = null;
  if (rawHod && !isStatusValue(rawHod)) {
    hodProfile = resolveRealHODProfile(rawHod) || { Name: rawHod, ID: rawHod, Email: '' };
  } else if (itRow && itRow.HOD) {
    hodProfile = resolveRealHODProfile(itRow.HOD) || { Name: itRow.HOD, ID: itRow.HOD, Email: '' };
  }

  // Resolve rawCs to full profile
  let csProfile = null;
  if (rawCs && !isStatusValue(rawCs)) {
    csProfile = resolveRealHODProfile(rawCs) || resolveCSuiteProfileForRequester(t.Department, reqId) || { Name: rawCs, ID: 'CSUITE', Email: '' };
  } else {
    csProfile = resolveCSuiteProfileForRequester(t.Department, reqId);
  }

  // Resolve rawHohr to full profile
  let hrProfile = null;
  if (rawHohr && !isStatusValue(rawHohr)) {
    hrProfile = resolveRealHODProfile(rawHohr) || getHOHREmailProfile() || { Name: rawHohr, ID: 'HOHR', Email: '' };
  } else {
    hrProfile = getHOHREmailProfile();
  }

  return {
    HOD: hodProfile,
    CSuite: csProfile,
    HOHR: hrProfile
  };
}

/**
 * Returns the assigned approver profile for a specific stage of a requisition.
 */
function getAssignedApproverForStage(requisition, stage) {
  const approvers = getAssignedApproversForRequisition(requisition);
  if (stage === 'HOD') return approvers.HOD;
  if (stage === 'CSuite') return approvers.CSuite;
  if (stage === 'HOHR') return approvers.HOHR;
  return null;
}


