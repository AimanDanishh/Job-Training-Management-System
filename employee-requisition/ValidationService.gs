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

/**
 * Validate employee requisition submission payload
 */
function validateEmployeeRequisitionData(data) {
  if (!data) return { valid: false, message: 'No requisition data provided.' };
  if (!data.TrainingName || String(data.TrainingName).trim() === '') {
    return { valid: false, message: 'Programme name / Title is required.' };
  }
  if (!data.StartDate || String(data.StartDate).trim() === '') {
    return { valid: false, message: 'Training Start Date is required.' };
  }
  if (!data.Venue || String(data.Venue).trim() === '') {
    return { valid: false, message: 'Training Venue is required.' };
  }

  // Mandatory Reason for Training (Reject empty / whitespace)
  const reason = String(data.Reason || data.Objectives || '').trim();
  if (!reason) {
    return { valid: false, message: 'Reason for training is required.' };
  }

  // Duration & 7 hours max per day
  const durationDays = Math.max(1, parseInt(data.Duration || 1, 10));
  const totalHours = parseFloat(data.TotalHours || 0);
  if (isNaN(totalHours) || totalHours <= 0) {
    return { valid: false, message: 'Duration (Hours) must be greater than 0.' };
  }
  if (totalHours > (durationDays * 7) || (totalHours / durationDays) > 7.001) {
    return { valid: false, message: 'Training duration cannot exceed 7 hours per day.' };
  }

  // Trainers validation
  const rawTrainers = Array.isArray(data.trainers) ? data.trainers : [];
  if (rawTrainers.length > 10) {
    return { valid: false, message: 'Maximum of 10 trainers allowed per training request.' };
  }
  const seenTrainers = new Set();
  for (let i = 0; i < rawTrainers.length; i++) {
    const t = String(rawTrainers[i] || '').trim();
    if (!t) {
      if (rawTrainers.length > 1) {
        return { valid: false, message: `Trainer #${i + 1} cannot be empty. Please enter trainer name or remove the empty field.` };
      }
      continue;
    }
    const key = t.toLowerCase();
    if (seenTrainers.has(key)) {
      return { valid: false, message: `Duplicate trainer name detected: "${t}".` };
    }
    seenTrainers.add(key);
  }

  // Attachments validation
  const allowedExtensions = ['.pdf', '.doc', '.docx', '.jpg', '.jpeg', '.png'];
  const maxFileSize = 10 * 1024 * 1024; // 10MB
  const files = Array.isArray(data.attachmentFiles) ? data.attachmentFiles : (data.BrochureFile && data.BrochureFile.data ? [data.BrochureFile] : []);
  const existingAtts = Array.isArray(data.existingAttachments) ? data.existingAttachments : (Array.isArray(data.attachments) ? data.attachments : []);

  if ((files.length + existingAtts.length) > 10) {
    return { valid: false, message: 'Maximum of 10 supporting documents allowed per request.' };
  }

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    if (!f || !f.name) continue;
    const ext = '.' + f.name.split('.').pop().toLowerCase();
    if (!allowedExtensions.includes(ext)) {
      return { valid: false, message: `File "${f.name}" has an unsupported format. Allowed types: ${allowedExtensions.join(', ')}` };
    }
    if (f.size && f.size > maxFileSize) {
      return { valid: false, message: `File "${f.name}" exceeds the maximum allowed size of 10MB.` };
    }
  }

  return { valid: true };
}
