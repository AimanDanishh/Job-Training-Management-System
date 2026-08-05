/**
 * Evaluation.gs — Training Evaluation & 6-Month Post Evaluation Service for Public Portal
 */

/**
 * Fetch basic public details of a training for display on evaluation forms
 */
function getTrainingInfo(trainingId) {
  try {
    if (!trainingId || String(trainingId).trim() === '') {
      return err('Training ID is required.');
    }
    const cleanId = String(trainingId).trim();
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return err('Trainings sheet unavailable.');

    const rows = sheetToJson(tSheet);
    const t = rows.find(r => String(r.ID || r.TrainingID || '').trim() === cleanId);

    if (!t) return err('Training programme not found.');

    return ok({
      ID: t.ID,
      Code: t.Code || '',
      Name: t.Name || '',
      Category: t.Category || '',
      Trainer: t.Trainer || '',
      StartDate: t.StartDate || '',
      EndDate: t.EndDate || ''
    });
  } catch (e) {
    Logger.log('getTrainingInfo error: ' + e.message);
    return err('Failed to load training details: ' + e.message);
  }
}

/**
 * Submit public Training Evaluation
 */
function saveTrainingEvaluation(data) {
  try {
    if (!data || typeof data !== 'object') {
      return err('Invalid submission payload.');
    }

    const trainingId = data.TrainingID || data.trainingId;
    const employeeId = data.EmployeeID || data.employeeId;

    // 1. Validate on Server Side
    const validation = validatePublicEvaluation(trainingId, employeeId);
    if (!validation.valid) {
      return err(validation.message);
    }

    const emp      = validation.employee;
    const training = validation.training;
    const empName  = (emp && emp.Name) ? emp.Name : (data.EmployeeName || employeeId);

    const scores = [data.Q1, data.Q2, data.Q3, data.Q4, data.Q5, data.Q6, data.Q7].map(Number);
    if (scores.some(s => isNaN(s) || s < 1 || s > 5)) {
      return err('All 7 evaluation questions (scale 1-5) are required.');
    }
    
    const avg = (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2);

    const sheet = getSheet(SHEET_NAMES.trainingEval);
    if (!sheet) return err('TrainingEval sheet not found.');

    // Columns: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt']
    const evalRow = [
      generateId('EVL'),
      trainingId,
      employeeId,
      empName,
      data.Q1, data.Q2, data.Q3,
      data.Q4, data.Q5, data.Q6, data.Q7,
      data.SectionB1 || '',
      data.SectionB2 || '',
      data.SectionB3 || '',
      avg,
      now()
    ];
    sheet.appendRow(evalRow);
    try { syncEvaluationToTrainingDriveSheet(trainingId, evalRow); } catch(e) {}

    return ok({ 
      message: 'Training evaluation submitted successfully!',
      avgScore: avg,
      trainingTitle: training.Name || ''
    });

  } catch (e) {
    Logger.log('saveTrainingEvaluation error: ' + e.message);
    return err('Failed to save evaluation: ' + e.message);
  }
}

/**
 * Submit public 6-Month Post Evaluation (Supervisor Assessment)
 */
function savePostEvaluation(data) {
  try {
    if (!data || typeof data !== 'object') {
      return err('Invalid post-evaluation payload.');
    }

    const trainingId   = data.TrainingID || data.trainingId;
    const employeeId   = data.EmployeeID || data.employeeId;
    const token        = data.Token || data.token;
    const evaluatorName = data.EvaluatorName || data.evaluatorName;

    if (!evaluatorName || String(evaluatorName).trim() === '') {
      return err('Evaluator (Supervisor) Name is required.');
    }

    const cb = Number(data.CompetencyBefore);
    const ca = Number(data.CompetencyAfter);
    if (isNaN(cb) || cb < 1 || cb > 5 || isNaN(ca) || ca < 1 || ca > 5) {
      return err('Competency levels before and after training (scale 1-5) are required.');
    }

    // 1. Validate on Server Side
    const validation = validatePublicPostEvaluation(trainingId, employeeId, token);
    if (!validation.valid) {
      return err(validation.message);
    }

    const finalTId   = validation.trainingId;
    const finalEmpId = validation.employeeId;

    const sheet = getSheet(SHEET_NAMES.postEval);
    if (!sheet) return err('PostEval sheet not found.');

    // Columns: ['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID', 'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply', 'FurtherTraining', 'Comments', 'SubmittedAt']
    const postRow = [
      generateId('PEV'),
      finalTId,
      finalEmpId,
      String(evaluatorName).trim(),
      data.EvaluatorID      || '',
      cb,
      ca,
      data.Improvement        || '',
      data.CanApply           || '',
      data.FurtherTraining    || '',
      data.Comments           || '',
      now()
    ];
    sheet.appendRow(postRow);
    try { syncPostEvalToTrainingDriveSheet(finalTId, postRow); } catch(e) {}

    return ok({ message: '6-Month Post-Training Evaluation submitted successfully!' });

  } catch (e) {
    Logger.log('savePostEvaluation error: ' + e.message);
    return err('Failed to save post-evaluation: ' + e.message);
  }
}


