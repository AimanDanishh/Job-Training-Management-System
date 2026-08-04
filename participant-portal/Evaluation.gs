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

    const scores = [data.Q1, data.Q2, data.Q3, data.Q4, data.Q5, data.Q6, data.Q7]
      .map(Number).filter(n => !isNaN(n) && n > 0);
    
    const avg = scores.length > 0
      ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2)
      : '0.00';

    const sheet = getSheet(SHEET_NAMES.trainingEval);
    if (!sheet) return err('TrainingEval sheet not found.');

    // Columns: ['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt']
    sheet.appendRow([
      generateId('EVL'),
      trainingId,
      employeeId,
      empName,
      data.Q1 || 0, data.Q2 || 0, data.Q3 || 0,
      data.Q4 || 0, data.Q5 || 0, data.Q6 || 0, data.Q7 || 0,
      data.SectionB1 || '',
      data.SectionB2 || '',
      data.SectionB3 || '',
      avg,
      now()
    ]);

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
    sheet.appendRow([
      generateId('PEV'),
      finalTId,
      finalEmpId,
      String(evaluatorName).trim(),
      data.EvaluatorID      || '',
      data.CompetencyBefore   || 0,
      data.CompetencyAfter    || 0,
      data.Improvement        || '',
      data.CanApply           || '',
      data.FurtherTraining    || '',
      data.Comments           || '',
      now()
    ]);

    return ok({ message: '6-Month Post-Training Evaluation submitted successfully!' });

  } catch (e) {
    Logger.log('savePostEvaluation error: ' + e.message);
    return err('Failed to save post-evaluation: ' + e.message);
  }
}
