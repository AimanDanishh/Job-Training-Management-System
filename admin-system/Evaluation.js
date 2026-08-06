/**
 * Evaluation.gs — Training Evaluation and Post-Training (6-month) Evaluation
 */

// ─── Training Evaluation ────────────────────────────────────────────────────────
function getTrainingEvaluations(trainingId) {
  try {
    if (!trainingId) return ok([]);
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return ok([]);
    const sheet = ss.getSheetByName('TrainingEval');
    if (!sheet) return ok([]);
    const rows  = sheetToJson(sheet);
    return ok(rows);
  } catch (e) {
    return err(e.message);
  }
}

function saveTrainingEvaluation(data) {
  try {
    if (!data.TrainingID || !data.EmployeeID)
      return err('TrainingID and EmployeeID are required.');

    const ss = getTrainingDataSpreadsheet(data.TrainingID);
    if (!ss) return err('Could not open per-training sheet for ID: ' + data.TrainingID);

    let sheet = ss.getSheetByName('TrainingEval');
    if (!sheet) {
      sheet = ss.insertSheet('TrainingEval');
      sheet.appendRow(['ID', 'TrainingID', 'EmployeeID', 'EmployeeName', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'SectionB1', 'SectionB2', 'SectionB3', 'AvgScore', 'SubmittedAt']);
      sheet.getRange('A1:P1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // Prevent duplicate submission
    const rows = sheetToJson(sheet);
    const exists = rows.find(r =>
      String(r.TrainingID || '').trim() === String(data.TrainingID).trim() &&
      String(r.EmployeeID || '').trim().toLowerCase() === String(data.EmployeeID).trim().toLowerCase()
    );
    if (exists) return err('You have already submitted an evaluation for this training.');

    const scores = [data.Q1, data.Q2, data.Q3, data.Q4, data.Q5, data.Q6, data.Q7].map(Number);
    if (scores.some(s => isNaN(s) || s < 1 || s > 5)) {
      return err('All 7 evaluation questions (scale 1-5) are required.');
    }

    const avg = (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2);

    const evalRow = [
      generateId('EVL'),
      data.TrainingID,
      data.EmployeeID,
      data.EmployeeName  || '',
      data.Q1, data.Q2, data.Q3,
      data.Q4, data.Q5, data.Q6, data.Q7,
      data.SectionB1     || '',
      data.SectionB2     || '',
      data.SectionB3     || '',
      avg,
      now()
    ];
    sheet.appendRow(evalRow);

    return ok({ message: 'Training evaluation submitted. Average score: ' + avg });
  } catch (e) {
    return err('Failed to save evaluation: ' + e.message);
  }
}

// ─── Post-Training Evaluation (6-month) ────────────────────────────────────────
function getPostEvaluations(trainingId) {
  try {
    if (trainingId) {
      const ss = getTrainingDataSpreadsheet(trainingId);
      if (!ss) return ok([]);
      const sheet = ss.getSheetByName('PostEval');
      if (!sheet) return ok([]);
      return ok(sheetToJson(sheet));
    }

    // Iterate across all trainings if trainingId omitted
    const tSheet = getSheet(SHEET_NAMES.trainings);
    if (!tSheet) return ok([]);
    const trainings = sheetToJson(tSheet);
    let allPosts = [];
    trainings.forEach(t => {
      if (t.ID) {
        const ss = getTrainingDataSpreadsheet(t.ID);
        if (ss) {
          const sheet = ss.getSheetByName('PostEval');
          if (sheet) allPosts = allPosts.concat(sheetToJson(sheet));
        }
      }
    });
    return ok(allPosts);
  } catch (e) {
    return err(e.message);
  }
}

function savePostEvaluation(data) {
  try {
    if (!data.TrainingID || !data.EmployeeID || !data.EvaluatorName)
      return err('TrainingID, EmployeeID, and Evaluator Name are required.');

    const cb = Number(data.CompetencyBefore);
    const ca = Number(data.CompetencyAfter);
    if (isNaN(cb) || cb < 1 || cb > 5 || isNaN(ca) || ca < 1 || ca > 5) {
      return err('Competency levels before and after training (scale 1-5) are required.');
    }

    const ss = getTrainingDataSpreadsheet(data.TrainingID);
    if (!ss) return err('Could not open per-training sheet for ID: ' + data.TrainingID);

    let sheet = ss.getSheetByName('PostEval');
    if (!sheet) {
      sheet = ss.insertSheet('PostEval');
      sheet.appendRow(['ID', 'TrainingID', 'EmployeeID', 'EvaluatorName', 'EvaluatorID', 'CompetencyBefore', 'CompetencyAfter', 'Improvement', 'CanApply', 'FurtherTraining', 'Comments', 'SubmittedAt']);
      sheet.getRange('A1:L1').setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const rows = sheetToJson(sheet);
    const exists = rows.find(r =>
      String(r.TrainingID || '').trim() === String(data.TrainingID).trim() &&
      String(r.EmployeeID || '').trim().toLowerCase() === String(data.EmployeeID).trim().toLowerCase()
    );
    if (exists) return err('A post-training evaluation has already been submitted for this employee.');

    const postRow = [
      generateId('PEV'),
      data.TrainingID,
      data.EmployeeID,
      data.EvaluatorName      || '',
      data.EvaluatorID        || '',
      cb,
      ca,
      data.Improvement        || '',
      data.CanApply           || '',
      data.FurtherTraining    || '',
      data.Comments           || '',
      now()
    ];
    sheet.appendRow(postRow);

    // Auto-advance training stage if all post-evals are done
    tryAdvanceToEvaluationCompleted(data.TrainingID);

    return ok({ message: 'Post-training evaluation submitted successfully.' });
  } catch (e) {
    return err('Failed to save post-evaluation: ' + e.message);
  }
}

// ─── Evaluation Summary ─────────────────────────────────────────────────────────
function getEvaluationSummary(trainingId) {
  try {
    if (!trainingId) return ok({ evalCompleted: 0, avgScore: null, postCompleted: 0 });
    const ss = getTrainingDataSpreadsheet(trainingId);
    if (!ss) return ok({ evalCompleted: 0, avgScore: null, postCompleted: 0 });

    const evalSheet = ss.getSheetByName('TrainingEval');
    const evalRows  = evalSheet ? sheetToJson(evalSheet) : [];

    const postSheet = ss.getSheetByName('PostEval');
    const postRows  = postSheet ? sheetToJson(postSheet) : [];

    const scores = evalRows.map(r => Number(r.AvgScore)).filter(n => n > 0);
    const avgScore = scores.length > 0
      ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2)
      : null;

    return ok({
      evalCompleted:  evalRows.length,
      avgScore:       avgScore,
      postCompleted:  postRows.length,
    });
  } catch (e) {
    return err(e.message);
  }
}

// ─── Internal: auto-advance stage ──────────────────────────────────────────────
function tryAdvanceToEvaluationCompleted(trainingId) {
  try {
    const tSheet = getSheet(SHEET_NAMES.trainings);
    const tRows  = sheetToJson(tSheet);
    const t = tRows.find(r => r.ID === trainingId);
    if (!t) return;

    const ss = getTrainingDataSpreadsheet(trainingId);
    const eSheet = ss ? ss.getSheetByName('TrainingEval') : null;
    const evalCount = eSheet ? sheetToJson(eSheet).length : 0;

    if (evalCount >= Number(t.Participants) && t.Stage === 'Training Completed') {
      updateTrainingStage(trainingId, 'Evaluation Completed');
    }
  } catch (e) {
    Logger.log('Stage auto-advance error: ' + e.message);
  }
}

