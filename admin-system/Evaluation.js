/**
 * Evaluation.gs — Training Evaluation and Post-Training (6-month) Evaluation
 */

// ─── Training Evaluation ────────────────────────────────────────────────────────
function getTrainingEvaluations(trainingId) {
  try {
    const sheet = getSheet(SHEET_NAMES.trainingEval);
    const rows  = sheetToJson(sheet).filter(r => r.TrainingID === trainingId);
    return ok(rows);
  } catch (e) {
    return err(e.message);
  }
}

function saveTrainingEvaluation(data) {
  try {
    if (!data.TrainingID || !data.EmployeeID)
      return err('TrainingID and EmployeeID are required.');

    // Prevent duplicate submission
    const sheet = getSheet(SHEET_NAMES.trainingEval);
    const rows  = sheetToJson(sheet);
    const exists = rows.find(r =>
      r.TrainingID === data.TrainingID && r.EmployeeID === data.EmployeeID
    );
    if (exists) return err('You have already submitted an evaluation for this training.');

    const scores = [data.Q1, data.Q2, data.Q3, data.Q4, data.Q5, data.Q6, data.Q7].map(Number);
    if (scores.some(s => isNaN(s) || s < 1 || s > 5)) {
      return err('All 7 evaluation questions (scale 1-5) are required.');
    }

    const avg = (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2);

    sheet.appendRow([
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
    ]);
    return ok({ message: 'Training evaluation submitted. Average score: ' + avg });
  } catch (e) {
    return err('Failed to save evaluation: ' + e.message);
  }
}

// ─── Post-Training Evaluation (6-month) ────────────────────────────────────────
function getPostEvaluations(trainingId) {
  try {
    const sheet = getSheet(SHEET_NAMES.postEval);
    const rows  = sheetToJson(sheet).filter(r => r.TrainingID === trainingId);
    return ok(rows);
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

    const sheet = getSheet(SHEET_NAMES.postEval);
    const rows  = sheetToJson(sheet);
    const exists = rows.find(r =>
      r.TrainingID === data.TrainingID && r.EmployeeID === data.EmployeeID
    );
    if (exists) return err('A post-training evaluation has already been submitted for this employee.');

    sheet.appendRow([
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
    ]);

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
    const evalSheet = getSheet(SHEET_NAMES.trainingEval);
    const evalRows  = sheetToJson(evalSheet).filter(r => r.TrainingID === trainingId);

    const postSheet = getSheet(SHEET_NAMES.postEval);
    const postRows  = sheetToJson(postSheet).filter(r => r.TrainingID === trainingId);

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

    const eSheet = getSheet(SHEET_NAMES.trainingEval);
    const evalCount = sheetToJson(eSheet).filter(r => r.TrainingID === trainingId).length;

    if (evalCount >= Number(t.Participants) && t.Stage === 'Training Completed') {
      updateTrainingStage(trainingId, 'Evaluation Completed');
    }
  } catch (e) {
    Logger.log('Stage auto-advance error: ' + e.message);
  }
}

/**
 * Bulk email 6-month post-training evaluation form link to supervisor
 */
function sendSupervisorPostEvalEmails(trainingId, supervisorEmail, selectedEmpIds) {
  try {
    if (!trainingId || !supervisorEmail) {
      return err('Training ID and Supervisor Email are required.');
    }
    const cleanEmail = String(supervisorEmail).trim();
    const cleanBase = getPublicPortalUrl().split('?')[0];
    const postUrl = `${cleanBase}?page=post&id=${encodeURIComponent(trainingId)}`;

    const count = (Array.isArray(selectedEmpIds) && selectedEmpIds.length > 0) ? selectedEmpIds.length : 1;

    const subject = `[Action Required] 6-Month Post-Training Evaluation — TrainHub`;
    const body = `Dear Supervisor,\n\nYou have been requested to complete the 6-Month Post-Training Evaluation for your team member(s).\n\nPlease click the link below to open the evaluation form on the TrainHub Participant Portal:\n${postUrl}\n\nThank you,\nTrainHub Management System`;

    MailApp.sendEmail(cleanEmail, subject, body);
    return ok({ message: `Sent 6-month evaluation link to ${cleanEmail} for ${count} participant(s).` });
  } catch (e) {
    return err('Failed to send supervisor email: ' + e.message);
  }
}