/**
 * QRService.gs — QR Code and Attendance URL generation service using QuickChart API
 */

/**
 * Generate ONE Attendance URL for a specific Session.
 * URL Format: https://YOUR_WEBAPP_URL?page=attendance&session=SESSION_ID
 * 
 * @param {string} sessionId - The unique ID of the session (e.g. SES0001)
 * @returns {string} Fully qualified web app attendance URL
 */
function generateAttendanceURL(sessionId) {
  if (!sessionId) return '';
  let baseUrl = getPublicPortalUrl();
  if (!baseUrl) {
    baseUrl = 'https://script.google.com/macros/s/CURRENT_DEPLOYMENT_ID/exec';
  }
  const cleanBase = String(baseUrl || '').split('?')[0];
  return `${cleanBase}?page=attendance&session=${encodeURIComponent(sessionId)}`;
}

/**
 * Generate QR Code URL using QuickChart API with optional center company logo.
 * Converts Google Drive share links (e.g. https://drive.google.com/file/d/FILE_ID/view)
 * into direct public CDN image URLs (https://lh3.googleusercontent.com/d/FILE_ID).
 * Sets ecLevel=H (High Error Correction) so the QR code remains 100% scannable with logo.
 * 
 * @param {string} text - The URL or text to encode inside the QR code
 * @param {string} [customLogoUrl] - Optional Google Drive link or direct image URL for company logo
 * @returns {string} QuickChart QR image URL
 */
function generateQRCode(text, customLogoUrl) {
  if (!text) return '';
  const encodedText = encodeURIComponent(text);

  // Retrieve logo URL from parameter or Script Property (COMPANY_LOGO_URL)
  const rawLogo = customLogoUrl || getCompanyLogoUrl();
  const directLogoUrl = convertDriveLinkToDirectImageUrl(rawLogo);

  // ecLevel=H enables High Error Correction (30%) so center logo overlay doesn't affect scanning
  let qrUrl = `https://quickchart.io/qr?size=400&ecLevel=H&text=${encodedText}`;

  if (directLogoUrl) {
    qrUrl += `&centerImageUrl=${encodeURIComponent(directLogoUrl)}&centerImageSizeRatio=0.22`;
  }

  return qrUrl;
}

/**
 * Configure Company Logo URL in Script Properties and regenerate all session QR codes.
 * 
 * @param {string} logoUrlOrDriveLink - Google Drive share link or direct image URL
 * @returns {string} JSON response with status
 */
function updateCompanyLogoAndRegenerateQRs(logoUrlOrDriveLink) {
  try {
    setCompanyLogoUrl(logoUrlOrDriveLink || '');
    return generateMissingQRCodes(true); // Force regenerate all QRs
  } catch (e) {
    return err('Failed to update company logo: ' + e.message);
  }
}

/**
 * Utility function to scan the TrainingSessions sheet and automatically
 * generate missing Attendance URLs and QR Code URLs for all sessions.
 * 
 * @param {boolean} [forceRegenerate=false] - If true, re-generates all QR code URLs with updated logo settings
 * @returns {string} JSON response with status and count of updated sessions
 */
function generateMissingQRCodes(forceRegenerate = false) {
  try {
    const sheet = getSheet(SHEET_NAMES.trainingSessions);
    if (!sheet) return err('TrainingSessions sheet not found.');

    ensureTrainingSessionsSheetColumns(sheet);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return ok({ message: 'No sessions found.', count: 0 });

    const headers = data[0];

    const findIndex = (possibleNames) => {
      const cleans = possibleNames.map(n => n.toLowerCase().replace(/[^a-z0-9]/g, ''));
      return headers.findIndex(h => cleans.includes(String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '')));
    };

    const sessionIdx = findIndex(['SessionID', 'Session ID', 'Session_ID', 'ID']);
    const attUrlIdx  = findIndex(['AttendanceURL', 'Attendance URL', 'Attendance_URL', 'AttURL']);
    const qrUrlIdx   = findIndex(['QRCodeURL', 'QR Code URL', 'QR_Code_URL', 'QRCode', 'QRURL']);

    if (sessionIdx === -1 || attUrlIdx === -1 || qrUrlIdx === -1) {
      return err('Invalid header structure in TrainingSessions sheet.');
    }

    let updatedCount = 0;

    for (let i = 1; i < data.length; i++) {
      const sessionId = String(data[i][sessionIdx]).trim();
      if (!sessionId) continue;

      let attUrl = data[i][attUrlIdx];
      let qrUrl  = data[i][qrUrlIdx];

      let updated = false;
      if (!attUrl || String(attUrl).trim() === '') {
        attUrl = generateAttendanceURL(sessionId);
        sheet.getRange(i + 1, attUrlIdx + 1).setValue(attUrl);
        updated = true;
      }

      if (forceRegenerate || !qrUrl || String(qrUrl).trim() === '') {
        qrUrl = generateQRCode(attUrl || generateAttendanceURL(sessionId));
        sheet.getRange(i + 1, qrUrlIdx + 1).setValue(qrUrl);
        updated = true;
      }

      if (updated) updatedCount++;
    }

    return ok({
      message: `Successfully generated QR codes for ${updatedCount} session(s).`,
      count: updatedCount
    });
  } catch (e) {
    Logger.log('generateMissingQRCodes error: ' + e.message);
    return err('Failed to generate missing QR codes: ' + e.message);
  }
}

/**
 * Generate Evaluation Form URLs and QR Codes targeting Participant Portal
 * Level 1 Participant Training Evaluation: page=evaluation&id=TRN-xxx
 * Level 3 6-Month Supervisor Evaluation:  page=post&id=TRN-xxx
 */
function getEvaluationQRCodes(trainingId) {
  try {
    if (!trainingId) return err('Training ID is required.');
    let baseUrl = getPublicPortalUrl();
    if (!baseUrl) {
      baseUrl = 'https://script.google.com/macros/s/CURRENT_DEPLOYMENT_ID/exec';
    }
    const cleanBase = String(baseUrl || '').split('?')[0];

    const evalUrl = `${cleanBase}?page=evaluation&id=${encodeURIComponent(trainingId)}`;
    const postUrl = `${cleanBase}?page=post&id=${encodeURIComponent(trainingId)}`;

    return ok({
      participantEval: {
        url: evalUrl,
        qrCodeUrl: generateQRCode(evalUrl)
      },
      supervisorPostEval: {
        url: postUrl,
        qrCodeUrl: generateQRCode(postUrl)
      }
    });
  } catch (e) {
    return err('Failed to generate evaluation QRs: ' + e.message);
  }
}

