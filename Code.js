// ===== CONFIG =====
// Set these in Apps Script → Project Settings → Script Properties:
//   SHEET_ID              — Google Sheet ID for form responses
//   SUBMISSIONS_FOLDER_ID — "Abstract Submissions" folder ID in Shared Drive
//   TEMPLATE_DOC_ID       — Abstract PDF template Google Doc ID
//   EDIT_DEADLINE         — (optional) Cut-off date for edits, format: "2026-08-15"
const PROPS               = PropertiesService.getScriptProperties();
const SHEET_ID            = PROPS.getProperty('SHEET_ID');
const SUBMISSIONS_FOLDER_ID = PROPS.getProperty('SUBMISSIONS_FOLDER_ID');
const TEMPLATE_DOC_ID       = PROPS.getProperty('TEMPLATE_DOC_ID');

// Single source of truth for all column positions (0-based)
// A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V   W
// 0   1   2   3   4   5   6   7   8   9   10  11  12  13  14  15  16  17  18  19  20  21  22
const COL = {
  TIMESTAMP:        0,   // A — auto
  STUDENT_FIRST:    1,   // B — form
  STUDENT_LAST:     2,   // C — form
  STUDENT_EMAIL:    3,   // D — form
  STUDENTSHIP:      4,   // E — form
  STIPEND_OTHER:    5,   // F — form (populated when Stipend = "Other")
  SUP_FIRST:        6,   // G — form
  SUP_LAST:         7,   // H — form
  SUP_EMAIL:        8,   // I — form
  DEPARTMENT:       9,   // J — form
  INSTITUTE:        10,  // K — form (optional institute affiliation)
  TITLE:            11,  // L — form
  FIRST_AUTHOR:     12,  // M — form
  FIRST_AUTHOR_AFF: 13,  // N — form
  CO_AUTHORS:       14,  // O — form
  ABSTRACT_BODY:    15,  // P — form
  APPROVAL_STATUS:  16,  // Q — admin dropdown: Approved / Not Approved
  REVIEW_NOTES:     17,  // R — admin fills (included in rejection email)
  PDF_STATUS:       18,  // S — auto
  PDF_LINK:         19,  // T — auto (Drive URL)
  EMAIL_STATUS:     20,  // U — auto
  POSTER_NUMBER:    21,  // V — auto (assigned via assignPosterNumbers)
  EDIT_TOKEN:       22   // W — system (powers edit links)
};

// ===== SERVE FORM =====
// Data is injected by replacing </head> with a <script> block — no template syntax needed,
// so Index.html stays as plain HTML with no IDE errors.
function doGet(e) {
  const token = e && e.parameter && e.parameter.token;
  let prefillJson   = 'null';
  let editTokenJson = 'null';

  if (token) {
    const deadline = PROPS.getProperty('EDIT_DEADLINE');
    if (deadline && new Date() > new Date(deadline)) {
      return errorPage(
        'Submissions Closed',
        'The abstract editing period has ended. Please contact ' +
        '<a href="mailto:fmdugrad@ualberta.ca">fmdugrad@ualberta.ca</a> if you need assistance.'
      );
    }
    const existing = getSubmissionByToken(token);
    if (!existing) {
      return errorPage(
        'Invalid Link',
        'This edit link is invalid or has expired. Please contact ' +
        '<a href="mailto:fmdugrad@ualberta.ca">fmdugrad@ualberta.ca</a>.'
      );
    }
    prefillJson   = JSON.stringify(existing);
    editTokenJson = JSON.stringify(token);
  }

  const pageHtml  = HtmlService.createHtmlOutputFromFile('Index').getContent();
  const dataBlock = '<script>var PREFILL=' + prefillJson + ';var EDIT_TOKEN=' + editTokenJson + ';</script>';
  const finalHtml = pageHtml.replace('</head>', dataBlock + '\n</head>');

  return HtmlService.createHtmlOutput(finalHtml)
    .setTitle('2026 FoMD Summer Students\' Research Day – Abstract Submission');
}

function errorPage(title, body) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Roboto,sans-serif;padding:60px;text-align:center;color:#c0392b;">' +
    '<h2>' + title + '</h2><p>' + body + '</p></div>'
  );
}

// ===== ADMIN MENU =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Abstract Admin')
    .addItem('Open Admin Panel',                 'showSidebar')
    .addSeparator()
    .addItem('Assign Poster Numbers',            'assignPosterNumbers')
    .addSeparator()
    .addItem('Set Up Sheet Headers & Dropdowns', 'setupSheet')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Abstract Admin Tools');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ===== HANDLE FORM SUBMISSION (dispatcher) =====
function handleFormSubmit(formObj) {
  const isEdit = !!(formObj.editToken && formObj.editToken.trim());
  return isEdit ? handleEdit(formObj) : handleNewSubmission(formObj);
}

// ===== NEW SUBMISSION =====
function handleNewSubmission(formObj) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');

  let additionalAuthors = [];
  try { additionalAuthors = JSON.parse(formObj.additionalAuthors || '[]'); } catch (e) {}
  const coAuthorsFormatted = formatCoAuthors(additionalAuthors);

  const token = Utilities.getUuid();

  sheet.appendRow([
    new Date(),                                       // A: Timestamp
    formObj.studentFirstName,                         // B: Student First Name
    formObj.studentLastName,                          // C: Student Last Name
    formObj.studentEmail,                             // D: Student Email
    formObj.studentship,                              // E: Stipend Support
    formObj.stipendOther     || '',                   // F: Stipend Other
    formObj.supervisorFirstName,                      // G: Supervisor First Name
    formObj.supervisorLastName,                       // H: Supervisor Last Name
    formObj.supervisorEmail,                          // I: Supervisor Email
    formObj.supervisorDepartment,                     // J: Department
    formObj.supervisorInstitute  || '',               // K: Institute Affiliation
    formObj.title,                                    // L: Abstract Title
    formObj.firstAuthor,                              // M: First Author
    formObj.firstAuthorAffiliation,                   // N: First Author Affiliation
    coAuthorsFormatted,                               // O: Additional Authors
    formObj.abstractBody,                             // P: Abstract Body
    '',                                               // Q: Approval Status   ← admin
    '',                                               // R: Review Notes      ← admin
    '',                                               // S: PDF Status        ← auto
    '',                                               // T: PDF Drive Link    ← auto
    '',                                               // U: Email Status      ← auto
    '',                                               // V: Poster Number     ← auto
    token                                             // W: Edit Token        ← system
  ]);

  const lastRow = sheet.getLastRow();
  try {
    const pdf = generateAndSavePDF(formObj, additionalAuthors);
    const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    sheet.getRange(lastRow, COL.PDF_STATUS + 1).setValue('PDF Generated – ' + ts);
    sheet.getRange(lastRow, COL.PDF_LINK   + 1).setValue(pdf.url);
  } catch (err) {
    Logger.log('PDF generation failed: ' + err.message);
    sheet.getRange(lastRow, COL.PDF_STATUS + 1).setValue('PDF Error: ' + err.message);
  }

  const webAppUrl = ScriptApp.getService().getUrl();
  const editLink  = webAppUrl + '?token=' + token;

  MailApp.sendEmail({
    to: formObj.studentEmail,
    subject: "2026 FoMD Summer Students' Research Day – Abstract Submission Confirmation",
    htmlBody: `
      <p>Dear ${formObj.studentFirstName},</p>
      <p>Thank you for submitting your abstract to the <strong>2026 FoMD Summer Students' Research Day</strong>.</p>
      <p><strong>Abstract Title:</strong> ${formObj.title}</p>
      <p>Your abstract has been received and is under review. You will be contacted once a decision has been made.</p>
      <hr style="margin:20px 0;">
      <p><strong>Need to make changes?</strong><br>
         You can edit your abstract at any time before the submission deadline using the link below.
         Keep this email — it is the only way to access your submission for editing.</p>
      <p><a href="${editLink}" style="background:#005C2E;color:#fff;padding:10px 20px;border-radius:5px;text-decoration:none;font-weight:bold;">Edit My Abstract</a></p>
      <hr style="margin:20px 0;">
      <p>If you have any questions, please contact the FoMD Undergraduate Research Program Admin at
         <a href="mailto:fmdugrad@ualberta.ca">fmdugrad@ualberta.ca</a>.</p>
      <br>
      <p>Kind regards,<br>FoMD Undergraduate Research Program</p>
    `
  });

  return HtmlService.createHtmlOutputFromFile('Success').getContent();
}

// ===== EDIT EXISTING SUBMISSION =====
function handleEdit(formObj) {
  const sheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data   = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.EDIT_TOKEN] === formObj.editToken) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) throw new Error('Edit token not found. Your link may be invalid.');

  let additionalAuthors = [];
  try { additionalAuthors = JSON.parse(formObj.additionalAuthors || '[]'); } catch (e) {}
  const coAuthorsFormatted = formatCoAuthors(additionalAuthors);

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  sheet.getRange(rowIndex, COL.TIMESTAMP        + 1).setValue(new Date());
  sheet.getRange(rowIndex, COL.STUDENT_FIRST    + 1).setValue(formObj.studentFirstName);
  sheet.getRange(rowIndex, COL.STUDENT_LAST     + 1).setValue(formObj.studentLastName);
  sheet.getRange(rowIndex, COL.STUDENT_EMAIL    + 1).setValue(formObj.studentEmail);
  sheet.getRange(rowIndex, COL.STUDENTSHIP      + 1).setValue(formObj.studentship);
  sheet.getRange(rowIndex, COL.STIPEND_OTHER    + 1).setValue(formObj.stipendOther    || '');
  sheet.getRange(rowIndex, COL.SUP_FIRST        + 1).setValue(formObj.supervisorFirstName);
  sheet.getRange(rowIndex, COL.SUP_LAST         + 1).setValue(formObj.supervisorLastName);
  sheet.getRange(rowIndex, COL.SUP_EMAIL        + 1).setValue(formObj.supervisorEmail);
  sheet.getRange(rowIndex, COL.DEPARTMENT       + 1).setValue(formObj.supervisorDepartment);
  sheet.getRange(rowIndex, COL.INSTITUTE        + 1).setValue(formObj.supervisorInstitute || '');
  sheet.getRange(rowIndex, COL.TITLE            + 1).setValue(formObj.title);
  sheet.getRange(rowIndex, COL.FIRST_AUTHOR     + 1).setValue(formObj.firstAuthor);
  sheet.getRange(rowIndex, COL.FIRST_AUTHOR_AFF + 1).setValue(formObj.firstAuthorAffiliation);
  sheet.getRange(rowIndex, COL.CO_AUTHORS       + 1).setValue(coAuthorsFormatted);
  sheet.getRange(rowIndex, COL.ABSTRACT_BODY    + 1).setValue(formObj.abstractBody);
  // Q (Approval Status), R (Review Notes), V (Poster Number), W (Edit Token) preserved

  sheet.getRange(rowIndex, COL.PDF_STATUS + 1).setValue('');
  sheet.getRange(rowIndex, COL.PDF_LINK   + 1).setValue('');
  try {
    const pdf = generateAndSavePDF(formObj, additionalAuthors);
    sheet.getRange(rowIndex, COL.PDF_STATUS + 1).setValue('PDF Updated – ' + ts);
    sheet.getRange(rowIndex, COL.PDF_LINK   + 1).setValue(pdf.url);
  } catch (err) {
    Logger.log('PDF regen on edit failed: ' + err.message);
    sheet.getRange(rowIndex, COL.PDF_STATUS + 1).setValue('PDF Error: ' + err.message);
  }

  MailApp.sendEmail({
    to: formObj.studentEmail,
    subject: "2026 FoMD Summer Students' Research Day – Abstract Updated",
    htmlBody: `
      <p>Dear ${formObj.studentFirstName},</p>
      <p>Your abstract submission has been successfully updated.</p>
      <p><strong>Abstract Title:</strong> ${formObj.title}</p>
      <p>If you need to make further changes before the submission deadline, use the same edit link from your original confirmation email.</p>
      <p>If you have any questions, please contact
         <a href="mailto:fmdugrad@ualberta.ca">fmdugrad@ualberta.ca</a>.</p>
      <br>
      <p>Kind regards,<br>FoMD Undergraduate Research Program</p>
    `
  });

  return HtmlService.createHtmlOutputFromFile('Success').getContent();
}

// ===== TOKEN LOOKUP =====
function getSubmissionByToken(token) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.EDIT_TOKEN] !== token) continue;
    const row = data[i];

    const coAuthorsStr      = (row[COL.CO_AUTHORS] || '').toString();
    const additionalAuthors = coAuthorsStr
      .split(';')
      .map(s => {
        const m = s.trim().match(/^(.+?)\s*\((.+?)\)$/);
        return m ? { name: m[1].trim(), affiliation: m[2].trim() } : { name: s.trim(), affiliation: '' };
      })
      .filter(a => a.name !== '');

    return {
      studentFirstName:     row[COL.STUDENT_FIRST],
      studentLastName:      row[COL.STUDENT_LAST],
      studentEmail:         row[COL.STUDENT_EMAIL],
      studentship:          row[COL.STUDENTSHIP],
      stipendOther:         row[COL.STIPEND_OTHER],
      supervisorFirstName:  row[COL.SUP_FIRST],
      supervisorLastName:   row[COL.SUP_LAST],
      supervisorEmail:      row[COL.SUP_EMAIL],
      supervisorDepartment: row[COL.DEPARTMENT],
      supervisorInstitute:  row[COL.INSTITUTE],
      title:                row[COL.TITLE],
      firstAuthor:          row[COL.FIRST_AUTHOR],
      firstAuthorAffiliation: row[COL.FIRST_AUTHOR_AFF],
      additionalAuthors:    additionalAuthors,
      abstractBody:         row[COL.ABSTRACT_BODY]
    };
  }
  return null;
}

// ===== HELPER: format co-authors for the sheet =====
function formatCoAuthors(authors) {
  return authors
    .map(a => a.affiliation ? `${a.name} (${a.affiliation})` : a.name)
    .join('; ');
}

// ===== PDF GENERATION FROM GOOGLE DOC TEMPLATE =====
// Template Google Doc should contain these placeholder strings (each on its own line/paragraph):
//   {{TITLE}}                    {{STUDENT_NAME}}           {{STUDENT_EMAIL}}
//   {{STUDENTSHIP}}              {{SUPERVISOR_NAME}}        {{SUPERVISOR_EMAIL}}
//   {{DEPARTMENT}}               {{INSTITUTE}}              {{FIRST_AUTHOR}}
//   {{FIRST_AUTHOR_AFFILIATION}} {{CO_AUTHORS}}             {{ABSTRACT_BODY}}
//   {{SUBMISSION_DATE}}
//
// Design rules:
//   - Apply all desired formatting TO the placeholder text itself — it is inherited by content.
//   - {{ABSTRACT_BODY}} and {{CO_AUTHORS}} must each be alone in their own paragraph.
//   - All other placeholders can sit inline with label text, e.g. "Title: {{TITLE}}"
function generateAndSavePDF(formObj, additionalAuthors) {
  if (!SUBMISSIONS_FOLDER_ID) throw new Error('SUBMISSIONS_FOLDER_ID not set in Script Properties');
  if (!TEMPLATE_DOC_ID)       throw new Error('TEMPLATE_DOC_ID not set in Script Properties');

  const folder   = DriveApp.getFolderById(SUBMISSIONS_FOLDER_ID);
  const template = DriveApp.getFileById(TEMPLATE_DOC_ID);

  const safeName = 'Abstract_' + formObj.studentLastName + '_' +
    formObj.title.replace(/[^a-zA-Z0-9 ]/g, '').substring(0, 40).trim();

  const docCopy = template.makeCopy(safeName, folder);
  const doc     = DocumentApp.openById(docCopy.getId());
  const body    = doc.getBody();

  const coAuthorsLines = additionalAuthors.length > 0
    ? additionalAuthors.map(a => a.affiliation ? a.name + ', ' + a.affiliation : a.name)
    : ['None'];
  replaceWithParagraphs(body, '{{CO_AUTHORS}}',    coAuthorsLines);
  replaceWithParagraphs(body, '{{ABSTRACT_BODY}}', (formObj.abstractBody || '').split('\n'));

  const stipendDisplay = formObj.studentship === 'Other' && formObj.stipendOther
    ? 'Other – ' + formObj.stipendOther
    : (formObj.studentship || '');

  const singleLine = {
    '{{TITLE}}':                    formObj.title               || '',
    '{{STUDENT_NAME}}':             formObj.studentFirstName + ' ' + formObj.studentLastName,
    '{{STUDENT_EMAIL}}':            formObj.studentEmail         || '',
    '{{STUDENTSHIP}}':              stipendDisplay,
    '{{SUPERVISOR_NAME}}':          formObj.supervisorFirstName + ' ' + formObj.supervisorLastName,
    '{{SUPERVISOR_EMAIL}}':         formObj.supervisorEmail      || '',
    '{{DEPARTMENT}}':               formObj.supervisorDepartment || '',
    '{{INSTITUTE}}':                formObj.supervisorInstitute  || '',
    '{{FIRST_AUTHOR}}':             formObj.firstAuthor          || '',
    '{{FIRST_AUTHOR_AFFILIATION}}': formObj.firstAuthorAffiliation || '',
    '{{SUBMISSION_DATE}}':          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy')
  };

  for (const [placeholder, value] of Object.entries(singleLine)) {
    body.replaceText(placeholder, escapeDollarSigns(value));
  }

  doc.saveAndClose();

  const pdfBlob = docCopy.getAs('application/pdf').setName(safeName + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  docCopy.setTrashed(true);

  return { blob: pdfBlob, url: pdfFile.getUrl() };
}

// Replaces a placeholder paragraph with one paragraph per line, inheriting the placeholder's style.
function replaceWithParagraphs(body, placeholder, lines) {
  const result = body.findText(placeholder);
  if (!result) return;

  const element   = result.getElement();
  const para      = element.getType() === DocumentApp.ElementType.TEXT ? element.getParent() : element;
  const paraIndex = body.getChildIndex(para);
  const paraAttrs = para.getAttributes();
  const textAttrs = element.getType() === DocumentApp.ElementType.TEXT ? element.getAttributes() : {};

  const nonEmpty = lines.filter(l => l.trim() !== '');
  if (nonEmpty.length === 0) nonEmpty.push('');

  nonEmpty.slice().reverse().forEach(line => {
    const newPara = body.insertParagraph(paraIndex, line);
    newPara.setAttributes(paraAttrs);
    if (newPara.editAsText) newPara.editAsText().setAttributes(textAttrs);
  });

  body.removeChild(para);
}

function escapeDollarSigns(str) {
  return str.replace(/\$/g, '$$$$');
}

// ===== BATCH PDF GENERATION =====
function generateAbstractPDFs() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data  = sheet.getDataRange().getValues();
  let processed = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[COL.PDF_STATUS] && row[COL.PDF_STATUS].toString().trim() !== '') continue;

    const coAuthorsStr = (row[COL.CO_AUTHORS] || '').toString();
    const additionalAuthors = coAuthorsStr
      .split(';')
      .map(s => {
        const m = s.trim().match(/^(.+?)\s*\((.+?)\)$/);
        return m ? { name: m[1].trim(), affiliation: m[2].trim() } : { name: s.trim(), affiliation: '' };
      })
      .filter(a => a.name !== '');

    const formObj = {
      studentFirstName:      row[COL.STUDENT_FIRST],
      studentLastName:       row[COL.STUDENT_LAST],
      studentEmail:          row[COL.STUDENT_EMAIL],
      studentship:           row[COL.STUDENTSHIP],
      stipendOther:          row[COL.STIPEND_OTHER],
      supervisorFirstName:   row[COL.SUP_FIRST],
      supervisorLastName:    row[COL.SUP_LAST],
      supervisorEmail:       row[COL.SUP_EMAIL],
      supervisorDepartment:  row[COL.DEPARTMENT],
      supervisorInstitute:   row[COL.INSTITUTE],
      title:                 row[COL.TITLE],
      firstAuthor:           row[COL.FIRST_AUTHOR],
      firstAuthorAffiliation:row[COL.FIRST_AUTHOR_AFF],
      abstractBody:          row[COL.ABSTRACT_BODY]
    };

    try {
      const pdf = generateAndSavePDF(formObj, additionalAuthors);
      const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      sheet.getRange(i + 1, COL.PDF_STATUS + 1).setValue('PDF Generated – ' + ts);
      sheet.getRange(i + 1, COL.PDF_LINK   + 1).setValue(pdf.url);
      processed++;
    } catch (err) {
      Logger.log('Row ' + (i + 1) + ' PDF error: ' + err.message);
      sheet.getRange(i + 1, COL.PDF_STATUS + 1).setValue('Error: ' + err.message);
    }
  }

  return 'Generated ' + processed + ' new PDF(s).';
}

// ===== ASSIGN POSTER NUMBERS =====
function assignPosterNumbers() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data  = sheet.getDataRange().getValues();

  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const existing = (data[i][COL.POSTER_NUMBER] || '').toString();
    const match = existing.match(/(\d+)/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1], 10));
  }

  let assigned = 0;
  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    const approval = (row[COL.APPROVAL_STATUS] || '').toString().trim().toLowerCase();
    const hasPoster = row[COL.POSTER_NUMBER] && row[COL.POSTER_NUMBER].toString().trim() !== '';
    if (approval === 'approved' && !hasPoster) {
      maxNum++;
      sheet.getRange(i + 1, COL.POSTER_NUMBER + 1).setValue('P-' + maxNum.toString().padStart(3, '0'));
      assigned++;
    }
  }

  const msg = 'Assigned ' + assigned + ' poster number(s). Next available: P-' + (maxNum + 1).toString().padStart(3, '0') + '.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ===== BATCH APPROVAL EMAILS =====
function sendBatchApprovalEmails() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data  = sheet.getDataRange().getValues();
  let sent    = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((row[COL.APPROVAL_STATUS] || '').toString().trim().toLowerCase() !== 'approved') continue;
    if (row[COL.EMAIL_STATUS] && row[COL.EMAIL_STATUS].toString().trim() !== '') continue;

    const name      = row[COL.STUDENT_FIRST];
    const email     = row[COL.STUDENT_EMAIL];
    const title     = row[COL.TITLE];
    const posterNum = row[COL.POSTER_NUMBER]
      ? '<p><strong>Poster Number:</strong> ' + row[COL.POSTER_NUMBER] + '</p>' : '';

    MailApp.sendEmail({
      to: email,
      subject: "2026 FoMD Summer Students' Research Day – Abstract Approved",
      htmlBody: `
        <p>Dear ${name},</p>
        <p>We are pleased to inform you that your abstract submission has been <strong>approved</strong> for
           the <strong>2026 FoMD Summer Students' Research Day</strong>.</p>
        <p><strong>Abstract Title:</strong> ${title}</p>
        ${posterNum}
        <p>Further details regarding the event will be sent to you in due course.</p>
        <p>Congratulations, and we look forward to seeing your presentation!</p>
        <br>
        <p>Kind regards,<br>FoMD Undergraduate Research Program</p>
      `
    });

    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    sheet.getRange(i + 1, COL.EMAIL_STATUS + 1).setValue('Approval Sent – ' + ts);
    sent++;
  }

  return 'Sent ' + sent + ' approval email(s).';
}

// ===== BATCH REJECTION EMAILS =====
function sendBatchRejectionEmails() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data  = sheet.getDataRange().getValues();
  let sent    = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((row[COL.APPROVAL_STATUS] || '').toString().trim().toLowerCase() !== 'not approved') continue;
    if (row[COL.EMAIL_STATUS] && row[COL.EMAIL_STATUS].toString().trim() !== '') continue;

    const name    = row[COL.STUDENT_FIRST];
    const email   = row[COL.STUDENT_EMAIL];
    const title   = row[COL.TITLE];
    const message = (row[COL.REVIEW_NOTES] || '').toString().trim();

    MailApp.sendEmail({
      to: email,
      subject: "2026 FoMD Summer Students' Research Day – Abstract Submission Update",
      htmlBody: `
        <p>Dear ${name},</p>
        <p>Thank you for submitting your abstract to the <strong>2026 FoMD Summer Students' Research Day</strong>.</p>
        <p><strong>Abstract Title:</strong> ${title}</p>
        <p>After careful review, we regret to inform you that your abstract was not selected for presentation at this year's event.</p>
        ${message ? `<p><strong>Reviewer Comments:</strong></p><p>${message}</p>` : ''}
        <p>We encourage you to continue your research and hope to see you at future events.</p>
        <p>If you have any questions, please contact
           <a href="mailto:fmdugrad@ualberta.ca">fmdugrad@ualberta.ca</a>.</p>
        <br>
        <p>Kind regards,<br>FoMD Undergraduate Research Program</p>
      `
    });

    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    sheet.getRange(i + 1, COL.EMAIL_STATUS + 1).setValue('Rejection Sent – ' + ts);
    sent++;
  }

  return 'Sent ' + sent + ' rejection email(s).';
}

// ===== SHEET SETUP (run once) =====
function setupSheet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');

  const headers = [
    'Timestamp',                // A
    'Student First Name',       // B
    'Student Last Name',        // C
    'Student Email',            // D
    'Stipend Support',          // E
    'Stipend (Other)',          // F
    'Supervisor First Name',    // G
    'Supervisor Last Name',     // H
    'Supervisor Email',         // I
    'Department',               // J
    'Institute Affiliation',    // K
    'Abstract Title',           // L
    'First Author',             // M
    'First Author Affiliation', // N
    'Additional Authors',       // O
    'Abstract Body',            // P
    'Approval Status',          // Q ← admin dropdown
    'Review Notes',             // R ← admin fills
    'PDF Status',               // S ← auto
    'PDF Drive Link',           // T ← auto
    'Email Status',             // U ← auto
    'Poster Number',            // V ← auto
    'Edit Token'                // W ← system
  ];

  // If row 1 already has submission data (not a header), insert a blank row first
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell !== '' && firstCell !== 'Timestamp') {
    sheet.insertRowBefore(1);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#003C1E')
    .setFontColor('#ffffff');

  const approvalRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Approved', 'Not Approved'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, COL.APPROVAL_STATUS + 1, sheet.getMaxRows() - 1, 1).setDataValidation(approvalRule);

  Logger.log('Sheet headers and dropdowns set up successfully.');
}

// ===== UTILITIES =====
function ping() {
  return true;
}
