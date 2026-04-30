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
// A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V
// 0   1   2   3   4   5   6   7   8   9   10  11  12  13  14  15  16  17  18  19  20  21
const COL = {
  TIMESTAMP:       0,   // A — auto
  STUDENT_FIRST:   1,   // B — form
  STUDENT_LAST:    2,   // C — form
  STUDENT_EMAIL:   3,   // D — form
  STUDENTSHIP:     4,   // E — form
  STIPEND_OTHER:   5,   // F — form (populated when Stipend = "Other")
  SUP_FIRST:       6,   // G — form
  SUP_LAST:        7,   // H — form
  SUP_EMAIL:       8,   // I — form
  DEPARTMENT:      9,   // J — form
  INSTITUTE:       10,  // K — form (optional institute affiliation)
  TITLE:           11,  // L — form
  FIRST_AUTHOR:    12,  // M — form
  CO_AUTHORS:      13,  // N — form
  ABSTRACT_BODY:   14,  // O — form
  APPROVAL_STATUS: 15,  // P — admin dropdown: Approved / Not Approved
  REVIEW_NOTES:    16,  // Q — admin fills (included in rejection email)
  PDF_STATUS:      17,  // R — auto
  PDF_LINK:        18,  // S — auto (Drive URL)
  EMAIL_STATUS:    19,  // T — auto
  POSTER_NUMBER:   20,  // U — auto (assigned via assignPosterNumbers)
  EDIT_TOKEN:      21   // V — system (powers edit links)
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
    .addItem('Open Admin Sidebar',               'showSidebar')
    .addItem('Open Abstract Review (expanded)',  'showReviewDialog')
    .addItem('Customize Email Templates',        'showEmailTemplateDialog')
    .addSeparator()
    .addItem('Send Approval Emails',             'sendBatchApprovalEmails')
    .addItem('Send Rejection Emails',            'sendBatchRejectionEmails')
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

function showReviewDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReviewDialog')
    .setWidth(800)
    .setHeight(620);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Abstract Review');
}

function showEmailTemplateDialog() {
  const html = HtmlService.createHtmlOutputFromFile('EmailTemplateDialog')
    .setWidth(820)
    .setHeight(640);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Email Template Editor');
}

// Install trigger to ensure menu appears on sheet open
function installTrigger() {
  // Menu trigger
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SHEET_ID)
    .onOpen()
    .create();
  
  // Color coding trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SHEET_ID)
    .onEdit()
    .create();
  
  return 'Triggers installed. Refresh the sheet.';
}

// ===== ON EDIT TRIGGER FOR ROW COLOR CODING + AUTO POSTER ASSIGNMENT =====
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Form Responses') return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (col === COL.APPROVAL_STATUS + 1 && row > 1) {
    const status = (e.value || '').toString().trim().toLowerCase();
    const rowRange = sheet.getRange(row, 1, 1, 22);

    if (status === 'approved') {
      rowRange.setBackgroundColor('#D9F2D9');
      const existingPoster = sheet.getRange(row, COL.POSTER_NUMBER + 1).getValue();
      if (!existingPoster || existingPoster.toString().trim() === '') {
        const next = getNextPosterNumber(sheet);
        sheet.getRange(row, COL.POSTER_NUMBER + 1)
          .setValue(next.toString().padStart(3, '0'));
      }
    } else if (status === 'not approved') {
      rowRange.setBackgroundColor('#F2D9D9');
      sheet.getRange(row, COL.POSTER_NUMBER + 1).setValue('');
    } else {
      rowRange.setBackgroundColor('#FFFACD');
      sheet.getRange(row, COL.POSTER_NUMBER + 1).setValue('');
    }
  }
}

// Returns the next sequential poster number (integer) across all existing assignments.
function getNextPosterNumber(sheet) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const match = (data[i][COL.POSTER_NUMBER] || '').toString().match(/(\d+)/);
    if (match) max = Math.max(max, parseInt(match[1], 10));
  }
  return max + 1;
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
    coAuthorsFormatted,                               // N: Additional Authors
    formObj.abstractBody,                             // O: Abstract Body
    'Unprocessed',                                    // P: Approval Status   ← admin
    '',                                               // Q: Review Notes      ← admin
    '',                                               // R: PDF Status        ← auto
    '',                                               // S: PDF Drive Link    ← auto
    '',                                               // T: Email Status      ← auto
    '',                                               // U: Poster Number     ← auto
    token                                             // V: Edit Token        ← system
  ]);

  const lastRow = sheet.getLastRow();
  
  // Apply yellow background for Unprocessed status
  sheet.getRange(lastRow, 1, 1, 22).setBackgroundColor('#FFFACD'); // Soft yellow (lemon chiffon)
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
      .map(s => ({ name: s.trim() }))
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
      additionalAuthors:    additionalAuthors,
      abstractBody:         row[COL.ABSTRACT_BODY]
    };
  }
  return null;
}

// ===== HELPER: format co-authors for the sheet =====
function formatCoAuthors(authors) {
  return authors.map(a => a.name).join('; ');
}

// ===== PDF GENERATION FROM GOOGLE DOC TEMPLATE =====
// Template Google Doc should contain these placeholder strings (each on its own line/paragraph):
//   {{TITLE}}           {{STUDENT_NAME}}    {{STUDENT_EMAIL}}
//   {{STUDENTSHIP}}     {{SUPERVISOR_NAME}} {{SUPERVISOR_EMAIL}}
//   {{DEPARTMENT}}      {{INSTITUTE}}       {{FIRST_AUTHOR}}
//   {{CO_AUTHORS}}      {{ABSTRACT_BODY}}   {{SUBMISSION_DATE}}
//
// Design rules:
//   - Apply all desired formatting TO the placeholder text itself — it is inherited by content.
//   - {{ABSTRACT_BODY}} and {{CO_AUTHORS}} must each be alone in their own paragraph.
//   - All other placeholders can sit inline with label text, e.g. "Title: {{TITLE}}"
function generateAndSavePDF(formObj, additionalAuthors) {
  if (!SUBMISSIONS_FOLDER_ID) throw new Error('SUBMISSIONS_FOLDER_ID not set in Script Properties');
  if (!TEMPLATE_DOC_ID)       throw new Error('TEMPLATE_DOC_ID not set in Script Properties');

  const submissionDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const safeName = formObj.studentLastName + '_' + submissionDate;

  const copyMeta = Drive.Files.copy({ title: safeName }, TEMPLATE_DOC_ID, { supportsAllDrives: true });
  const copyId   = copyMeta.id;

  const doc  = DocumentApp.openById(copyId);
  const body = doc.getBody();

  const coAuthorNames = additionalAuthors.map(a => a.name).filter(Boolean);
  const coAuthorsLine = coAuthorNames.length > 0 ? ', ' + coAuthorNames.join(', ') : '';
  body.replaceText('\\{\\{CO_AUTHORS\\}\\}', escapeDollarSigns(coAuthorsLine));
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
    '{{FIRST_AUTHOR}}':    formObj.firstAuthor || '',
    '{{SUBMISSION_DATE}}': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy')
  };

  for (const [placeholder, value] of Object.entries(singleLine)) {
    body.replaceText(placeholder, escapeDollarSigns(value));
  }

  doc.saveAndClose();

  // Export as PDF using the Document service's built-in method (more reliable than URL fetch)
  const pdfBlob = doc.getAs('application/pdf').setName(safeName + '.pdf');

  // Create the PDF in the submissions folder using Drive Advanced Service
  // Use DriveApp for better compatibility with Shared Drives
  const folder = DriveApp.getFolderById(SUBMISSIONS_FOLDER_ID);
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setName(safeName + '.pdf');

  // Trash the temporary doc copy
  Drive.Files.trash(copyId, { supportsAllDrives: true });

  return { blob: pdfBlob, url: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view' };
}

// Replaces a placeholder paragraph with one paragraph per line, inheriting the placeholder's style.
// The first line replaces the placeholder text in-place so the original paragraph (with all
// of its formatting) is preserved exactly. Additional lines are inserted as copies.
// insertParagraph() can strip paragraph-level indent attributes from the inserted clone, so
// we re-apply them explicitly using direct setter methods rather than setAttributes().
function replaceWithParagraphs(body, placeholder, lines) {
  const result = body.findText(placeholder);
  if (!result) return;

  const element   = result.getElement();
  const para      = element.getType() === DocumentApp.ElementType.TEXT ? element.getParent() : element;

  const nonEmpty = lines.filter(l => l.trim() !== '');
  if (nonEmpty.length === 0) nonEmpty.push('');

  // Replace text in the first paragraph (preserves original paragraph object and its attributes)
  para.editAsText().setText(nonEmpty[0]);

  // Read paragraph-level style from the template paragraph once, before the loop
  const alignment       = para.getAlignment();
  const heading         = para.getHeading();
  const indentStart     = para.getIndentStart();
  const indentFirstLine = para.getIndentFirstLine();
  const indentEnd       = para.getIndentEnd();
  const lineSpacing     = para.getLineSpacing();
  const spacingBefore   = para.getSpacingBefore();
  const spacingAfter    = para.getSpacingAfter();

  // Insert additional lines as new paragraphs, re-applying paragraph formatting explicitly
  for (let i = 1; i < nonEmpty.length; i++) {
    const insertIndex = body.getChildIndex(para) + i;
    const newPara = para.copy();
    newPara.editAsText().setText(nonEmpty[i]);
    const inserted = body.insertParagraph(insertIndex, newPara);
    // Re-apply paragraph-level formatting via direct setters (more reliable than setAttributes)
    if (alignment       !== null) inserted.setAlignment(alignment);
    if (heading         !== null) inserted.setHeading(heading);
    if (indentStart     !== null) inserted.setIndentStart(indentStart);
    if (indentFirstLine !== null) inserted.setIndentFirstLine(indentFirstLine);
    if (indentEnd       !== null) inserted.setIndentEnd(indentEnd);
    if (lineSpacing     !== null) inserted.setLineSpacing(lineSpacing);
    if (spacingBefore   !== null) inserted.setSpacingBefore(spacingBefore);
    if (spacingAfter    !== null) inserted.setSpacingAfter(spacingAfter);
  }
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
      .map(s => ({ name: s.trim() }))
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
      firstAuthor:  row[COL.FIRST_AUTHOR],
      abstractBody: row[COL.ABSTRACT_BODY]
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
      sheet.getRange(i + 1, COL.POSTER_NUMBER + 1).setValue(maxNum.toString().padStart(3, '0'));
      assigned++;
    }
  }

  const msg = 'Assigned ' + assigned + ' poster number(s). Next available: ' + (maxNum + 1).toString().padStart(3, '0') + '.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ===== EMAIL TEMPLATE MANAGEMENT =====
function getDefaultEmailTemplates() {
  return {
    approvalSubject:  "2026 FoMD Summer Students' Research Day – Abstract Approved",
    approvalBody:
      '<p>Dear {{STUDENT_FIRST}},</p>\n' +
      '<p>We are pleased to inform you that your abstract submission has been <strong>approved</strong> for the <strong>2026 FoMD Summer Students’ Research Day</strong>.</p>\n' +
      '<p><strong>Abstract Title:</strong> {{TITLE}}</p>\n' +
      '{{POSTER_BLOCK}}\n' +
      '<p>Further details regarding the event will be sent to you in due course.</p>\n' +
      '<p>Congratulations, and we look forward to seeing your presentation!</p>\n' +
      '<br>\n' +
      '<p>Kind regards,<br>FoMD Undergraduate Research Program</p>',

    rejectionSubject: "2026 FoMD Summer Students' Research Day – Abstract Submission Update",
    rejectionBody:
      '<p>Dear {{STUDENT_FIRST}},</p>\n' +
      '<p>Thank you for submitting your abstract to the <strong>2026 FoMD Summer Students’ Research Day</strong>.</p>\n' +
      '<p><strong>Abstract Title:</strong> {{TITLE}}</p>\n' +
      '<p>After careful review, we regret to inform you that your abstract was not selected for presentation at this year’s event.</p>\n' +
      '{{REVIEWER_NOTES_BLOCK}}\n' +
      '<p>We encourage you to continue your research and hope to see you at future events.</p>\n' +
      '<p>If you have any questions, please contact <a href="mailto:fmdugrad@ualberta.ca">fmdugrad@ualberta.ca</a>.</p>\n' +
      '<br>\n' +
      '<p>Kind regards,<br>FoMD Undergraduate Research Program</p>'
  };
}

function getEmailTemplates() {
  const defaults = getDefaultEmailTemplates();
  return {
    approvalSubject:  PROPS.getProperty('APPROVAL_EMAIL_SUBJECT')  || defaults.approvalSubject,
    approvalBody:     PROPS.getProperty('APPROVAL_EMAIL_BODY')     || defaults.approvalBody,
    rejectionSubject: PROPS.getProperty('REJECTION_EMAIL_SUBJECT') || defaults.rejectionSubject,
    rejectionBody:    PROPS.getProperty('REJECTION_EMAIL_BODY')    || defaults.rejectionBody
  };
}

function saveEmailTemplates(data) {
  PROPS.setProperty('APPROVAL_EMAIL_SUBJECT',  data.approvalSubject  || '');
  PROPS.setProperty('APPROVAL_EMAIL_BODY',     data.approvalBody     || '');
  PROPS.setProperty('REJECTION_EMAIL_SUBJECT', data.rejectionSubject || '');
  PROPS.setProperty('REJECTION_EMAIL_BODY',    data.rejectionBody    || '');
}

function resetEmailTemplatesToDefault() {
  const defaults = getDefaultEmailTemplates();
  PROPS.setProperty('APPROVAL_EMAIL_SUBJECT',  defaults.approvalSubject);
  PROPS.setProperty('APPROVAL_EMAIL_BODY',     defaults.approvalBody);
  PROPS.setProperty('REJECTION_EMAIL_SUBJECT', defaults.rejectionSubject);
  PROPS.setProperty('REJECTION_EMAIL_BODY',    defaults.rejectionBody);
  return defaults;
}

function applyEmailTemplate(template, values) {
  const posterBlock = values.posterNumber
    ? '<p><strong>Poster Number:</strong> ' + values.posterNumber + '</p>'
    : '';
  const notesBlock = values.reviewerNotes
    ? '<p><strong>Reviewer Comments:</strong></p><p>' + values.reviewerNotes + '</p>'
    : '';

  return template
    .replace(/\{\{STUDENT_FIRST\}\}/g,         values.studentFirst   || '')
    .replace(/\{\{STUDENT_LAST\}\}/g,          values.studentLast    || '')
    .replace(/\{\{STUDENT_NAME\}\}/g,          values.studentFirst   || '')
    .replace(/\{\{TITLE\}\}/g,                 values.title          || '')
    .replace(/\{\{POSTER_NUMBER\}\}/g,         values.posterNumber   || '')
    .replace(/\{\{POSTER_BLOCK\}\}/g,          posterBlock)
    .replace(/\{\{REVIEWER_NOTES\}\}/g,        values.reviewerNotes  || '')
    .replace(/\{\{REVIEWER_NOTES_BLOCK\}\}/g,  notesBlock);
}

// ===== BATCH APPROVAL EMAILS =====
function sendBatchApprovalEmails() {
  const sheet     = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data      = sheet.getDataRange().getValues();
  const templates = getEmailTemplates();
  let sent        = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((row[COL.APPROVAL_STATUS] || '').toString().trim().toLowerCase() !== 'approved') continue;
    if (row[COL.EMAIL_STATUS] && row[COL.EMAIL_STATUS].toString().trim() !== '') continue;

    const values = {
      studentFirst:  (row[COL.STUDENT_FIRST]  || '').toString(),
      studentLast:   (row[COL.STUDENT_LAST]   || '').toString(),
      title:         (row[COL.TITLE]          || '').toString(),
      posterNumber:  (row[COL.POSTER_NUMBER]  || '').toString().trim()
    };

    MailApp.sendEmail({
      to:       row[COL.STUDENT_EMAIL],
      subject:  applyEmailTemplate(templates.approvalSubject, values),
      htmlBody: applyEmailTemplate(templates.approvalBody,    values)
    });

    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    sheet.getRange(i + 1, COL.EMAIL_STATUS + 1).setValue('Approval Sent – ' + ts);
    sent++;
  }

  return 'Sent ' + sent + ' approval email(s).';
}

// ===== BATCH REJECTION EMAILS =====
function sendBatchRejectionEmails() {
  const sheet     = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data      = sheet.getDataRange().getValues();
  const templates = getEmailTemplates();
  let sent        = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((row[COL.APPROVAL_STATUS] || '').toString().trim().toLowerCase() !== 'not approved') continue;
    if (row[COL.EMAIL_STATUS] && row[COL.EMAIL_STATUS].toString().trim() !== '') continue;

    const values = {
      studentFirst:  (row[COL.STUDENT_FIRST]  || '').toString(),
      studentLast:   (row[COL.STUDENT_LAST]   || '').toString(),
      title:         (row[COL.TITLE]          || '').toString(),
      reviewerNotes: (row[COL.REVIEW_NOTES]   || '').toString().trim()
    };

    MailApp.sendEmail({
      to:       row[COL.STUDENT_EMAIL],
      subject:  applyEmailTemplate(templates.rejectionSubject, values),
      htmlBody: applyEmailTemplate(templates.rejectionBody,    values)
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
    'Timestamp',             // A
    'Student First Name',    // B
    'Student Last Name',     // C
    'Student Email',         // D
    'Stipend Support',       // E
    'Stipend (Other)',       // F
    'Supervisor First Name', // G
    'Supervisor Last Name',  // H
    'Supervisor Email',      // I
    'Department',            // J
    'Institute Affiliation', // K
    'Abstract Title',        // L
    'First Author',          // M
    'Additional Authors',    // N
    'Abstract Body',         // O
    'Approval Status',       // P ← admin dropdown
    'Review Notes',          // Q ← admin fills
    'PDF Status',            // R ← auto
    'PDF Drive Link',        // S ← auto
    'Email Status',          // T ← auto
    'Poster Number',         // U ← auto
    'Edit Token'             // V ← system
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

// ===== SIDEBAR: GET ALL ABSTRACTS FOR REVIEW =====
function getAbstractsForReview() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data  = sheet.getDataRange().getValues();
  const abstracts = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.TIMESTAMP]) continue;

    abstracts.push({
      rowIndex:      i + 1,
      studentFirst:  row[COL.STUDENT_FIRST]  || '',
      studentLast:   row[COL.STUDENT_LAST]   || '',
      studentEmail:  row[COL.STUDENT_EMAIL]  || '',
      supervisorFirst: row[COL.SUP_FIRST]    || '',
      supervisorLast:  row[COL.SUP_LAST]     || '',
      department:    row[COL.DEPARTMENT]     || '',
      title:         row[COL.TITLE]          || '',
      firstAuthor:   row[COL.FIRST_AUTHOR]   || '',
      coAuthors:     row[COL.CO_AUTHORS]     || '',
      abstractBody:  row[COL.ABSTRACT_BODY]  || '',
      approvalStatus: (row[COL.APPROVAL_STATUS] || '').toString().trim() || 'Unprocessed',
      reviewNotes:   row[COL.REVIEW_NOTES]   || '',
      posterNumber:  row[COL.POSTER_NUMBER]  || '',
      pdfLink:       row[COL.PDF_LINK]       || ''
    });
  }

  return abstracts;
}

// ===== SIDEBAR: SAVE APPROVAL DECISION =====
// Returns { posterNumber } so the sidebar can display the auto-assigned number immediately.
function saveAbstractDecision(rowIndex, status, notes) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  sheet.getRange(rowIndex, COL.APPROVAL_STATUS + 1).setValue(status || 'Unprocessed');
  sheet.getRange(rowIndex, COL.REVIEW_NOTES    + 1).setValue(notes  || '');

  const statusLower = (status || '').toLowerCase();
  const rowRange    = sheet.getRange(rowIndex, 1, 1, 22);
  let posterNumber  = '';

  if (statusLower === 'approved') {
    rowRange.setBackgroundColor('#D9F2D9');
    const existing = sheet.getRange(rowIndex, COL.POSTER_NUMBER + 1).getValue();
    if (existing && existing.toString().trim() !== '') {
      posterNumber = existing.toString().trim();
    } else {
      const next = getNextPosterNumber(sheet);
      posterNumber = next.toString().padStart(3, '0');
      sheet.getRange(rowIndex, COL.POSTER_NUMBER + 1).setValue(posterNumber);
    }
  } else if (statusLower === 'not approved') {
    rowRange.setBackgroundColor('#F2D9D9');
    sheet.getRange(rowIndex, COL.POSTER_NUMBER + 1).setValue('');
  } else {
    rowRange.setBackgroundColor('#FFFACD');
    sheet.getRange(rowIndex, COL.POSTER_NUMBER + 1).setValue('');
  }

  return { posterNumber: posterNumber };
}

// Converts "First Last" → "Last, First". Handles multi-part first names.
function formatAuthorLastFirst(fullName) {
  const name = (fullName || '').trim();
  if (!name) return name;
  const lastSpace = name.lastIndexOf(' ');
  if (lastSpace === -1) return name;
  return name.substring(lastSpace + 1) + ', ' + name.substring(0, lastSpace);
}

// ===== SIDEBAR: GENERATE ABSTRACT LIST PDF =====
// Creates a styled Google Doc with a numbered table, exports it as a PDF to the
// submissions folder, and returns the Drive view URL.
// Throws if any approved abstract is missing a poster number.
function generateAbstractListPDF() {
  if (!SUBMISSIONS_FOLDER_ID) throw new Error('SUBMISSIONS_FOLDER_ID not set in Script Properties');

  const abstracts = generateAbstractList(); // validates + sorts

  const ts       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const docTitle = "2026 FoMD SSRD – Abstract List " + ts;

  const doc  = DocumentApp.create(docTitle);
  const body = doc.getBody();
  body.clear();

  // Title block
  const heading = body.appendParagraph("2026 FoMD Summer Students’ Research Day");
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  heading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  heading.editAsText().setForegroundColor('#003C1E');

  const sub = body.appendParagraph('Abstract List');
  sub.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('');

  // Build table rows — poster number as plain integer (or '.' kept as-is)
  const rows = [['#', 'First Author', 'Abstract Title']];
  abstracts.forEach(function(a) {
    const num = parseInt(a.posterNumber, 10);
    rows.push([isNaN(num) ? a.posterNumber.trim() : num + '.', formatAuthorLastFirst(a.firstAuthor), a.title]);
  });

  const table = body.appendTable(rows);

  // Style header row
  const headerRow = table.getRow(0);
  for (let c = 0; c < 3; c++) {
    headerRow.getCell(c).setBackgroundColor('#003C1E');
    headerRow.getCell(c).editAsText().setForegroundColor('#ffffff').setBold(true);
  }

  // Alternate row shading
  for (let r = 1; r < table.getNumRows(); r++) {
    if (r % 2 === 0) {
      for (let c = 0; c < 3; c++) {
        table.getRow(r).getCell(c).setBackgroundColor('#e8f5e8');
      }
    }
  }

  doc.saveAndClose();

  const pdfBlob = DocumentApp.openById(doc.getId())
    .getAs('application/pdf')
    .setName(docTitle + '.pdf');
  const folder  = DriveApp.getFolderById(SUBMISSIONS_FOLDER_ID);
  const pdfFile = folder.createFile(pdfBlob);

  Drive.Files.trash(doc.getId(), { supportsAllDrives: true });

  return 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view';
}

// ===== SIDEBAR: VALIDATE & RETURN ABSTRACT LIST DATA =====
// Throws if any approved abstract is still missing a poster number.
function generateAbstractList() {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Form Responses');
  const data    = sheet.getDataRange().getValues();
  const list    = [];
  const missing = [];

  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    if (!row[COL.TIMESTAMP]) continue;
    const approval = (row[COL.APPROVAL_STATUS] || '').toString().trim().toLowerCase();
    if (approval !== 'approved') continue;

    const posterNumber = (row[COL.POSTER_NUMBER] || '').toString().trim();
    if (!posterNumber) {
      missing.push(
        (row[COL.STUDENT_FIRST] || '') + ' ' + (row[COL.STUDENT_LAST] || '')
      );
      continue;
    }

    list.push({
      posterNumber: posterNumber,
      firstAuthor:  (row[COL.FIRST_AUTHOR] || '').toString(),
      title:        (row[COL.TITLE]        || '').toString()
    });
  }

  if (missing.length > 0) {
    throw new Error(
      missing.length + ' approved abstract(s) have no poster number yet: ' +
      missing.join(', ') + '. Open each submission in the Abstract Review panel and re-save, or run "Fix Missing Poster Numbers" from the menu.'
    );
  }

  list.sort(function(a, b) {
    const rawA = a.posterNumber.trim();
    const rawB = b.posterNumber.trim();
    if (rawA === '.') return 1;   // '.' entries go to the end
    if (rawB === '.') return -1;
    const nA = parseInt((rawA.match(/\d+/) || ['0'])[0], 10);
    const nB = parseInt((rawB.match(/\d+/) || ['0'])[0], 10);
    return nA - nB;
  });

  return list;
}

// ===== UTILITIES =====
function ping() {
  return true;
}
