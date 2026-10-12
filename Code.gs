
/*******************************************************
 * IFEXA ACADEMY
 * COMPLETE DATA ANALYSIS CERTIFICATE AUTOMATION
 *
 * FEATURES
 * - Generate certificates and certificate IDs
 * - Create PDFs with QR verification codes
 * - Regenerate certificates in batches
 * - Archive old certificates
 * - Track progress and retry failures
 * - Send certificates to students not yet marked SENT
 * - Resend successfully regenerated certificates
 * - Verify certificates through the web app
 *******************************************************/

const CONFIG = {
  TRACKER_SHEET: "Certificate Tracker",
  TEMPLATE_ID: "1dd0Fe80xdXj3td2WhMNUeZZI_W53eRZdYRnl8lB4D2M",
  CERTIFICATE_FOLDER_NAME: "IFEXA September 2026 Certificates",
  OLD_CERTIFICATE_FOLDER_NAME: "Old Certificates",
  COURSE_NAME: "Data Analysis",
  ORGANIZATION: "IFEXA Academy",
  CERTIFICATE_PREFIX: "IFEXA-DA-CERT-",
  VERIFICATION_WEB_APP_URL:
    "https://script.google.com/macros/s/AKfycbyLR6ODt9Bji_UZKM8eqIPPmzvWUbpXItTaD36wr9ucITrzKTuuE3atlhAu-TAxbGZJaw/exec",
  QR_API: "https://quickchart.io/qr",
  QR_SIZE: 500,
  BATCH_SIZE: 15
};


/*******************************************************
 * CUSTOM MENU
 *******************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("IFEXA Certificates")
    .addItem("Generate Certificates", "generateCertificates")
    .addItem("Prepare Full Regeneration", "prepareFullRegeneration")
    .addItem("Regenerate Certificates", "regenerateExistingCertificatesInBatches")
    .addItem("Check Regeneration Progress", "checkQRRegenerationProgress")
    .addItem("Retry Failed Certificates", "retryFailedCertificates")
    .addSeparator()
    .addItem("Send Current Certificates", "sendAllCurrentCertificates")
    .addItem("RESEND All Regenerated Certificates", "resendAllRegeneratedCertificates")
    .addSeparator()
    .addItem("Test Certificate", "testCreateCertificate")
    .addItem("Test Verification", "testVerification")
    .addToUi();
}


/*******************************************************
 * TRACKER SHEET
 *******************************************************/

function getTrackerSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!sheet) {
    throw new Error('Sheet "' + CONFIG.TRACKER_SHEET + '" was not found.');
  }

  return sheet;
}


/*******************************************************
 * CERTIFICATE FOLDERS
 *******************************************************/

function getCertificateFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.CERTIFICATE_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(CONFIG.CERTIFICATE_FOLDER_NAME);
}


function getOldCertificatesFolder() {
  const parentFolder = getCertificateFolder();
  const folders = parentFolder.getFoldersByName(CONFIG.OLD_CERTIFICATE_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(CONFIG.OLD_CERTIFICATE_FOLDER_NAME);
}


/*******************************************************
 * FILE HELPERS
 *******************************************************/

function extractFileId(urlOrId) {
  if (!urlOrId) return null;

  const value = String(urlOrId).trim();

  if (value.indexOf("/") === -1 && value.indexOf("http") !== 0) {
    return value;
  }

  const match = value.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}


function archiveOldCertificate(certificateLink) {
  if (!certificateLink) return;

  try {
    const fileId = extractFileId(certificateLink);
    if (!fileId) return;

    const file = DriveApp.getFileById(fileId);
    file.moveTo(getOldCertificatesFolder());

  } catch (error) {
    Logger.log("Could not archive old certificate: " + error.message);
  }
}


/*******************************************************
 * CERTIFICATE IDS
 *******************************************************/

function getNextCertificateNumber() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return 1;

  const values = sheet.getRange(2, 10, lastRow - 1, 1).getValues();
  let highest = 0;

  values.forEach(function(row) {
    if (!row[0]) return;

    const match = String(row[0]).match(/(\d+)$/);

    if (match) {
      highest = Math.max(highest, parseInt(match[1], 10));
    }
  });

  return highest + 1;
}


function makeCertificateId(number) {
  return CONFIG.CERTIFICATE_PREFIX + String(number).padStart(4, "0");
}


function generateStudentId(rowNumber) {
  return "IFEXA-DA-" + String(rowNumber - 1).padStart(4, "0");
}


/*******************************************************
 * ISSUE DATE
 *******************************************************/

function formatIssueDate(issueDate) {
  if (!issueDate) {
    throw new Error("Issue Date is missing.");
  }

  const date = new Date(issueDate);

  if (isNaN(date.getTime())) {
    throw new Error("Invalid Issue Date: " + issueDate);
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "dd MMMM yyyy"
  );
}


/*******************************************************
 * CREATE CERTIFICATE PDF
 *******************************************************/

function createCertificate(fullName, studentId, certificateId, issueDate) {
  const folder = getCertificateFolder();
  const formattedDate = formatIssueDate(issueDate);

  const verificationUrl =
    CONFIG.VERIFICATION_WEB_APP_URL +
    "?student=" +
    encodeURIComponent(studentId);

  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID);

  const presentationFile = templateFile.makeCopy(
    fullName + " - " + certificateId,
    folder
  );

  const presentationId = presentationFile.getId();
  const presentation = SlidesApp.openById(presentationId);

  try {
    presentation.replaceAllText("{{FULL_NAME}}", String(fullName));
    presentation.replaceAllText("{{STUDENT_ID}}", String(studentId));
    presentation.replaceAllText("{{CERTIFICATE_ID}}", String(certificateId));
    presentation.replaceAllText("{{COURSE}}", CONFIG.COURSE_NAME);
    presentation.replaceAllText("{{ISSUE_DATE}}", formattedDate);

    insertQRCode(presentation, verificationUrl);
    presentation.saveAndClose();

    Utilities.sleep(1500);

    const pdfBlob = exportPresentationAsPDF(presentationId);
    pdfBlob.setName(fullName + " - " + certificateId + ".pdf");

    return folder.createFile(pdfBlob);

  } finally {
    // The temporary Slides copy is not the final certificate.
    presentationFile.setTrashed(true);
  }
}


/*******************************************************
 * INSERT QR CODE INTO TEMPLATE
 *******************************************************/

function insertQRCode(presentation, verificationUrl) {
  const qrUrl =
    CONFIG.QR_API +
    "?text=" + encodeURIComponent(verificationUrl) +
    "&size=" + CONFIG.QR_SIZE +
    "&margin=2&ecLevel=H";

  const response = UrlFetchApp.fetch(qrUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      "QR code generation failed. HTTP status: " +
      response.getResponseCode()
    );
  }

  const qrBlob = response.getBlob().setName("certificate-qr.png");
  const slides = presentation.getSlides();
  let qrInserted = false;

  for (let s = 0; s < slides.length && !qrInserted; s++) {
    const slide = slides[s];
    const elements = slide.getPageElements();

    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];

      if (
        element.getPageElementType() !==
        SlidesApp.PageElementType.SHAPE
      ) {
        continue;
      }

      const shape = element.asShape();
      const text = shape.getText().asString().trim();

      if (text !== "{{QR_CODE}}") continue;

      const left = element.getLeft();
      const top = element.getTop();
      const width = element.getWidth();
      const height = element.getHeight();

      element.remove();
      slide.insertImage(qrBlob, left, top, width, height);

      qrInserted = true;
      break;
    }
  }

  if (!qrInserted) {
    throw new Error(
      "QR_CODE placeholder was not found in the certificate template."
    );
  }
}


/*******************************************************
 * EXPORT SLIDES PRESENTATION AS PDF
 *******************************************************/

function exportPresentationAsPDF(presentationId) {
  const exportUrl =
    "https://docs.google.com/presentation/d/" +
    presentationId +
    "/export/pdf";

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      "PDF export failed. HTTP status: " +
      response.getResponseCode()
    );
  }

  return response.getBlob();
}


/*******************************************************
 * PREPARE FULL REGENERATION
 *******************************************************/

function prepareFullRegeneration() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("No certificate records found.");
  }

  sheet.getRange(1, 13).setValue("QR Status");

  // Clears regeneration statuses, not certificate IDs or PDF links.
  sheet.getRange(2, 13, lastRow - 1, 1).clearContent();

  SpreadsheetApp.flush();

  Logger.log("Full regeneration prepared.");
  Logger.log("Use Regenerate Certificates to process batches.");
}


/*******************************************************
 * REGENERATE CERTIFICATES IN BATCHES
 *******************************************************/

function regenerateExistingCertificatesInBatches() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No certificate records found.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  let processed = 0;
  let completed = 0;
  let errors = 0;

  for (let i = 0; i < data.length; i++) {
    if (processed >= CONFIG.BATCH_SIZE) break;

    const rowNumber = i + 2;
    const studentId = data[i][0];
    const fullName = data[i][1];
    const certificateId = data[i][9];
    const issueDate = data[i][10];
    const oldCertificateLink = data[i][11];
    const qrStatus = String(data[i][12] || "").toUpperCase();

    if (qrStatus === "DONE") continue;

    if (!studentId || !fullName || !certificateId || !issueDate) {
      const missing = !studentId ? "Student ID" :
        !fullName ? "Name" :
        !certificateId ? "Certificate ID" : "Issue Date";

      sheet.getRange(rowNumber, 13).setValue("ERROR: Missing " + missing);
      errors++;
      processed++;
      continue;
    }

    try {
      sheet.getRange(rowNumber, 13).setValue("PROCESSING");
      SpreadsheetApp.flush();

      const newPdf = createCertificate(
        fullName,
        studentId,
        certificateId,
        issueDate
      );

      const newPdfUrl = newPdf.getUrl();

      // Save the new PDF link before archiving the old one.
      sheet.getRange(rowNumber, 12).setValue(newPdfUrl);
      sheet.getRange(rowNumber, 13).setValue("DONE");
      SpreadsheetApp.flush();

      if (oldCertificateLink && oldCertificateLink !== newPdfUrl) {
        archiveOldCertificate(oldCertificateLink);
      }

      Logger.log("REGENERATED: " + fullName + " | " + certificateId);
      completed++;

    } catch (error) {
      sheet.getRange(rowNumber, 13).setValue("ERROR: " + error.message);
      Logger.log(
        "REGENERATION ERROR: " + fullName +
        " | " + certificateId +
        " | " + error.message
      );
      errors++;
    }

    processed++;
  }

  Logger.log("BATCH COMPLETE");
  Logger.log("Processed: " + processed);
  Logger.log("Completed: " + completed);
  Logger.log("Errors: " + errors);
}


/*******************************************************
 * CHECK REGENERATION PROGRESS
 *******************************************************/

function checkQRRegenerationProgress() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No certificate records found.");
    return;
  }

  const statuses = sheet.getRange(2, 13, lastRow - 1, 1).getValues();

  let done = 0;
  let remaining = 0;
  let errors = 0;
  let processing = 0;

  statuses.forEach(function(row) {
    const status = String(row[0] || "").toUpperCase();

    if (status === "DONE") {
      done++;
    } else if (status === "PROCESSING") {
      processing++;
    } else if (status.indexOf("ERROR") === 0) {
      errors++;
    } else {
      remaining++;
    }
  });

  const total = statuses.length;

  Logger.log("CERTIFICATE REGENERATION PROGRESS");
  Logger.log("Total: " + total);
  Logger.log("Completed: " + done);
  Logger.log("Processing: " + processing);
  Logger.log("Remaining: " + remaining);
  Logger.log("Errors: " + errors);
  Logger.log(
    "Completion: " +
    (total ? (done / total * 100).toFixed(1) : "0") +
    "%"
  );
}


/*******************************************************
 * RETRY FAILED CERTIFICATES
 *******************************************************/

function retryFailedCertificates() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const statuses = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  let reset = 0;

  statuses.forEach(function(row, index) {
    const status = String(row[0] || "").toUpperCase();

    if (status.indexOf("ERROR") === 0) {
      sheet.getRange(index + 2, 13).clearContent();
      reset++;
    }
  });

  Logger.log("Reset " + reset + " failed certificates.");
}


/*******************************************************
 * GENERATE NEW CERTIFICATES
 *
 * Column A: Student ID
 * Column B: Full Name
 * Column C: Email
 * Column I: Certificate Status
 * Column J: Certificate ID
 * Column K: Issue Date
 * Column L: Certificate PDF Link
 * Column M: QR Status
 *******************************************************/

function generateCertificates() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("No students found.");
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  let nextNumber = getNextCertificateNumber();
  let generated = 0;
  let errors = 0;

  for (let i = 0; i < data.length; i++) {
    const rowNumber = i + 2;
    const studentId = data[i][0];
    const fullName = data[i][1];
    const status = String(data[i][8] || "").toUpperCase();
    let certificateId = data[i][9];
    let issueDate = data[i][10];

    if (status === "GENERATED" || status === "SENT") continue;
    if (!studentId || !fullName) continue;

    if (!certificateId) {
      certificateId = makeCertificateId(nextNumber);
      nextNumber++;
    }

    if (!issueDate) {
      issueDate = new Date();
    }

    try {
      const pdfFile = createCertificate(
        fullName,
        studentId,
        certificateId,
        issueDate
      );

      sheet.getRange(rowNumber, 10).setValue(certificateId);
      sheet.getRange(rowNumber, 11).setValue(issueDate);
      sheet.getRange(rowNumber, 12).setValue(pdfFile.getUrl());
      sheet.getRange(rowNumber, 9).setValue("GENERATED");
      sheet.getRange(rowNumber, 13).setValue("DONE");

      generated++;

      Logger.log("GENERATED: " + fullName + " | " + certificateId);

    } catch (error) {
      errors++;
      Logger.log(
        "GENERATION ERROR: " + fullName + " | " + error.message
      );
    }
  }

  Logger.log("Certificates generated: " + generated);
  Logger.log("Generation errors: " + errors);
}


/*******************************************************
 * SEND CERTIFICATES NOT YET MARKED SENT
 *
 * Sends only rows with Status = GENERATED.
 * Rows marked SENT are skipped to prevent accidental
 * repeat sending.
 *
 * Column C = Email
 * Column I = Certificate Status
 * Column J = Certificate ID
 * Column L = PDF link
 * Column N = Last-sent timestamp
 *******************************************************/

function sendAllCurrentCertificates() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No student records found.");
    return;
  }

  sheet.getRange(1, 14).setValue("Certificate Last Sent");

  let quota = MailApp.getRemainingDailyQuota();

  Logger.log("Available email quota: " + quota);

  if (quota <= 0) {
    Logger.log("No email quota available. No certificates were sent.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  let sent = 0;
  let skipped = 0;
  let errors = 0;

  for (let i = 0; i < data.length; i++) {
    if (quota <= 0) {
      Logger.log("Email quota exhausted. Remaining rows were not processed.");
      break;
    }

    const rowNumber = i + 2;
    const studentId = data[i][0];
    const fullName = data[i][1];
    const email = String(data[i][2] || "").trim();
    const status = String(data[i][8] || "").trim().toUpperCase();
    const certificateId = data[i][9];
    const certificateLink = data[i][11];

    if (status !== "GENERATED") {
      skipped++;
      continue;
    }

    if (!studentId || !fullName || !email || !certificateLink) {
      Logger.log("SKIPPED row " + rowNumber + ": Missing required information.");
      skipped++;
      continue;
    }

    try {
      const fileId = extractFileId(certificateLink);

      if (!fileId) {
        throw new Error("Invalid certificate link.");
      }

      const file = DriveApp.getFileById(fileId);

      if (file.getMimeType() !== "application/pdf") {
        throw new Error("Certificate file is not a PDF.");
      }

      MailApp.sendEmail({
        to: email,
        subject: "Your IFEXA Academy Data Analysis Certificate",
        body:
          "Hello " + fullName + ",\n\n" +
          "Congratulations!\n\n" +
          "Your IFEXA Academy Data Analysis Certificate is attached to this email.\n\n" +
          "Student ID: " + studentId + "\n" +
          "Certificate ID: " + certificateId + "\n\n" +
          "Thank you for being part of IFEXA Academy's Data Analysis training.\n\n" +
          "Best regards,\nIFEXA Academy",
        attachments: [file.getBlob()]
      });

      // Only mark SENT after MailApp accepts the request.
      sheet.getRange(rowNumber, 9).setValue("SENT");
      sheet.getRange(rowNumber, 14).setValue(new Date());

      quota--;
      sent++;

      Logger.log("SENT: " + fullName + " | " + email);

    } catch (error) {
      errors++;
      Logger.log(
        "SEND ERROR: " + fullName +
        " | " + email +
        " | " + error.message
      );
    }
  }

  SpreadsheetApp.flush();

  Logger.log("CERTIFICATE SENDING SUMMARY");
  Logger.log("Successfully sent: " + sent);
  Logger.log("Skipped: " + skipped);
  Logger.log("Errors: " + errors);
  Logger.log("Estimated quota remaining: " + quota);
}


/*******************************************************
 * RESEND ALL SUCCESSFULLY REGENERATED CERTIFICATES
 *
 * Sends the current PDF in Column L to students whose
 * Column M regeneration status is DONE.
 *
 * This intentionally resends to students already marked
 * SENT, because the regenerated certificate may be new.
 *******************************************************/

function resendAllRegeneratedCertificates() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No students found.");
    return;
  }

  sheet.getRange(1, 14).setValue("Certificate Last Sent");

  let quota = MailApp.getRemainingDailyQuota();

  Logger.log("Available email quota: " + quota);

  if (quota <= 0) {
    Logger.log("No email quota available. No certificates were resent.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  let sent = 0;
  let skipped = 0;
  let errors = 0;

  for (let i = 0; i < data.length; i++) {
    if (quota <= 0) {
      Logger.log("Email quota exhausted. Remaining rows were not processed.");
      break;
    }

    const rowNumber = i + 2;
    const studentId = data[i][0];
    const fullName = data[i][1];
    const email = String(data[i][2] || "").trim();
    const status = String(data[i][8] || "").trim().toUpperCase();
    const certificateId = data[i][9];
    const certificateLink = data[i][11];
    const qrStatus = String(data[i][12] || "").trim().toUpperCase();

    if (qrStatus !== "DONE") {
      Logger.log("SKIPPED: " + fullName + " | Regeneration is not DONE.");
      skipped++;
      continue;
    }

    if (!studentId || !fullName || !email || !certificateLink) {
      Logger.log("SKIPPED: " + fullName + " | Missing required information.");
      skipped++;
      continue;
    }

    if (status !== "GENERATED" && status !== "SENT") {
      Logger.log("SKIPPED: " + fullName + " | Certificate status is " + status);
      skipped++;
      continue;
    }

    try {
      const fileId = extractFileId(certificateLink);

      if (!fileId) {
        throw new Error("Invalid certificate link.");
      }

      const file = DriveApp.getFileById(fileId);

      if (file.getMimeType() !== "application/pdf") {
        throw new Error("Certificate file is not a PDF.");
      }

      MailApp.sendEmail({
        to: email,
        subject: "Your Updated IFEXA Academy Data Analysis Certificate",
        body:
          "Hello " + fullName + ",\n\n" +
          "Congratulations!\n\n" +
          "Your updated IFEXA Academy Data Analysis Certificate is attached.\n\n" +
          "Student ID: " + studentId + "\n" +
          "Certificate ID: " + certificateId + "\n\n" +
          "This is the latest version of your certificate and includes a verification QR code.\n\n" +
          "Thank you for being part of IFEXA Academy.\n\n" +
          "Best regards,\nIFEXA Academy",
        attachments: [file.getBlob()]
      });

      sheet.getRange(rowNumber, 9).setValue("SENT");
      sheet.getRange(rowNumber, 14).setValue(new Date());

      quota--;
      sent++;

      Logger.log("RESENT: " + fullName + " | " + email);

    } catch (error) {
      errors++;
      Logger.log(
        "RESEND ERROR: " + fullName +
        " | " + email +
        " | " + error.message
      );
    }
  }

  SpreadsheetApp.flush();

  Logger.log("REGENERATED CERTIFICATE RESEND SUMMARY");
  Logger.log("Successfully sent: " + sent);
  Logger.log("Skipped: " + skipped);
  Logger.log("Errors: " + errors);
  Logger.log("Estimated quota remaining: " + quota);
}


/*******************************************************
 * VERIFICATION WEB APP
 *
 * Keep your existing Verification.html file in the
 * same Apps Script project.
 *******************************************************/

function doGet(e) {
  return HtmlService
    .createTemplateFromFile("Verification")
    .evaluate()
    .setTitle("IFEXA Certificate Verification")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/*******************************************************
 * VERIFY CERTIFICATE USING STUDENT ID
 *******************************************************/

function verifyCertificate(studentId) {
  if (!studentId) {
    return {
      success: false,
      verified: false,
      message: "Please enter a Student ID."
    };
  }

  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: false,
      verified: false,
      message: "No certificate records found."
    };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const searchId = String(studentId).trim().toUpperCase();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowStudentId = String(row[0] || "").trim().toUpperCase();

    if (rowStudentId !== searchId) continue;

    const status = String(row[8] || "").trim().toUpperCase();

    if (status !== "GENERATED" && status !== "SENT") {
      return {
        success: false,
        verified: false,
        message: "This certificate has not been issued."
      };
    }

    try {
      return {
        success: true,
        verified: true,
        fullName: row[1],
        studentId: rowStudentId,
        certificateId: row[9],
        course: CONFIG.COURSE_NAME,
        issueDate: formatIssueDate(row[10]),
        certificateLink: row[11],
        organization: CONFIG.ORGANIZATION
      };
    } catch (error) {
      return {
        success: false,
        verified: false,
        message: "Certificate record has an invalid issue date."
      };
    }
  }

  return {
    success: false,
    verified: false,
    message: "Certificate not found. Please check the Student ID and try again."
  };
}


/*******************************************************
 * TEST CERTIFICATE CREATION
 *******************************************************/

function testCreateCertificate() {
  const sheet = getTrackerSheet();
  const row = 2;

  const data = sheet.getRange(row, 1, 1, 14).getValues()[0];

  const studentId = data[0];
  const fullName = data[1];
  const issueDate = data[10];
  const certificateId =
    data[9] || makeCertificateId(getNextCertificateNumber());

  if (!studentId || !fullName || !issueDate) {
    throw new Error(
      "Row 2 must contain a Student ID, Full Name, and Issue Date."
    );
  }

  const pdf = createCertificate(
    fullName,
    studentId,
    certificateId,
    issueDate
  );

  Logger.log("TEST CERTIFICATE CREATED: " + pdf.getUrl());
}


/*******************************************************
 * TEST VERIFICATION
 *******************************************************/

function testVerification() {
  const sheet = getTrackerSheet();
  const studentId = sheet.getRange(2, 1).getValue();

  const result = verifyCertificate(studentId);

  Logger.log(JSON.stringify(result, null, 2));
}

function generateAndSendLastSixCertificates() {
  const sheet = getTrackerSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("No student records found.");
  }

  const targetIds = [
    "IFEXA-DA-0376",
    "IFEXA-DA-0501",
    "IFEXA-DA-0789",
    "IFEXA-DA-0081",
    "IFEXA-DA-0476",
    "IFEXA-DA-0064"
  ];

  // Column O tracks this special batch.
  sheet.getRange(1, 15).setValue("Selected Batch Status");

  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

  let nextNumber = getNextCertificateNumber();
  let generated = 0;
  let sent = 0;
  let skipped = 0;
  let errors = 0;

  const foundIds = [];

  for (let i = 0; i < data.length; i++) {
    const rowNumber = i + 2;

    const studentId = String(data[i][0] || "").trim();

    if (targetIds.indexOf(studentId) === -1) {
      continue;
    }

    foundIds.push(studentId);

    const fullName = String(data[i][1] || "").trim();
    const email = String(data[i][2] || "").trim();

    let certificateId = String(data[i][9] || "").trim();
    let issueDate = data[i][10];
    let pdfLink = String(data[i][11] || "").trim();

    const batchStatus = String(data[i][14] || "").trim();

    if (batchStatus === "TARGETED_SENT") {
      Logger.log("Already completed: " + studentId);
      skipped++;
      continue;
    }

    try {
      if (!fullName || !email) {
        throw new Error("Missing student name or email.");
      }

      if (!certificateId) {
        certificateId = makeCertificateId(nextNumber);
        nextNumber++;
      }

      if (!issueDate) {
        issueDate = new Date();
      }

      // Create a new PDF unless one is already prepared
      // and waiting to be emailed by this function.
      if (batchStatus !== "TARGETED_READY" || !pdfLink) {
        const pdfFile = createCertificate(
          fullName,
          studentId,
          certificateId,
          issueDate
        );

        pdfLink = pdfFile.getUrl();

        sheet.getRange(rowNumber, 10).setValue(certificateId);
        sheet.getRange(rowNumber, 11).setValue(issueDate);
        sheet.getRange(rowNumber, 12).setValue(pdfLink);
        sheet.getRange(rowNumber, 13).setValue("DONE");
        sheet.getRange(rowNumber, 15).setValue("TARGETED_READY");

        SpreadsheetApp.flush();
        generated++;

        Logger.log("Certificate generated: " + studentId);
      }

      // Check email quota before sending.
      if (MailApp.getRemainingDailyQuota() <= 0) {
        Logger.log("Email quota exhausted. Pending: " + studentId);
        continue;
      }

      const fileId = extractFileId(pdfLink);

      if (!fileId) {
        throw new Error("Invalid certificate PDF link.");
      }

      const pdf = DriveApp.getFileById(fileId);

      if (pdf.getMimeType() !== "application/pdf") {
        throw new Error("Certificate file is not a PDF.");
      }

      MailApp.sendEmail({
        to: email,
        subject: "Your IFEXA Academy Data Analysis Certificate",
        body:
          "Hello " + fullName + ",\n\n" +
          "Congratulations on completing the IFEXA Academy Data Analysis Training.\n\n" +
          "Your certificate is attached to this email.\n\n" +
          "Student ID: " + studentId + "\n" +
          "Certificate ID: " + certificateId + "\n\n" +
          "Thank you for learning with IFEXA Academy.\n\n" +
          "Best regards,\n" +
          "IFEXA Academy",
        attachments: [pdf.getBlob()],
        name: "IFEXA Academy"
      });

      sheet.getRange(rowNumber, 9).setValue("SENT");
      sheet.getRange(rowNumber, 14).setValue(new Date());
      sheet.getRange(rowNumber, 15).setValue("TARGETED_SENT");

      SpreadsheetApp.flush();

      sent++;

      Logger.log("Email accepted for sending: " + email);

    } catch (error) {
      errors++;

      Logger.log(
        "ERROR for " + studentId + ": " + error.message
      );
    }
  }

  for (let i = 0; i < targetIds.length; i++) {
    if (foundIds.indexOf(targetIds[i]) === -1) {
      Logger.log("Student ID not found: " + targetIds[i]);
    }
  }

  Logger.log("========== SUMMARY ==========");
  Logger.log("New certificates generated: " + generated);
  Logger.log("Emails accepted for sending: " + sent);
  Logger.log("Already completed: " + skipped);
  Logger.log("Errors: " + errors);
}


/**
 * IFEXA ACADEMY — MISSING STUDENT CERTIFICATE BATCH
 *
 * Registration tab: Form Responses 1
 * Tracker tab: Certificate Tracker
 *
 * Adds only the four missing students, avoids duplicate Student IDs,
 * and calls the existing certificate generation/email function.
 *
 * Existing functions required:
 *   getTrackerSheet()
 *   generateAndSendLastSixCertificates()
 */

function appendMissingFourStudentsAndSendCertificates() {
  const REGISTRATION_SHEET_NAME = "Form Responses 1";

  const TARGET_STUDENT_IDS = [
    "IFEXA-DA-0789",
    "IFEXA-DA-0081",
    "IFEXA-DA-0476",
    "IFEXA-DA-0064"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registrationSheet = ss.getSheetByName(
    REGISTRATION_SHEET_NAME
  );

  if (!registrationSheet) {
    throw new Error(
      'Cannot find the registration tab "' +
      REGISTRATION_SHEET_NAME + '".'
    );
  }

  const trackerSheet = getTrackerSheet();

  // Read the registration records.
  const registrationData =
    registrationSheet.getDataRange().getValues();

  if (registrationData.length < 2) {
    throw new Error("No registration records were found.");
  }

  // Identify the required columns using the header row.
  function normalizeHeader(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  }

  const headers = registrationData[0].map(normalizeHeader);

  const nameColumn = headers.findIndex(function(header) {
    return header === "fullname" ||
           header === "name" ||
           header === "studentname";
  });

  const idColumn = headers.findIndex(function(header) {
    return header === "studentid" ||
           header === "studentidnumber";
  });

  const emailColumn = headers.findIndex(function(header) {
    return header === "emailaddress" ||
           header === "email" ||
           header === "studentemail";
  });

  if (
    nameColumn === -1 ||
    idColumn === -1 ||
    emailColumn === -1
  ) {
    throw new Error(
      "Required headers not found. Expected Full Name, " +
      "Student ID, and Email Address in the registration tab."
    );
  }

  // Create a lookup of the four target students.
  const registrationLookup = {};

  for (let i = 1; i < registrationData.length; i++) {
    const row = registrationData[i];
    const studentId = String(row[idColumn] || "").trim();

    if (
      TARGET_STUDENT_IDS.indexOf(studentId) !== -1 &&
      !registrationLookup[studentId]
    ) {
      registrationLookup[studentId] = {
        fullName: String(row[nameColumn] || "").trim(),
        email: String(row[emailColumn] || "").trim()
      };
    }
  }

  // Read existing Student IDs from tracker column A.
  const trackerLastRow = trackerSheet.getLastRow();
  const existingIds = new Set();

  if (trackerLastRow >= 2) {
    const existingIdValues = trackerSheet
      .getRange(2, 1, trackerLastRow - 1, 1)
      .getDisplayValues();

    existingIdValues.forEach(function(row) {
      const id = String(row[0] || "").trim();

      if (id) {
        existingIds.add(id);
      }
    });
  }

  let added = 0;
  let alreadyPresent = 0;
  let notFound = 0;
  let missingDetails = 0;

  TARGET_STUDENT_IDS.forEach(function(studentId) {
    // Never append a Student ID already in the tracker.
    if (existingIds.has(studentId)) {
      Logger.log(
        "Already in tracker; not duplicated: " + studentId
      );
      alreadyPresent++;
      return;
    }

    const student = registrationLookup[studentId];

    if (!student) {
      Logger.log(
        "Student ID not found in registration tab: " + studentId
      );
      notFound++;
      return;
    }

    if (!student.fullName || !student.email) {
      Logger.log(
        "Missing full name or email; not added: " + studentId
      );
      missingDetails++;
      return;
    }

    // Append a new tracker row.
    const newRow = trackerSheet.getLastRow() + 1;

    // A = Student ID
    // B = Full Name
    // C = Email
    trackerSheet.getRange(newRow, 1, 1, 3).setValues([[
      studentId,
      student.fullName,
      student.email
    ]]);

    // O = Selected Batch Status.
    // Keep this row pending so the existing batch function
    // can generate and send its certificate.
    trackerSheet.getRange(newRow, 15)
      .setValue("TARGETED_PENDING");

    existingIds.add(studentId);
    added++;

    Logger.log(
      "Added to Certificate Tracker: " +
      student.fullName + " (" + studentId + ")"
    );
  });

  SpreadsheetApp.flush();

  Logger.log("========== APPEND SUMMARY ==========");
  Logger.log("New students added: " + added);
  Logger.log("Already in tracker: " + alreadyPresent);
  Logger.log("Missing from registration: " + notFound);
  Logger.log("Missing name/email: " + missingDetails);

  // Stop if the registration data needs correction.
  if (notFound > 0 || missingDetails > 0) {
    throw new Error(
      "Some students could not be added. Review the execution log " +
      "before continuing with certificate generation."
    );
  }

  // Use your existing, tested certificate function.
  // It handles the six target IDs and skips TARGETED_SENT records.
  generateAndSendLastSixCertificates();

  Logger.log(
    "Finished processing the selected certificate batch. " +
    "Check the tracker and execution log for final status."
  );
}
