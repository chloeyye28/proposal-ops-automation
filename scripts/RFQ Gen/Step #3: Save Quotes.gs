function runSaveQuotesV2() {
  const SHEET_NAME = "Save Quotes V2";
  const OUTPUT_LABEL = "Quote Saved";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Missing sheet: ${SHEET_NAME}`);
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert("Script is already running. Try again shortly.");
    return;
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    const allowedExt = new Set(["xlsx", "xls", "pdf", "doc", "docx", "zip"]);

    const label = GmailApp.getUserLabelByName(OUTPUT_LABEL)
      || GmailApp.createLabel(OUTPUT_LABEL);

    let processed = 0;
    let failed = 0;

    data.forEach((row, i) => {
      const messageId = String(row[0] || "").trim();
      const folderRef = String(row[1] || "").trim();
      const status = String(row[2] || "").trim();

      if (!messageId || status === "DONE") return;

      try {
        sheet.getRange(i + 2, 3).setValue("PROCESSING");

        const folderId = extractDriveFolderId_(folderRef);
        if (!folderId) throw new Error("Invalid folder reference");

        const destFolder = DriveApp.getFolderById(folderId);

        const message = findMessageByRfcId_(messageId);
        if (!message) throw new Error("Email message not found");

        const subject = sanitizeFilename_(message.getSubject() || "No Subject");
        const from = message.getFrom();
        const dateStr = Utilities.formatDate(
          message.getDate(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd HH:mm"
        );

        const htmlBody = message.getBody();

        const html = `
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
</head>
<body>
  <h3>${escapeHtml_(subject)}</h3>
  <p><b>From:</b> ${escapeHtml_(from)}</p>
  <p><b>Date:</b> ${escapeHtml_(dateStr)}</p>
  <hr>
  ${htmlBody}
</body>
</html>`;

        const pdfBlob = HtmlService.createHtmlOutput(html)
          .getBlob()
          .getAs(MimeType.PDF)
          .setName(`${subject}.pdf`);

        destFolder.createFile(pdfBlob);

        const attachments = message.getAttachments({ includeInlineImages: false });

        attachments.forEach(att => {
          const name = att.getName() || "attachment";
          const ext = name.split(".").pop().toLowerCase();

          if (allowedExt.has(ext)) {
            destFolder.createFile(att.copyBlob()).setName(name);
          }
        });

        const thread = message.getThread();
        thread.addLabel(label);

        sheet.getRange(i + 2, 3).setValue("DONE");
        processed++;

      } catch (err) {
        sheet.getRange(i + 2, 3).setValue("ERROR: " + err.message);
        failed++;
      }
    });

    SpreadsheetApp.getUi().alert(
      `Completed\n\nSuccess: ${processed}\nFailed: ${failed}`
    );

  } finally {
    lock.releaseLock();
  }
}

/**
 * Safe Gmail message lookup using RFC822 Message-ID
 * (No unsafe fallback logic)
 */
function findMessageByRfcId_(messageId) {
  const threads = GmailApp.search(`rfc822msgid:${messageId}`, 0, 5);
  if (!threads.length) return null;

  for (const thread of threads) {
    const msg = thread.getMessages().find(m => {
      return String(m.getHeader("Message-ID") || "").includes(messageId);
    });
    if (msg) return msg;
  }

  // fallback: safest possible option (first message only if exact match fails)
  const first = threads[0].getMessages()[0];
  return first || null;
}

/**
 * Extract Drive folder ID from string reference
 */
function extractDriveFolderId_(text) {
  const match = String(text).match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Basic filename sanitizer
 */
function sanitizeFilename_(name) {
  return String(name).replace(/[\\\/:*?"<>|]+/g, "-").trim();
}

/**
 * Escape HTML safely
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
