function runSaveQuotesV2() {
  const SHEET_NAME = "Save Quotes V2";
  const OUTPUT_LABEL = "Quote Saved";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${SHEET_NAME}" not found.`);
    return;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

  const allowedExt = new Set(["xlsx", "xls", "pdf", "doc", "docx", "zip"]);

  const outLabel = GmailApp.getUserLabelByName(OUTPUT_LABEL)
    || GmailApp.createLabel(OUTPUT_LABEL);

  let processed = 0;
  let failed = 0;

  data.forEach((row, i) => {
    const messageId = row[0];
    const folderRef = row[1];
    const status = row[2];

    // ✅ Skip anything already done
    if (!messageId || status === "DONE") return;

    try {
      // ---- mark as processing (optional but useful) ----
      sheet.getRange(i + 2, 3).setValue("PROCESSING");

      const folderId = extractDriveFolderId_(folderRef);
      const destFolder = DriveApp.getFolderById(folderId);

      // ---- fetch message ----
      const threads = GmailApp.search(`rfc822msgid:${messageId}`);
      if (!threads.length) throw new Error("Message not found");

      const thread = threads[0];
      const msg = thread.getMessages().find(m => m.getId() === messageId)
                || thread.getMessages().pop();

      const subject = sanitizeFilename_(msg.getSubject() || "No Subject");
      const from = msg.getFrom();
      const dateStr = Utilities.formatDate(
        msg.getDate(),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm"
      );

      const htmlBody = msg.getBody();

      // ---- PDF ----
      const html = `
<!doctype html>
<html>
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
        .setName(`${subject} - Email.pdf`);

      destFolder.createFile(pdfBlob);

      // ---- attachments ----
      msg.getAttachments({ includeInlineImages: false }).forEach(att => {
        const name = att.getName() || "attachment";
        const ext = name.split(".").pop().toLowerCase();

        if (allowedExt.has(ext)) {
          destFolder.createFile(att).setName(name);
        }
      });

      // ---- label thread (optional tracking) ----
      thread.addLabel(outLabel);

      // ---- mark DONE ----
      sheet.getRange(i + 2, 3).setValue("DONE");

      processed++;

    } catch (err) {
      sheet.getRange(i + 2, 3).setValue("ERROR");
      failed++;
    }
  });

  SpreadsheetApp.getUi().alert(
    `Save Quotes V2 Completed\n\nSuccess: ${processed}\nFailed: ${failed}`
  );
}
