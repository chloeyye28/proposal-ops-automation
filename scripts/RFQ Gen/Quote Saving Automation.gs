function saveQuoteAttachments() {

  // -----------------------------------
  // Configuration
  // -----------------------------------

  const SHEET_NAME = "Quote Processing";
  const OUTPUT_LABEL = "Processed Quotes";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {

    SpreadsheetApp
      .getUi()
      .alert(`Sheet "${SHEET_NAME}" not found.`);

    return;
  }

  // -----------------------------------
  // Read sheet data
  // -----------------------------------

  const data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 3)
    .getValues();

  // Allowed attachment types

  const allowedExtensions = new Set([
    "xlsx",
    "xls",
    "pdf",
    "doc",
    "docx",
    "zip"
  ]);

  // Create/reuse Gmail label

  const processedLabel =
    GmailApp.getUserLabelByName(OUTPUT_LABEL)
    || GmailApp.createLabel(OUTPUT_LABEL);

  let processed = 0;
  let failed = 0;

  // -----------------------------------
  // Process each email row
  // -----------------------------------

  data.forEach((row, index) => {

    const messageId = row[0];
    const folderReference = row[1];
    const status = row[2];

    // Skip completed rows

    if (!messageId || status === "DONE") return;

    try {

      // Mark processing state

      sheet
        .getRange(index + 2, 3)
        .setValue("PROCESSING");

      // -----------------------------------
      // Resolve destination folder
      // -----------------------------------

      const folderId =
        extractDriveFolderId_(folderReference);

      const destinationFolder =
        DriveApp.getFolderById(folderId);

      // -----------------------------------
      // Locate Gmail message
      // -----------------------------------

      const threads =
        GmailApp.search(`rfc822msgid:${messageId}`);

      if (!threads.length) {
        throw new Error("Email message not found");
      }

      const thread = threads[0];

      const message =
        thread
          .getMessages()
          .find(m => m.getId() === messageId)
        || thread.getMessages().pop();

      // -----------------------------------
      // Build metadata
      // -----------------------------------

      const subject =
        sanitizeFilename_(
          message.getSubject() || "No Subject"
        );

      const sender =
        message.getFrom();

      const formattedDate =
        Utilities.formatDate(
          message.getDate(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd HH:mm"
        );

      const htmlBody =
        message.getBody();

      // -----------------------------------
      // Generate PDF copy of email
      // -----------------------------------

      const html = `
<!doctype html>
<html>
<body>
  <h3>${escapeHtml_(subject)}</h3>

  <p>
    <b>From:</b>
    ${escapeHtml_(sender)}
  </p>

  <p>
    <b>Date:</b>
    ${escapeHtml_(formattedDate)}
  </p>

  <hr>

  ${htmlBody}

</body>
</html>`;

      const pdfBlob =
        HtmlService
          .createHtmlOutput(html)
          .getBlob()
          .getAs(MimeType.PDF)
          .setName(`${subject} - Email.pdf`);

      destinationFolder.createFile(pdfBlob);

      // -----------------------------------
      // Save allowed attachments
      // -----------------------------------

      message
        .getAttachments({
          includeInlineImages: false
        })
        .forEach(attachment => {

          const fileName =
            attachment.getName() || "attachment";

          const extension =
            fileName
              .split(".")
              .pop()
              .toLowerCase();

          if (allowedExtensions.has(extension)) {

            destinationFolder
              .createFile(attachment)
              .setName(fileName);
          }
        });

      // -----------------------------------
      // Apply Gmail label
      // -----------------------------------

      thread.addLabel(processedLabel);

      // -----------------------------------
      // Mark complete
      // -----------------------------------

      sheet
        .getRange(index + 2, 3)
        .setValue("DONE");

      processed++;

    } catch (error) {

      sheet
        .getRange(index + 2, 3)
        .setValue("ERROR");

      failed++;
    }
  });

  // -----------------------------------
  // Completion summary
  // -----------------------------------

  SpreadsheetApp
    .getUi()
    .alert(
      `Quote Processing Complete\n\nSuccessful: ${processed}\nFailed: ${failed}`
    );
}
