/**
 * Batch export Google Docs to PDF
 *
 * SHEET SETUP
 * -------------------------------------------------
 * Tab Name: Batch PDFing
 *
 * B1 = Destination folder URL
 * A7:A = Google Doc links
 * B7:B = Status output
 * C7:C = Generated PDF links
 */

function batchConvertFromSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Batch PDFing");

  if (!sheet) {
    throw new Error('Tab "Batch PDFing" not found.');
  }

  // Destination folder URL
  const destFolderUrl = String(sheet.getRange("B1").getValue()).trim();

  if (!destFolderUrl) {
    throw new Error("Please provide a destination folder URL in B1.");
  }

  const destFolderId = extractDriveId_(destFolderUrl);
  const destFolder = DriveApp.getFolderById(destFolderId);

  const lastRow = sheet.getLastRow();
  if (lastRow < 7) return;

  // Read document links
  const links = sheet.getRange(7, 1, lastRow - 6, 1).getValues();

  const statuses = [];
  const outputLinks = [];

  for (let i = 0; i < links.length; i++) {
    const link = String(links[i][0] || "").trim();

    if (!link) {
      statuses.push([""]);
      outputLinks.push([""]);
      continue;
    }

    try {
      const fileId = extractDriveId_(link);
      const file = DriveApp.getFileById(fileId);

      const pdfBlob = exportFileAsPdf_(fileId);

      const cleanName = sanitizeFilename_(file.getName());
      const pdfName = cleanName.endsWith(".pdf")
        ? cleanName
        : `${cleanName}.pdf`;

      const createdPdf = destFolder.createFile(pdfBlob).setName(pdfName);

      statuses.push(["SUCCESS"]);
      outputLinks.push([createdPdf.getUrl()]);

    } catch (err) {
      statuses.push([`FAILED - ${err.message || err}`]);
      outputLinks.push([""]);
    }
  }

  // Write results back to sheet
  sheet.getRange(7, 2, statuses.length, 1).setValues(statuses);
  sheet.getRange(7, 3, outputLinks.length, 1).setValues(outputLinks);
}
