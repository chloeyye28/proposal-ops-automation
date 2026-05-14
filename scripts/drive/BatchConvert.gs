/**
 * Batch export Google Docs to PDF
 *
 * SHEET SETUP
 * -------------------------------------------------
 * Tab Name: Batch PDFing
 *
 * B1 = Destination folder URL
 * A7:A = Google Doc links
 * B7:B = Status feedback
 * C7:C = PDF link output
 */

function batchConvertFromSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Batch PDFing");

  if (!sheet) {
    throw new Error('Tab "Batch PDFing" not found.');
  }

  // Destination folder
  const destFolderUrl = String(sheet.getRange("B1").getValue()).trim();

  if (!destFolderUrl) {
    throw new Error("Please paste destination folder URL in B1.");
  }

  const destFolderId = extractDriveId_(destFolderUrl);
  const destFolder = DriveApp.getFolderById(destFolderId);

  const lastRow = sheet.getLastRow();

  if (lastRow < 7) return;

  // Read links from A7:A
  const links = sheet.getRange(7, 1, lastRow - 6, 1).getValues();

  const statuses = [];
  const outputLinks = [];

  for (let i = 0; i < links.length; i++) {
    const row = i + 7;
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
        : cleanName + ".pdf";

      const createdPdf = destFolder.createFile(pdfBlob).setName(pdfName);

      statuses.push(["SUCCESS"]);
      outputLinks.push([createdPdf.getUrl()]);

    } catch (err) {

      statuses.push([
        `FAILED - ${err.message || err}`
      ]);

      outputLinks.push([""]);
    }
  }

  // Write results
  sheet.getRange(7, 2, statuses.length, 1).setValues(statuses); // B
  sheet.getRange(7, 3, outputLinks.length, 1).setValues(outputLinks); // C
}


/**
 * Export Google file as PDF
 */
function exportFileAsPdf_(fileId) {

  const url =
    `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/pdf`;

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    },
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(
      `Export failed (HTTP ${responseCode})`
    );
  }

  return response.getBlob().setContentType("application/pdf");
}


/**
 * Extract Google Drive ID
 */
function extractDriveId_(url) {

  const match =
    url.match(/\/d\/([a-zA-Z0-9_-]+)/) ||
    url.match(/folders\/([a-zA-Z0-9_-]+)/) ||
    url.match(/id=([a-zA-Z0-9_-]+)/);

  if (match && match[1]) {
    return match[1];
  }

  // Raw ID support
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url)) {
    return url;
  }

  throw new Error("Invalid Google Drive URL");
}


/**
 * Clean filename
 */
function sanitizeFilename_(name) {
  return String(name)
    .replace(/[\/\\:*?"<>|]/g, " ")
    .trim();
}
