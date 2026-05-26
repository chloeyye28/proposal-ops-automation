// 🔹 Extract Sheet ID from URL in E1 and output to B1
function extractSheetIdFromUrl() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetUrl = sheet.getRange("E1").getValue().toString().trim();

  if (sheetUrl) {
    const idMatch = sheetUrl.match(/[-\w]{25,}/);

    if (idMatch) {
      sheet.getRange("B1").setValue(idMatch[0]);
    }
  }
}

// 🔹 Main function
function createRFQDraft() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Run helper function
  extractSheetIdFromUrl();

  // 📧 Email fields from sheet
  const toEmail = sheet.getRange("B6").getValue();

  const bccEmail = sheet.getRange("B7")
    .getDisplayValue()
    .split(/[,\n;]/)
    .map(e => e.trim())
    .filter(String)
    .join(", ");

  const subjectLine = sheet.getRange("B8").getValue();

  // 📌 Bid / RFQ reference
  const bidNumber = sheet.getRange("E2").getValue();

  const bidKey = bidNumber.toString().slice(-2);
  const bidPrefix = bidNumber.toString().substring(0, 8);

  // =========================================================
  // CONFIGURATION
  // Replace with your own parent folder ID
  // =========================================================
  const PARENT_FOLDER_ID = "YOUR_PARENT_FOLDER_ID";

  // 📂 Step 1: Locate project folder
  const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);

  const projectFolders = parentFolder.getFolders();

  let bidFolder = null;

  while (projectFolders.hasNext()) {

    const folder = projectFolders.next();

    if (folder.getName().startsWith(bidPrefix)) {
      bidFolder = folder;
      break;
    }
  }

  if (!bidFolder) {

    SpreadsheetApp.getUi().alert(
      "No matching project folder found."
    );

    return;
  }

  // 📂 Step 2: Locate subfolder
  const subFolders = bidFolder.getFolders();

  let workflowFolder = null;

  while (subFolders.hasNext()) {

    const folder = subFolders.next();

    if (
      folder
        .getName()
        .toLowerCase()
        .includes("subtrade")
    ) {

      workflowFolder = folder;
      break;
    }
  }

  if (!workflowFolder) {

    SpreadsheetApp.getUi().alert(
      "Required workflow folder not found."
    );

    return;
  }

  // 📂 Step 3: Find Google Sheet
  const files = workflowFolder.getFiles();

  let generatorFile = null;

  while (files.hasNext()) {

    const file = files.next();

    if (
      file.getName().toLowerCase().includes("generator") &&
      file.getMimeType() === MimeType.GOOGLE_SHEETS
    ) {

      generatorFile = file;
      break;
    }
  }

  if (!generatorFile) {

    SpreadsheetApp.getUi().alert(
      "Generator sheet not found."
    );

    return;
  }

  const generatorFileId = generatorFile.getId();

  // 🔗 Write generator file link to sheet
  sheet.getRange("E3").setFormula(
    `=HYPERLINK(
      "https://docs.google.com/spreadsheets/d/${generatorFileId}",
      "${generatorFile.getName()}"
    )`
  );

  // 🔗 Write parent folder link to sheet
  const workflowFolderId = workflowFolder.getId();

  sheet.getRange("E4").setFormula(
    `=HYPERLINK(
      "https://drive.google.com/drive/folders/${workflowFolderId}",
      "${workflowFolder.getName()}"
    )`
  );

  // 📑 Open generator spreadsheet
  const generatorSpreadsheet =
    SpreadsheetApp.openById(generatorFileId);

  const tabs = generatorSpreadsheet.getSheets();

  let targetSheet = null;

  for (let s of tabs) {

    if (
      s.getName()
        .toLowerCase()
        .includes(bidKey.toLowerCase())
    ) {

      targetSheet = s;
      break;
    }
  }

  if (!targetSheet) {

    SpreadsheetApp.getUi().alert(
      "Matching tab not found."
    );

    return;
  }

  // 🔎 Locate content anchor
  const values = targetSheet
    .getDataRange()
    .getValues();

  let anchorRow = -1;

  for (let r = 0; r < values.length; r++) {

    if (values[r].includes("Email Content Generator")) {

      anchorRow = r + 1;
      break;
    }
  }

  if (anchorRow === -1) {

    SpreadsheetApp.getUi().alert(
      "Content section not found."
    );

    return;
  }

  // 📑 Extract formatted content
  const lastRow = targetSheet.getLastRow();

  const range = targetSheet.getRange(
    anchorRow + 1,
    1,
    lastRow - anchorRow,
    4
  );

  const richTexts = range.getRichTextValues();

  // 🎨 Convert sheet formatting into HTML
  let htmlBody =
    "<table style='border-collapse:collapse;'>";

  for (let r = 0; r < richTexts.length; r++) {

    htmlBody += "<tr>";

    for (let c = 0; c < richTexts[r].length; c++) {

      const rt = richTexts[r][c];

      let cellText = rt
        .getText()
        .replace(/\n/g, "<br>");

      let style = "";

      const textStyle = rt.getTextStyle();

      if (textStyle.isBold()) {
        style += "font-weight:bold;";
      }

      if (textStyle.isItalic()) {
        style += "font-style:italic;";
      }

      if (textStyle.isUnderline()) {
        style += "text-decoration:underline;";
      }

      htmlBody += `
        <td style="padding:4px; ${style}">
          ${cellText}
        </td>
      `;
    }

    htmlBody += "</tr>";
  }

  htmlBody += "</table>";

  // 📨 Build Gmail draft payload
  const raw = Utilities.base64EncodeWebSafe(

    Utilities.newBlob(

      "Content-Type: text/html; charset=UTF-8\r\n" +
      "MIME-Version: 1.0\r\n" +
      "to: " + toEmail + "\r\n" +
      "bcc: " + bccEmail + "\r\n" +
      "subject: " + subjectLine + "\r\n\r\n" +
      htmlBody

    ).getBytes()
  );

  // ✉️ Create Gmail draft
  Gmail.Users.Drafts.create(
    {
      message: {
        raw: raw
      }
    },
    "me"
  );

  SpreadsheetApp.getUi().alert(
    "Draft created successfully!"
  );
}
  
