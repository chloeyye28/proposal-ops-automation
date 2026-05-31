function createRFQDraft() {

  const PARENT_FOLDER_ID = "REPLACE_WITH_FOLDER_ID";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert("Script already running. Try again shortly.");
    return;
  }

  try {

    // -----------------------------
    // Extract sheet ID safely
    // -----------------------------
    extractSheetIdFromUrlSafe_(sheet);

    const toEmail = String(sheet.getRange("B6").getValue()).trim();

    const bccEmail = String(sheet.getRange("B7").getDisplayValue())
      .split(/[,\n;]/)
      .map(s => s.trim())
      .filter(Boolean)
      .join(", ");

    const subjectLine = String(sheet.getRange("B8").getValue()).trim();

    const bidNumber = String(sheet.getRange("E2").getValue()).trim();

    if (!bidNumber) throw new Error("Missing bid number (E2)");

    const bidPrefix = bidNumber.substring(0, 8);
    const bidKey = bidNumber.slice(-2);

    // -----------------------------
    // Drive navigation
    // -----------------------------
    const root = DriveApp.getFolderById(PARENT_FOLDER_ID);

    const bidFolder = findFolderByPrefix_(root, bidPrefix);
    if (!bidFolder) throw new Error("Project folder not found");

    const workflowFolder = findFolderContains_(bidFolder, "subtrade");
    if (!workflowFolder) throw new Error("Workflow folder not found");

    const generatorFile = findSheetFile_(workflowFolder, "generator");
    if (!generatorFile) throw new Error("Generator file not found");

    const generatorId = generatorFile.getId();

    sheet.getRange("E3").setFormula(
      `=HYPERLINK("https://docs.google.com/spreadsheets/d/${generatorId}", "${generatorFile.getName()}")`
    );

    sheet.getRange("E4").setFormula(
      `=HYPERLINK("https://drive.google.com/drive/folders/${workflowFolder.getId()}", "${workflowFolder.getName()}")`
    );

    // -----------------------------
    // Open generator sheet
    // -----------------------------
    const genSheet = SpreadsheetApp.openById(generatorId);
    const tabs = genSheet.getSheets();

    const targetTab = tabs.find(s =>
      s.getName().toLowerCase().includes(bidKey)
    );

    if (!targetTab) throw new Error("Matching tab not found");

    // -----------------------------
    // Find anchor
    // -----------------------------
    const anchorRow = findAnchorRow_(targetTab, "Email Content Generator");
    if (!anchorRow) throw new Error("Content anchor not found");

    // -----------------------------
    // Extract HTML content
    // -----------------------------
    const htmlBody = buildHtmlFromRange_(targetTab, anchorRow);

    // -----------------------------
    // Build Gmail draft
    // -----------------------------
    const rawMessage = buildMimeMessage_(
      toEmail,
      bccEmail,
      subjectLine,
      htmlBody
    );

    const encoded = Utilities.base64EncodeWebSafe(
      Utilities.newBlob(rawMessage).getBytes()
    );

    Gmail.Users.Drafts.create(
      { message: { raw: encoded } },
      "me"
    );

    SpreadsheetApp.getUi().alert("Draft created successfully!");

  } finally {
    lock.releaseLock();
  }
}

/* =========================
   HELPERS (SAFE + MODULAR)
========================= */

function extractSheetIdFromUrlSafe_(sheet) {
  const url = String(sheet.getRange("E1").getValue()).trim();
  const match = url.match(/[-\w]{25,}/);
  if (match) sheet.getRange("B1").setValue(match[0]);
}

function findFolderByPrefix_(root, prefix) {
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().startsWith(prefix)) return f;
  }
  return null;
}

function findFolderContains_(parent, text) {
  const folders = parent.getFolders();
  const needle = text.toLowerCase();

  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().toLowerCase().includes(needle)) return f;
  }
  return null;
}

function findSheetFile_(folder, keyword) {
  const files = folder.getFiles();
  const needle = keyword.toLowerCase();

  while (files.hasNext()) {
    const f = files.next();
    if (
      f.getName().toLowerCase().includes(needle) &&
      f.getMimeType() === MimeType.GOOGLE_SHEETS
    ) return f;
  }
  return null;
}

function findAnchorRow_(sheet, text) {
  const values = sheet.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i].includes(text)) return i + 1;
  }
  return null;
}

function buildHtmlFromRange_(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(startRow + 1, 1, lastRow - startRow, 4);
  const rich = range.getRichTextValues();

  let html = "<table style='border-collapse:collapse;'>";

  rich.forEach(row => {
    html += "<tr>";

    row.forEach(cell => {
      const text = cell.getText().replace(/\n/g, "<br>");
      const style = cell.getTextStyle();

      let css = "";
      if (style.isBold()) css += "font-weight:bold;";
      if (style.isItalic()) css += "font-style:italic;";
      if (style.isUnderline()) css += "text-decoration:underline;";

      html += `<td style="padding:4px;${css}">${text}</td>`;
    });

    html += "</tr>";
  });

  html += "</table>";
  return html;
}

function buildMimeMessage_(to, bcc, subject, html) {
  return (
    "Content-Type: text/html; charset=UTF-8\r\n" +
    "MIME-Version: 1.0\r\n" +
    "to: " + to + "\r\n" +
    "bcc: " + bcc + "\r\n" +
    "subject: " + subject + "\r\n\r\n" +
    html
  );
}
