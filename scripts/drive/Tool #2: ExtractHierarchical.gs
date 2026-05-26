function extractHierarchicalFromB1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("List of Attachment Tool");

  // B1 contains a Folder URL
  const folderUrl = sheet.getRange("B1").getValue().toString().trim();

  if (!folderUrl) {
    Logger.log("❌ No Folder URL in B1.");
    return;
  }

  // Extract folder ID from URL
  const match = folderUrl.match(/[-\w]{25,}/);
  if (!match) {
    Logger.log("❌ Invalid Folder URL in B1.");
    return;
  }

  const folderId = match[0];

  try {
    const rootFolder = DriveApp.getFolderById(folderId);
    const output = [];

    extractFolderContents(rootFolder, output);

    // Clear old output area from A5 down
    const lastRow = sheet.getLastRow();
    if (lastRow >= 5) {
      sheet.getRange(5, 1, lastRow - 4, 2).clear();
    }

    // Write results starting at A5
    sheet.getRange(5, 1, output.length, 2).setValues(output);

    Logger.log(`✅ Done! ${output.length} items extracted.`);
  } catch (e) {
    Logger.log("❌ Error: Invalid Folder URL or permission denied.");
    Logger.log(e);
  }
}


// Recursive folder traversal
function extractFolderContents(folder, output) {
  const folderName = folder.getName();
  const folderMatch = folderName.match(/^([\d\.]+)\s*(.*)/);

  if (folderMatch) {
    output.push([folderMatch[1], folderMatch[2]]);
  } else {
    output.push(["", folderName]);
  }

  // Files
  const files = [];
  const fileIter = folder.getFiles();

  while (fileIter.hasNext()) {
    files.push(fileIter.next());
  }

  files.sort((a, b) => naturalCompare(a.getName(), b.getName()));

  for (const file of files) {
    const fileName = file.getName();
    const fileMatch = fileName.match(/^([\d\.]+)\s*(.*)/);

    if (fileMatch) {
      output.push([fileMatch[1], fileMatch[2]]);
    } else {
      output.push(["", fileName]);
    }
  }

  // Subfolders
  const subfolders = [];
  const folderIter = folder.getFolders();

  while (folderIter.hasNext()) {
    subfolders.push(folderIter.next());
  }

  subfolders.sort((a, b) => naturalCompare(a.getName(), b.getName()));

  for (const subfolder of subfolders) {
    extractFolderContents(subfolder, output);
  }
}


// Natural sort function
function naturalCompare(a, b) {
  const ax = [], bx = [];

  a.replace(/(\d+)|(\D+)/g, function(_, $1, $2) {
    ax.push([$1 || Infinity, $2 || ""]);
  });

  b.replace(/(\d+)|(\D+)/g, function(_, $1, $2) {
    bx.push([$1 || Infinity, $2 || ""]);
  });

  while (ax.length && bx.length) {
    const an = ax.shift();
    const bn = bx.shift();
    const nn = (an[0] - bn[0]) || an[1].localeCompare(bn[1]);
    if (nn) return nn;
  }

  return ax.length - bx.length;
}


// ================================
// PDF EXPORT FUNCTION
// ================================
function exportHierarchyToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("List of Attachment Tool");

  const data = sheet.getRange(5, 1, sheet.getLastRow() - 4, 2).getValues();

  if (!data || data.length === 0) {
    Logger.log("❌ No data to export.");
    return;
  }

  let content = "Google Drive Folder Structure Export\n\n";

  data.forEach(row => {
    const number = row[0] || "";
    const name = row[1] || "";
    content += `${number ? number + " " : ""}${name}\n`;
  });

  const blob = Utilities.newBlob(content, "text/plain", "Drive_Hierarchy.txt");
  const pdf = blob.getAs("application/pdf");

  const file = DriveApp.createFile(pdf);
  file.setName("Drive_Hierarchy_Export.pdf");

  Logger.log("✅ PDF created: " + file.getUrl());
}
