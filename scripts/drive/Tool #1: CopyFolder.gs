function copyFolderFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sourceId = extractDriveId_(sheet.getRange("B1").getValue().toString().trim());
  const destinationId = extractDriveId_(sheet.getRange("B2").getValue().toString().trim());
  
  function extractDriveId_(input) {
  if (!input) return "";

  // If it's already an ID (no slashes), just return it.
  if (input.indexOf("/") === -1) return input;

  // Common Google Drive folder URL patterns:
  // https://drive.google.com/drive/folders/<ID>
  // https://drive.google.com/open?id=<ID>
  // Also handles .../folders/<ID>?...
  const m = input.match(/(?:\/folders\/|[?&]id=)([-\w]{10,})/);
  if (m && m[1]) return m[1];

  // Fallback: try to find a long-ish Drive-looking token
  const m2 = input.match(/[-\w]{25,}/);
  return (m2 && m2[0]) ? m2[0] : "";
}
 
  const ui = SpreadsheetApp.getUi();

  if (!sourceId || !destinationId) {
    ui.alert("❗ Missing Input", "Please enter both Source and Destination Folder IDs in cells B1 and B2.", ui.ButtonSet.OK);
    return;
  }

  try {
    const sourceFolder = DriveApp.getFolderById(sourceId);
    const destinationParent = DriveApp.getFolderById(destinationId);
    copyContentsRecursive(sourceFolder, destinationParent);
    ui.alert("✅ Success", `Copied contents of '${sourceFolder.getName()}' into '${destinationParent.getName()}'`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("❌ Error", e.message, ui.ButtonSet.OK);
  }
}

function copyContentsRecursive(source, target) {
  const files = source.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.makeCopy(file.getName(), target);
  }

  const subfolders = source.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const newSubfolder = target.createFolder(subfolder.getName());
    copyContentsRecursive(subfolder, newSubfolder);
  }
}

