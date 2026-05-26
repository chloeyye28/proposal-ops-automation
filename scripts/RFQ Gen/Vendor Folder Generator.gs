function createVendorFolders() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Vendor Folder Creator");

  if (!sheet) return;

  // -----------------------------------
  // Extract folder ID from URL input
  // -----------------------------------

  const folderUrl = sheet.getRange("B1").getValue().toString().trim();

  if (folderUrl) {
    const idMatch = folderUrl.match(/[-\w]{25,}/);

    if (idMatch) {
      sheet.getRange("B2").setValue(idMatch[0]);
    }
  }

  // -----------------------------------
  // Read project/bid prefix
  // Example: Q25-1103
  // -----------------------------------

  const projectPrefix = sheet
    .getRange("A8")
    .getValue()
    .toString()
    .trim()
    .substring(0, 8);

  // -----------------------------------
  // Root projects folder
  // Replace with your own folder ID
  // -----------------------------------

  const ROOT_FOLDER_ID = "REPLACE_WITH_FOLDER_ID";

  const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);

  // -----------------------------------
  // Locate matching project folder
  // -----------------------------------

  let targetFolder = null;

  const folders = rootFolder.getFolders();

  while (folders.hasNext()) {

    const folder = folders.next();

    if (
      folder
        .getName()
        .toLowerCase()
        .includes(projectPrefix.toLowerCase())
    ) {
      targetFolder = folder;
      break;
    }
  }

  if (!targetFolder) {

    Logger.log(
      "No matching project folder found."
    );

    return;
  }

  // -----------------------------------
  // Locate sub-management folder
  // -----------------------------------

  let managementFolder = null;

  const subfolders = targetFolder.getFolders();

  while (subfolders.hasNext()) {

    const folder = subfolders.next();

    if (
      folder
        .getName()
        .toLowerCase()
        .includes("subtrade management")
    ) {
      managementFolder = folder;
      break;
    }
  }

  if (!managementFolder) {

    Logger.log(
      "Management folder not found."
    );

    return;
  }

  // -----------------------------------
  // Read vendor + bid data
  // -----------------------------------

  const lastRow = sheet.getLastRow();

  if (lastRow < 10) {

    SpreadsheetApp
      .getUi()
      .alert("No vendor data found.");

    return;
  }

  // Column C = Vendor
  // Column D = Bid Number

  const data = sheet
    .getRange(10, 3, lastRow - 9, 2)
    .getValues();

  // -----------------------------------
  // Process each vendor row
  // -----------------------------------

  for (let i = 0; i < data.length; i++) {

    let [vendor, bidNumber] = data[i];

    if (!vendor || !bidNumber) continue;

    let currentFolder = null;

    // -----------------------------------
    // Search by full bid number
    // -----------------------------------

    const bidFolders = managementFolder.getFolders();

    while (bidFolders.hasNext()) {

      const folder = bidFolders.next();

      if (
        folder
          .getName()
          .toLowerCase()
          .includes(bidNumber.toLowerCase())
      ) {
        currentFolder = folder;
        break;
      }
    }

    // -----------------------------------
    // Fallback search by suffix
    // -----------------------------------

    if (!currentFolder) {

      const shortBid = bidNumber.slice(-2);

      const regex = new RegExp(
        shortBid.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
        'i'
      );

      const altFolders = managementFolder.getFolders();

      while (altFolders.hasNext()) {

        const folder = altFolders.next();

        if (regex.test(folder.getName())) {
          currentFolder = folder;
          break;
        }
      }
    }

    if (!currentFolder) continue;

    // -----------------------------------
    // Locate or create "Received" folder
    // -----------------------------------

    let receivedFolder = null;

    const innerFolders = currentFolder.getFolders();

    while (innerFolders.hasNext()) {

      const folder = innerFolders.next();

      if (
        folder
          .getName()
          .toLowerCase()
          .includes("received")
      ) {
        receivedFolder = folder;
        break;
      }
    }

    if (!receivedFolder) {

      receivedFolder =
        currentFolder.createFolder("Received");

      Logger.log(
        `Created folder inside: ${currentFolder.getName()}`
      );
    }

    // -----------------------------------
    // Create vendor folder
    // Uses first two words of vendor name
    // -----------------------------------

    const vendorFolderName = vendor
      .trim()
      .split(/\s+/)
      .slice(0, 2)
      .join(" ");

    const existing =
      receivedFolder.getFoldersByName(vendorFolderName);

    const vendorFolder = existing.hasNext()
      ? existing.next()
      : receivedFolder.createFolder(vendorFolderName);

    // -----------------------------------
    // Output folder URL to sheet
    // -----------------------------------

    sheet
      .getRange(10 + i, 5)
      .setValue(vendorFolder.getUrl());
  }

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast("Vendor folders processed successfully.");
}
