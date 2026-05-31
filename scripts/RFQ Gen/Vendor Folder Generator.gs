function createVendorFolders() {

  const SHEET_NAME = "Vendor Folder Creator";
  const ROOT_FOLDER_ID = "REPLACE_WITH_FOLDER_ID";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Missing sheet: ${SHEET_NAME}`);
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert("Script already running. Try again shortly.");
    return;
  }

  try {

    // -----------------------------
    // Extract folder ID safely
    // -----------------------------
    const folderUrl = String(sheet.getRange("B1").getValue()).trim();
    if (folderUrl) {
      const match = folderUrl.match(/[-\w]{25,}/);
      if (match) sheet.getRange("B2").setValue(match[0]);
    }

    // -----------------------------
    // Project prefix
    // -----------------------------
    const projectPrefix = String(sheet.getRange("A8").getValue())
      .trim()
      .substring(0, 8);

    if (!projectPrefix) {
      throw new Error("Missing project prefix in A8");
    }

    const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);

    const targetFolder = findProjectFolder_(rootFolder, projectPrefix);
    if (!targetFolder) throw new Error("Project folder not found");

    const managementFolder = findChildFolder_(targetFolder, "subtrade management");
    if (!managementFolder) throw new Error("Subtrade Management folder not found");

    const lastRow = sheet.getLastRow();
    if (lastRow < 10) {
      SpreadsheetApp.getUi().alert("No vendor data found.");
      return;
    }

    const data = sheet.getRange(10, 3, lastRow - 9, 2).getValues();

    data.forEach((row, i) => {

      const vendor = String(row[0] || "").trim();
      const bidNumber = String(row[1] || "").trim();

      if (!vendor || !bidNumber) return;

      try {

        const bidFolder = findBidFolder_(managementFolder, bidNumber);
        if (!bidFolder) throw new Error(`Bid folder not found: ${bidNumber}`);

        let receivedFolder = findChildFolder_(bidFolder, "received");

        if (!receivedFolder) {
          receivedFolder = bidFolder.createFolder("Received");
        }

        const vendorFolderName = buildVendorFolderName_(vendor);

        let vendorFolder = receivedFolder.getFoldersByName(vendorFolderName);
        vendorFolder = vendorFolder.hasNext()
          ? vendorFolder.next()
          : receivedFolder.createFolder(vendorFolderName);

        sheet.getRange(10 + i, 5).setValue(vendorFolder.getUrl());

      } catch (err) {
        sheet.getRange(10 + i, 5).setValue("ERROR: " + err.message);
      }

    });

    SpreadsheetApp.getActiveSpreadsheet()
      .toast("Vendor folders processed successfully.");

  } finally {
    lock.releaseLock();
  }
}

/* =========================
   SAFE HELPERS
========================= */

function findProjectFolder_(root, prefix) {
  const folders = root.getFolders();

  const safePrefix = prefix.toLowerCase();

  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().toLowerCase().includes(safePrefix)) {
      return f;
    }
  }
  return null;
}

function findChildFolder_(parent, name) {
  const folders = parent.getFolders();
  const target = name.toLowerCase();

  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().toLowerCase().includes(target)) {
      return f;
    }
  }
  return null;
}

/**
 * Safer bid matching (reduces false positives)
 */
function findBidFolder_(managementFolder, bidNumber) {

  const folders = managementFolder.getFolders();

  const safe = bidNumber.toLowerCase();

  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().toLowerCase().includes(safe)) {
      return f;
    }
  }

  // fallback: only last 2 digits, but still constrained
  const suffix = bidNumber.slice(-2);
  const regex = new RegExp(`(^|\\D)${suffix}(\\D|$)`);

  const retry = managementFolder.getFolders();

  while (retry.hasNext()) {
    const f = retry.next();
    if (regex.test(f.getName())) return f;
  }

  return null;
}

function buildVendorFolderName_(vendor) {
  return vendor
    .split(/\s+/)
    .slice(0, 2)
    .join(" ")
    .trim();
}
