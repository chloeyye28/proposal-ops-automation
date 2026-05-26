function runFolderNumbering() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Folder Numbering");

  if (!sheet) {
    ui.alert(
      "❗ Sheet not found",
      `Missing sheet: "Folder Numbering". Available sheets:\n\n` +
        ss.getSheets().map(s => s.getName()).join("\n"),
      ui.ButtonSet.OK
    );
    return;
  }

  // Folder URL from B1
  const folderUrl = String(sheet.getRange("B1").getValue() || "").trim();

  const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
  if (!folderIdMatch) {
    ui.alert(
      "❗ Invalid URL",
      "Please enter a valid Google Drive folder URL in B1.",
      ui.ButtonSet.OK
    );
    return;
  }

  const folderId = folderIdMatch[0];

  try {
    const rootFolder = DriveApp.getFolderById(folderId);
    const rootName = rootFolder.getName().trim();

    // Detect prefix like: 3, 3.1, 3.1.2
    let prefixMatch = rootName.match(/^(\d+(?:\.\d+)*)\s*(?:[.)-])?/);

    // Fallback to parent folder if needed
    if (!prefixMatch) {
      const parents = rootFolder.getParents();
      if (parents.hasNext()) {
        const parentName = parents.next().getName().trim();
        prefixMatch = parentName.match(/^(\d+(?:\.\d+)*)\s*(?:[.)-])?/);
      }
    }

    const rootPrefix = prefixMatch ? prefixMatch[1] : "1";

    function renameChildren(folder, prefix) {
      let itemCounter = 1;

      // Subfolders
      const subfolders = [];
      const subIterator = folder.getFolders();
      while (subIterator.hasNext()) subfolders.push(subIterator.next());

      subfolders.sort((a, b) => a.getName().localeCompare(b.getName()));

      for (const subfolder of subfolders) {
        const currentPrefix = `${prefix}.${itemCounter}`;
        const originalName = subfolder.getName();

        const newName = originalName.startsWith(currentPrefix)
          ? originalName
          : `${currentPrefix} ${originalName}`;

        subfolder.setName(newName);

        renameChildren(subfolder, currentPrefix);
        itemCounter++;
      }

      // Files
      const files = [];
      const fileIterator = folder.getFiles();
      while (fileIterator.hasNext()) files.push(fileIterator.next());

      files.sort((a, b) => a.getName().localeCompare(b.getName()));

      for (const file of files) {
        const currentPrefix = `${prefix}.${itemCounter}`;
        const originalName = file.getName();

        const newName = originalName.startsWith(currentPrefix)
          ? originalName
          : `${currentPrefix} ${originalName}`;

        file.setName(newName);
        itemCounter++;
      }
    }

    renameChildren(rootFolder, rootPrefix);

    ui.alert(`✅ Numbering complete! Root prefix used: ${rootPrefix}`);

  } catch (e) {
    ui.alert("⚠️ Error: " + e.message);
  }
}
