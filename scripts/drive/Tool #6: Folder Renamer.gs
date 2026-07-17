function startUniversalRenamer() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Batch Files Rename Tool");

  // Inputs
  const folderUrl = sheet.getRange("B1").getValue().toString().trim();
  const findText = sheet.getRange("B2").getValue().toString();
  const replaceText = sheet.getRange("B3").getValue().toString();

  const ui = SpreadsheetApp.getUi();


  // Validate inputs
  if (!folderUrl || findText === "") {

    ui.alert(
      "❗ Missing Input",
      "Please provide:\n\nB1 = Folder URL\nB2 = Text to Find",
      ui.ButtonSet.OK
    );

    return;
  }


  // Extract Folder ID
  const folderIdMatch =
    folderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/) ||
    folderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);


  const folderId = folderIdMatch ? folderIdMatch[1] : "";


  if (!folderId) {

    ui.alert(
      "❗ Invalid Folder URL",
      "Please enter a valid Google Drive folder URL.",
      ui.ButtonSet.OK
    );

    return;
  }


  try {

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();

    let checkedCount = 0;
    let renamedCount = 0;


    // Escape special characters so B2 works like CTRL + H
    const escapedFind = findText.replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&"
    );


    const replaceRegex = new RegExp(
      escapedFind,
      "g"
    );


    while (files.hasNext()) {

      const file = files.next();

      checkedCount++;

      const originalName = file.getName();

      let newName = originalName;


      // ------------------------------------
      // Remove all "Copy of " prefixes
      // ------------------------------------
      newName = newName.replace(
        /^(Copy of )+/i,
        ""
      );


      // ------------------------------------
      // CTRL + H style replacement
      // Keeps replacement at the same location
      // ------------------------------------
      newName = newName.replace(
        replaceRegex,
        function () {
          return replaceText;
        }
      );


      // Rename only if changed
      if (newName !== originalName) {

        file.setName(newName);

        renamedCount++;

      }

    }


    ui.alert(
      "✅ Rename Completed\n\n" +
      "Files Checked: " + checkedCount +
      "\nFiles Renamed: " + renamedCount
    );


  } catch (err) {

    ui.alert(
      "⚠️ Error",
      "Could not rename files:\n\n" + err.message,
      ui.ButtonSet.OK
    );

  }

}
