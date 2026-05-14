/**
 * Normalize Google Drive folder names from a folder URL entered in B1.
 *
 * Features:
 * - Recursively scans nested folders
 * - Converts MOSTLY-UPPERCASE names into Title Case
 * - Preserves known acronyms (QA, QC, ISO, etc.)
 * - Supports preview mode (Dry Run)
 *
 * Sheet Setup:
 * B1 = Source Google Drive Folder URL
 */

function normalizeDriveFolderNaming() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const folderUrl = sheet.getRange("B1").getValue().toString().trim();

  if (!folderUrl) {
    ui.alert(
      "Missing Input",
      "Please paste a Google Drive folder URL into cell B1.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Extract folder ID
  const match = folderUrl.match(/[-\w]{25,}/);

  if (!match) {
    ui.alert(
      "Invalid URL",
      "Could not extract a valid Google Drive Folder ID.",
      ui.ButtonSet.OK
    );
    return;
  }

  const rootFolder = DriveApp.getFolderById(match[0]);

  // ================= SETTINGS =================

  const DRY_RUN = false;

  // Rename if 85% or more of letters are uppercase
  const THRESHOLD = 0.85;

  const ACRONYMS = new Set([
    "QC",
    "QA",
    "NDE",
    "HSE",
    "ISO",
    "PQR",
    "WPS",
    "ITP",
    "NDT"
  ]);

  const SMALL_WORDS = new Set([
    "and",
    "or",
    "the",
    "of",
    "to",
    "in",
    "for",
    "a",
    "an",
    "on",
    "at",
    "by",
    "with",
    "from"
  ]);

  // ============================================

  let scanned = 0;
  let renamed = 0;

  const examples = [];

  function isMostlyUppercase(name, threshold) {

    const letters = name.match(/[A-Za-z]/g);

    if (!letters || letters.length === 0) {
      return false;
    }

    let uppercaseCount = 0;

    for (const char of letters) {
      if (char === char.toUpperCase()) {
        uppercaseCount++;
      }
    }

    return (uppercaseCount / letters.length) >= threshold;
  }

  function toSmartTitleCase(name) {

    const parts = name.split(/(\s+|[-/()]+|&+)/);

    let wordIndex = 0;

    return parts.map(token => {

      // Keep separators unchanged
      if (/^(\s+|[-/()]+|&+)$/.test(token)) {
        return token;
      }

      // Skip tokens without letters
      if (!/[A-Za-z]/.test(token)) {
        return token;
      }

      const upper = token.toUpperCase();
      const lower = token.toLowerCase();

      const lettersOnly =
        token.replace(/[^A-Za-z]/g, "").toUpperCase();

      // Preserve acronyms
      if (
        ACRONYMS.has(upper) ||
        ACRONYMS.has(lettersOnly)
      ) {
        wordIndex++;
        return token.replace(
          /[A-Za-z]+/g,
          match => match.toUpperCase()
        );
      }

      // Lowercase small words unless first word
      if (
        SMALL_WORDS.has(lower) &&
        wordIndex > 0
      ) {
        wordIndex++;
        return lower;
      }

      // Standard Title Case
      const titled =
        lower.replace(/[a-z]/, c => c.toUpperCase());

      wordIndex++;

      return titled;

    }).join("");
  }

  function normalizeFolder(folder) {

    scanned++;

    const oldName = folder.getName();
    let newName = oldName;

    if (isMostlyUppercase(oldName, THRESHOLD)) {
      newName = toSmartTitleCase(oldName);
    }

    if (newName !== oldName) {

      if (!DRY_RUN) {
        folder.setName(newName);
      }

      renamed++;

      if (examples.length < 10) {
        examples.push(`${oldName} → ${newName}`);
      }
    }

    // Process subfolders recursively
    const subfolders = folder.getFolders();

    while (subfolders.hasNext()) {
      normalizeFolder(subfolders.next());
    }
  }

  try {

    normalizeFolder(rootFolder);

    ui.alert(
      `${DRY_RUN ? "Preview Complete" : "Normalization Complete"}\n\n` +
      `Folders Scanned: ${scanned}\n` +
      `Folders Renamed: ${renamed}\n\n` +
      (examples.length
        ? `Examples:\n\n${examples.join("\n")}`
        : "No changes required.")
    );

  } catch (err) {

    ui.alert(
      "Error",
      err.message,
      ui.ButtonSet.OK
    );
  }
}
