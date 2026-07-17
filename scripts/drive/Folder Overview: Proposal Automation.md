# Proposal/Project Submission Package Automation

## Overview
Five Google Apps Script tools that automate repetitive Google Drive operations for proposal and project documentation workflows. All tools are input-driven via Google Sheets.

## Tech Stack
- Google Apps Script (JavaScript)
- Google Sheets
- Google Drive API
- Google Docs API

---

## Tool #1: Copy Folder Structure

**What it does:** Duplicates a folder and all its contents (including nested subfolders and files) from one location to another in Google Drive.

**Why:** Google Drive has no built-in "duplicate folder" feature. Manual copying of complex nested structures is error-prone and time-consuming.

**How to use:**
1. Enter source folder URL in B1
2. Enter destination folder URL in B2
3. Run `copyFolderFromSheet()`

**Output:** New copy of the folder structure in the destination location.

---

## Tool #2: Extract Folder Hierarchy

**What it does:** Scans a Google Drive folder recursively and exports its structure (folder names, file names, numbering) into a Google Sheet.

**Why:** Understanding large nested folder structures by clicking through Drive is slow. You need a readable outline you can review, sort, or export.

**How to use:**
1. Enter folder URL in B1 of "List of Attachment Tool" sheet
2. Run `extractHierarchicalFromB1()`
3. Optionally run `exportHierarchyToPDF()` to generate a PDF

**Output:** Two-column table in the sheet (numbering | name), starting at row 5. Natural sorting applied to match Drive display order.

---

## Tool #3: Apply Hierarchical Numbering

**What it does:** Recursively renames all folders and files in a directory with consistent hierarchical numbering (e.g., 1, 1.1, 1.1.1, 1.1.2, etc.).

**Why:** Proposal packages often need structured numbering for compliance or organization. Manual numbering across hundreds of files is impractical.

**How to use:**
1. Enter root folder URL in B1 of "Folder Numbering" sheet
2. Run `runFolderNumbering()`

**Behavior:**
- Detects existing prefix on root folder (e.g., "3.1")
- Uses that prefix; falls back to "1" if not found
- Applies sequential numbering to all subfolders and files
- Skips items already correctly numbered

**Output:** All folders and files in the tree renamed with hierarchical prefixes.

---

## Tool #4: Batch Convert Docs to PDF

**What it does:** Converts multiple Google Docs to PDFs and saves them to a destination folder, logging success/failure for each.

**Why:** Exporting docs one-by-one through the UI is tedious at scale. A batch process with status tracking eliminates manual work.

**How to use:**
1. Enter destination folder URL in B1 of "Batch PDFing" sheet
2. Paste Google Doc URLs in column A, starting at row 7
3. Run `batchConvertFromSheet()`

**Output:**
- Column B: SUCCESS or FAILED message for each doc
- Column C: URL of generated PDF (or empty if failed)

---

## Tool #5: Normalize Folder Names

**What it does:** Scans a folder tree and converts ALL-CAPS or inconsistently formatted folder names to Title Case, preserving technical acronyms (QA, ISO, HSE, etc.).

**Why:** Large shared folders accumulate inconsistent naming conventions. Standardizing them improves readability and professionalism without manual renaming.

**How to use:**
1. Enter folder URL in B1
2. Run `normalizeDriveFolderNaming()`

**Configurable settings in code:**
- `DRY_RUN`: Set to `true` to preview changes without applying them
- `THRESHOLD`: Percentage of uppercase letters (default 0.85 = 85%) required to trigger renaming
- `ACRONYMS`: Set of technical terms to preserve in uppercase (QC, QA, ISO, HSE, NDE, PQR, WPS, ITP, NDT)

**Output:** Folder names converted to Title Case with acronyms preserved. Shows count of scanned/renamed folders plus 10 example conversions.

---

## Quick Start

All tools expect:
- **Input:** Google Drive folder URLs or file URLs in specific Google Sheet cells
- **Output:** Operations logged in Drive or results written back to the sheet

To use these:
1. Set up a Google Sheet with tabs named: "Folder Numbering", "List of Attachment Tool", "Batch PDFing"
2. Copy the corresponding `.gs` script into Apps Script
3. Authorize the script to access Drive/Sheets
4. Enter URLs in the specified cells and run the function

Each tool handles its own error checking and reports status via alerts or logging.
