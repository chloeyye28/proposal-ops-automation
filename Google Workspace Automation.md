# Google Workspace Proposal Automation Toolkit

## About
A collection of Google Apps Script tools designed to automate proposal and project-related workflows inside Google Workspace, including Google Drive, Google Docs, and Google Sheets.This toolkit reduces manual effort involved in setting up, organizing, and managing proposal/project documentation structures.

## Tech Stack
- Google Apps Script (JavaScript-based)
- Google Sheets
- Google Drive API
- Google Docs API

## Features Overview
- Automated Google Drive folder creation and duplication  
- Recursive folder copying with full structure preservation  
- Proposal and project setup automation using templates  
- Input-driven execution via Google Sheets  
- Google Apps Script-based workflow engine  

---

## Tool #1: Google Drive Folder Duplication Automation

### Problem
Google Drive does not provide a native one-click method to duplicate folders with complex nested structures.

Common workarounds include downloading folders as ZIP files and re-uploading them. However, this approach becomes unreliable for:
- Large folder structures
- Deeply nested hierarchies
- Files with naming or size constraints during compression and extraction
- Long Downloading and uploading time

### Solution
A Google Apps Script-based automation tool that enables full folder duplication in Google Drive while preserving:
- Folder hierarchy
- File structure
- Nested subfolders

The process is executed directly within Google Workspace without requiring external tools.

### Impact
- Eliminates manual folder duplication work for proposal and project setup
- Supports complex folder structures without compression limitations
- Improves consistency and reliability of document setup workflows

## Workflow
Google Sheets (Source & Destination Folder URL input) → Apps Script Engine → Google Drive Automation

---

## Tool #2: Google Drive Hierarchical Structure Extractor

### Problem
Google Drive does not provide a native way to export or visualize folder structures in a structured, hierarchical format.
Manually inspecting nested folders is inefficient and does not scale for large documentation systems.

### Solution
A Google Apps Script tool that recursively traverses a Google Drive folder structure and extracts:

- Folder hierarchy
- File names
- Structured numbering patterns (if present in naming conventions)
The output is written into a Google Sheet in a structured, readable format. Easy to copy and paste into a table of contents template, and also has the option to export to a PDF file.

### Impact
- Enables quick visualization of complex Drive folder structures  
- Reduces manual effort required to audit or map document systems  
- Improves visibility into proposal/document organization  
- Supports structured documentation workflows in Google Workspace  

### Features
- Recursive folder traversal (nested structure support)  
- File and folder extraction in hierarchical order  
- Natural sorting of alphanumeric file names  
- Pattern recognition for numbered naming structures  
- Outputs structured data directly into Google Sheets  

## Workflow
Google Sheets (Folder URL input) → Apps Script → Google Drive traversal → Structured hierarchy output in Sheet --> Exported to PDF (optional)
