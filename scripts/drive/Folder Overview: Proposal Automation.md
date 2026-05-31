# Proposal/Project Submission Package Automation

## About
This toolkit reduces manual effort involved in setting up, organizing, and managing proposal/project documentation structures.

## Tech Stack
- Google Apps Script (JavaScript)
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

```text
Google Sheets (Source & Destination Folder URL input) 
↓
Apps Script Engine 
↓
Google Drive Automation
```

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

```text
Google Sheets (Folder URL input)
↓
Apps Script → Google Drive traversal
↓
Structured hierarchy output in Sheet
↓
Exported to PDF (optional)
```

---

## Tool #3: Google Drive Hierarchical Folder Numbering Automation

### Problem
Google Drive does not provide a fast way to enforce structured hierarchical numbering across folders and files folders. As a result, when folders such as proposal submission packages contain complex nested structures requiring hierarchical numbering, files must be renamed individually, which is extremely time-consuming. The only alternative is using third-party batch renaming add-ons, which introduces unnecessary risks of exposing project information.

### Solution
A Google Apps Script-based automation tool that recursively applies hierarchical numbering to folders and files in Google Drive.

The system:
- Detects existing numeric prefixes for parent folders (e.g., 1, 1.1, 1.1.1)  
- Applies structured hierarchical numbering to all subfolders and files  
- Maintains sorted order using natural sorting logic  
- Renames items directly in Google Drive while preserving structure  

This creates a consistent hierarchical structure across the entire folder system.

### Impact
- Enforces consistent folder and file naming structure
- Completing hirarchical numbering for complex folders in one-click  
- Reduces manual effort required for organizing proposal/project directories  
- Standardizes hierarchy across teams and workflows  

## Workflow

```text
Google Sheets (Root Folder URL input)
↓
Apps Script Engine
↓
Drive Structure Analysis
↓
Recursive Renaming Engine
↓
Google Drive (Updated Hierarchical Naming System)
```
---

## Tool #4: Batch Google Docs to PDF Export Automation

### Problem
Google Drive does not provide a scalable way to batch-convert multiple Google Docs into PDFs while also organizing outputs into a structured workflow.

As a result, users often rely on manual steps:

- Opening each document individually  
- Exporting to PDF one by one  
- Downloading and re-uploading files  
- Manually tracking conversion status  

This becomes inefficient and error-prone when dealing with large batches of documents.

### Solution
A Google Apps Script-based batch automation tool that converts multiple Google Docs into PDFs using a spreadsheet-driven workflow.

The system:

- Reads a list of Google Doc URLs from Google Sheets  
- Converts each document into PDF using Google Drive export API  
- Saves PDFs into a specified destination folder  
- Logs success/failure status per document  
- Outputs generated PDF links back into the sheet  

This creates a fully automated batch document processing pipeline inside Google Workspace.

### Impact
- Eliminates repetitive manual PDF exporting tasks  
- Enables bulk document processing from a single interface  
- Improves tracking with real-time success/failure status  
- Centralizes document conversion workflow inside Google Sheets  
- Reduces operational overhead for proposal/document preparation  

## Workflow

```text
Google Sheets (Doc URL Input + Destination Folder)
↓
Apps Script Engine
↓
Google Drive Export API
↓
PDF Generation
↓
Google Drive Storage
↓
Status Logging in Google Sheet
```

---

## Tool #5: Google Drive Folder Naming Normalization

### Problem

Large proposal and project folders often develop inconsistent naming conventions due to manual folder creation, including:

- FULLY CAPITALIZED folder names
- inconsistent formatting styles
- lack of standardization across projects

This reduced readability and created inconsistencies in shared document environments.

### Solution

A Google Apps Script automation tool that recursively scans Google Drive folder structures and automatically normalizes folder naming conventions using configurable formatting rules.

The tool:
- detects mostly-uppercase folder names
- converts names into standardized Title Case
- preserves technical acronyms (QA, QC, ISO, etc.)
- processes nested subfolders recursively
- supports preview mode before applying changes

### Impact

- Reduced manual folder/file naming cleanup work
- Standardized naming conventions across proposal/project environments
- Improved readability of large folder structures
 
### Features

- Recursive folder scanning
- Uppercase detection logic
- Smart Title Case conversion
- Acronym preservation system
- Configurable formatting thresholds
- Google Sheets input-driven workflow

## Workflow

```text
Google Sheets (Targeted Folder URL)
↓
Apps Script Engine
↓
Recursive Google Drive Processing
↓
Folder Naming Standardization
```
---
