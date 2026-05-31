# RFQ Automation

## End-to-End Workflow Map

```text
Step #1: RFQ Email Generator
        ↓
Step #2: Vendor Folder Generator
        ↓
Step #3: Save Quotes Automation
        ↓
Step #4: Logged Sent Quotes Tracker
```

## System Architecture Flow

```text
Google Sheets (Control Panel)
        ↓
Google Apps Script Engine
        ↓
Gmail + Google Drive APIs
        ↓
Automated Folder Structure + Email Generation + PDF Archiving
        ↓
Execution Logs Written Back to Google Sheets
```

---

## About

This project is a Google Apps Script-based automation suite designed to streamline RFQ (Request for Quote), vendor management, and document handling workflows within Google Workspace.

The system automates repetitive operational tasks including:

- RFQ email generation
- Vendor and project folder creation
- Email and attachment archiving
- Quote tracking and logging

By using Google Sheets as a control panel and Apps Script as the workflow engine, the solution creates a lightweight workflow orchestration system entirely within Google Workspace.

---

## Tech Stack

- Google Apps Script
- Google Sheets
- Gmail API
- Google Drive API
- Google Workspace

---

## Features Overview

- Automated RFQ email draft generation from structured sheet data
- Dynamic vendor folder creation in Google Drive
- Email archiving with PDF and attachment extraction
- Structured quote tracking and status logging
- Spreadsheet-driven workflow management
- Fully serverless Google Workspace automation

---

# Step #1: RFQ Email Generator

## Problem

Creating RFQ emails manually is repetitive and time-consuming, especially when information is stored across multiple spreadsheets and formatting must remain consistent across vendors.

## Solution

A Google Apps Script solution that generates ready-to-send Gmail drafts with a single click.

The system:

- Extracts structured RFQ data from Google Sheets
- Pulls formatted content from a generator spreadsheet
- Converts sheet content into HTML email format
- Generates Gmail drafts using the Gmail API
- Links supporting Drive resources into the workflow

## Impact

- Eliminates manual RFQ drafting
- Ensures consistent communication formatting
- Reduces human error
- Centralizes RFQ generation into a single workflow

## Workflow

```text
Google Sheets (RFQ Input)
        ↓
Apps Script Engine
        ↓
Drive Generator Sheet
        ↓
HTML Email Builder
        ↓
Gmail Draft Creation
```

---

# Step #2: Vendor Folder Generator

## Problem

Vendor and project folder structures often require repetitive manual setup, resulting in:

- Inconsistent folder organization
- Administrative overhead
- Reduced traceability

## Solution

A Google Drive automation tool that:

- Reads vendor and bid information from Google Sheets
- Locates the correct project structure in Google Drive
- Automatically creates missing "Received" folders
- Generates vendor-specific folders
- Outputs generated folder links back into the spreadsheet

## Impact

- Standardized vendor folder structures
- Reduced manual setup work
- Faster RFQ onboarding
- Improved document traceability

## Workflow

```text
Google Sheets (Vendor + Bid Data)
        ↓
Apps Script Engine
        ↓
Google Drive Navigation Engine
        ↓
Folder Creation System
        ↓
Sheet Output (Folder URLs)
```

---

# Step #3: Save Quotes Automation

## Problem

Saving vendor quotes manually requires:

- Opening emails individually
- Downloading attachments
- Organizing files into project folders
- Tracking processing status

This process becomes inefficient and error-prone when handling large volumes of RFQs.

## Solution

A Gmail and Google Drive automation workflow that:

- Uses a Google Chrome bookmark tool to capture Gmail Message IDs
- Stores Message IDs and destination folder references in Google Sheets
- Retrieves Gmail messages automatically
- Converts email content into PDF format
- Extracts and saves supported attachments
- Stores all files in the designated Drive folder
- Updates processing status in Google Sheets
- Labels processed Gmail conversations for tracking purposes

## Impact

- Eliminates manual email archiving
- Standardizes quote storage
- Improves RFQ traceability
- Reduces administrative effort

## Workflow

```text
Chrome Bookmark Tool
        ↓
Google Sheets (Message ID + Folder Mapping)
        ↓
Apps Script Engine
        ↓
Gmail Message Retrieval
        ↓
PDF + Attachment Processing
        ↓
Google Drive Storage
        ↓
Status Logging + Gmail Labeling
```

---

# Step #4: Logged Sent Quotes Tracker

## Problem

Tracking sent RFQs manually requires repetitive work between Gmail and Google Drive, making it difficult to maintain a consistent audit trail.

## Solution

A batch-processing workflow that:

- Processes sent RFQ emails
- Applies Gmail labels for state management
- Logs processed RFQs into designated project folders
- Maintains an auditable record of RFQ communication

## Impact

- Improved auditability of communication history
- Reduced duplicate processing
- Better visibility into RFQ status
- Improved workflow accountability

## Workflow

```text
Sent RFQ Emails
        ↓
Apps Script Engine
        ↓
Gmail Label Management
        ↓
Drive Archiving
        ↓
Google Sheets Logging
```

---

# System Summary

This repository demonstrates how Google Workspace can be used as a lightweight workflow orchestration platform.

The automation suite replaces manual operational work associated with:

- RFQ generation
- Vendor folder management
- Quote archiving
- Communication tracking

while maintaining a structured, auditable, and scalable process entirely within Google Workspace.

---

# Business Impact

- Reduced repetitive administrative work
- Standardized RFQ and vendor management processes
- Improved document organization and traceability
- Increased visibility into procurement workflows
- Reduced risk of human error
- Leveraged existing Google Workspace infrastructure without requiring additional software
