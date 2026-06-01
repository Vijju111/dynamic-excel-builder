# Software Requirements Specification (SRS)

## Dynamic Excel Builder & Live Auto-Fill System

**Version:** 1.0  
**Date:** 2026-06-01  
**Prepared For:** Project Stakeholders  
**Recommended Stack:** React + FastAPI, local browser storage for session persistence

---

## 1. Introduction

### 1.1 Purpose

This document defines the software requirements for a professional web application that enables users to upload a master Excel file, create customized working sheets, perform real-time lookup-based auto-fill operations, edit data in an Excel-like interface, and export the final output to Excel and PDF formats without data loss.

The system is intended for non-technical users who currently perform repetitive spreadsheet preparation manually using copy-paste, VLOOKUP/XLOOKUP, and multiple Excel files.

### 1.2 Scope

The application shall provide a browser-based workflow where users can:

1. Upload a master Excel file.
2. Detect and manage column headers.
3. Select, rename, reorder, and remove columns.
4. Add unlimited custom columns.
5. Select a lookup key column.
6. Generate a new working sheet.
7. Enter lookup values and receive instant auto-filled data.
8. Edit and manage sheet data in an Excel-like grid.
9. Preserve session edits using local storage.
10. Export the final sheet to `.xlsx` and `.pdf` formats.
11. Re-import previously generated sheets for continued editing.

### 1.3 Intended Users

- HR departments
- Schools and colleges
- Finance teams
- Inventory teams
- Administrative staff
- Data-entry operators
- Any non-technical user working with large Excel-based records

### 1.4 Definitions

| Term | Definition |
|---|---|
| Master Excel File | Original large Excel file containing complete source records. |
| Working Sheet / Sub-Sheet | Customized generated sheet containing selected columns and custom columns. |
| Lookup Column | Existing master column used as the matching key for auto-fill. |
| Custom Column | User-created column that does not exist in the master file. |
| Auto-Fill | Automatic population of matching row data after lookup value entry. |
| Local Storage | Browser-based storage used to preserve current session/project data. |

---

## 2. Overall Description

### 2.1 Product Perspective

The system shall be a standalone web application. It shall not require external cloud storage for core operation. Uploaded data and working sessions shall be stored locally in the browser unless optional future persistence is added.

Recommended architecture:

- **Frontend:** React
- **Grid Component:** AG Grid Enterprise, Handsontable, or equivalent Excel-like grid
- **Backend:** FastAPI
- **Excel Processing:** pandas, openpyxl, xlsxwriter
- **PDF Processing:** reportlab or equivalent
- **Storage:** Browser localStorage / IndexedDB for session persistence
- **Deployment:** Docker-compatible; supports Windows and Linux

### 2.2 Product Functions

The system shall provide the following major functions:

- Excel upload and validation
- Header row detection and correction
- Column selection and management
- Custom column creation
- Lookup key configuration
- Dynamic working sheet generation
- Real-time lookup and auto-fill
- Duplicate match handling
- Not-found handling
- Excel-like editing grid
- Local session auto-save
- Excel export
- PDF export
- Existing generated sheet import

### 2.3 User Characteristics

Users are expected to have basic Excel knowledge but may not have technical knowledge of databases, formulas, APIs, or programming.

The UI must be simple, guided, and error-resistant.

### 2.4 Constraints

- The system must support `.xlsx` and `.xls` files.
- Core session storage must use local browser storage.
- Uploaded files must not execute embedded code or macros.
- The application must handle at least 10,000 rows and 50+ columns.
- Lookup response should be under 200ms for 10,000+ rows after indexing.
- Password-protected files may be rejected unless password support is implemented.

### 2.5 Assumptions

- The first worksheet of the uploaded Excel file shall be used by default unless sheet selection is later implemented.
- The first row shall be assumed to contain headers by default.
- Users may override header row detection manually.
- Formula cells may be imported as calculated values by default, with optional formula preservation if implemented.

---

## 3. Functional Requirements

## 3.1 Master Excel Upload

### FR-001: Upload Excel File

The system shall allow users to upload master Excel files in the following formats:

- `.xlsx`
- `.xls`

### FR-002: Validate File Type

The system shall reject unsupported file types and display a clear message such as:

> Unsupported file type. Please upload a valid Excel file.

### FR-003: Detect Corrupted Files

The system shall detect corrupted or unreadable Excel files and display a user-friendly error message.

### FR-004: Handle Empty Excel Files

If the uploaded file contains no usable data, the system shall display:

> No data found.

### FR-005: Handle Password-Protected Files

If the uploaded file is password protected and password handling is not implemented, the system shall display:

> Password-protected files are not supported.

---

## 3.2 Column Header Detection

### FR-006: Automatic Header Detection

The system shall automatically detect column headers from the uploaded file. By default, the first row shall be treated as the header row.

### FR-007: Manual Header Row Selection

The system shall allow the user to select a different row as the header row when headers are not present in the first row.

Example: row 3 may be selected as the actual header row.

### FR-008: Manual Header Naming

If the uploaded file does not contain proper headers, the system shall allow the user to manually assign column names.

Example:

| Excel Column | User-Defined Name |
|---|---|
| Column A | Employee ID |
| Column B | Name |
| Column C | Department |

### FR-009: Duplicate Header Handling

If duplicate column names are detected, the system shall automatically rename them using suffixes.

Example:

- Name
- Name_1
- Name_2

---

## 3.3 Column Management

### FR-010: Display Available Columns

After successful upload and header detection, the system shall display all available columns from the master sheet.

### FR-011: Select Existing Columns

The system shall allow users to select one or more existing columns to include in the working sheet.

### FR-012: Reorder Columns

The system shall allow users to reorder selected columns before generating the working sheet.

### FR-013: Rename Columns

The system shall allow users to rename selected columns for the generated working sheet.

Example:

`Employee ID` may be renamed to `Emp Code`.

### FR-014: Remove Columns

The system shall allow users to remove selected columns from the working sheet configuration.

### FR-015: Empty Column Handling

The system shall detect empty columns and allow the user to remove them.

---

## 3.4 Custom Column Management

### FR-016: Add Custom Columns

The system shall allow users to add unlimited custom columns that do not exist in the master sheet.

Examples:

- Attendance
- Status
- Remarks
- Verification
- Comments

### FR-017: Rename Custom Columns

The system shall allow users to rename custom columns.

### FR-018: Delete Custom Columns

The system shall allow users to delete custom columns.

### FR-019: Reorder Custom Columns

The system shall allow users to reorder custom columns with existing selected columns.

### FR-020: Prevent Duplicate Custom Column Names

The system shall prevent creation of duplicate custom column names and display a clear validation message.

---

## 3.5 Lookup Column Configuration

### FR-021: Select Lookup Column

The system shall provide a dropdown for selecting the lookup column.

### FR-022: Restrict Lookup Column to Master Columns

Only columns that exist in the uploaded master sheet shall be available as lookup columns.

### FR-023: Missing Lookup Column Validation

The system shall block working sheet generation until a valid lookup column is selected.

### FR-024: Duplicate Lookup Warning

If the selected lookup column contains duplicate values, the system shall warn the user.

---

## 3.6 Dynamic Working Sheet Generation

### FR-025: Generate Working Sheet

The system shall generate a new empty working sheet based on:

- Selected master columns
- User-created custom columns
- Defined column order
- Renamed column labels

### FR-026: Preserve Column Mapping

The system shall internally preserve the mapping between renamed working sheet columns and original master columns.

Example:

| Master Column | Working Sheet Column |
|---|---|
| Employee ID | Emp Code |
| Department | Dept |

### FR-027: Initial Empty Rows

The generated working sheet shall initially contain empty editable rows or allow users to insert rows immediately.

---

## 3.7 Excel-Like Live Grid

### FR-028: Editable Grid

The generated sheet shall be displayed in a fully interactive spreadsheet-like grid.

### FR-029: Keyboard Navigation

The grid shall support keyboard navigation similar to Excel.

### FR-030: Copy and Paste

The grid shall support copy/paste, including multi-cell paste.

### FR-031: Multi-Cell Selection

The grid shall support multi-cell and range selection.

### FR-032: Row Insertion and Deletion

The grid shall allow users to insert and delete rows.

### FR-033: Column Resizing

The grid shall allow users to resize columns.

### FR-034: Column Reordering

The grid shall allow users to reorder columns within the working sheet.

### FR-035: Scrolling

The grid shall support horizontal and vertical scrolling for large sheets.

### FR-036: Undo and Redo

The grid shall support undo and redo operations.

### FR-037: Frozen Headers

The grid shall keep headers visible while scrolling.

### FR-038: Search and Filter

The grid shall provide search and filtering functionality.

---

## 3.8 Real-Time Auto-Fill Engine

### FR-039: Trigger Auto-Fill on Lookup Entry

When a user enters or changes a value in the lookup column, the system shall immediately search the master dataset.

### FR-040: Auto-Fill Matching Columns

If a matching record is found, the system shall populate all working sheet columns that map to columns in the master dataset.

### FR-041: Do Not Auto-Fill Custom Columns

The system shall never overwrite custom columns during auto-fill.

### FR-042: No Button Required

Auto-fill shall occur without requiring the user to click a button.

### FR-043: Optimized Lookup Index

The system shall create an in-memory dictionary/hash map index for lookup values to avoid scanning the full master sheet on every edit.

### FR-044: Lookup Performance

The auto-fill operation shall complete in less than 200ms for 10,000+ rows under normal operating conditions.

---

## 3.9 Matching Logic and Overwrite Protection

### FR-045: Match Only Shared Columns

Only columns that exist in both the master sheet and the working sheet mapping shall be auto-filled.

### FR-046: Protect Manual Edits

If the user manually edits an auto-filled value, the system shall not overwrite that value during subsequent auto-fill operations unless configured to do so.

### FR-047: Configurable Overwrite Behavior

The system shall support configurable overwrite modes:

1. Do not overwrite manually edited cells.
2. Overwrite after user confirmation.
3. Always overwrite.
4. Overwrite only when the row is refreshed.

### FR-048: Manual Refresh

The system shall allow users to manually refresh a row and accept updated master values.

### FR-049: Clear Row

The system shall allow users to clear a row, after which future lookup entry may re-trigger auto-fill.

---

## 3.10 Duplicate Lookup Handling

### FR-050: Duplicate Match Detection

If multiple master records match a lookup value, the system shall detect the duplicate match condition.

### FR-051: Duplicate Handling Options

The system shall support the following duplicate handling modes:

1. Use first match.
2. Use latest match.
3. Show selection popup.
4. Merge records, if applicable.

### FR-052: Duplicate Selection Popup

When configured to show a selection popup, the system shall display matching records and allow the user to choose one.

---

## 3.11 Not Found Handling

### FR-053: Preserve Lookup Value

If a lookup value is not found, the system shall keep the entered lookup value in the working sheet.

### FR-054: Leave Other Fields Blank

If no matching record is found, the system shall leave remaining auto-fill fields blank unless existing data must be preserved according to overwrite settings.

### FR-055: Highlight Not Found Row

The system shall visually highlight rows where lookup values are not found.

### FR-056: Display Not Found Message

The system shall display a status message such as:

> Record Not Found

---

## 3.12 Large Dataset Support

### FR-057: Support Large Files

The system shall support master datasets containing at least:

- 10,000+ rows
- 50+ columns

### FR-058: Avoid Full Scans During Editing

The system shall not perform full-sheet scans for every lookup edit. It shall use an indexed lookup structure.

### FR-059: Efficient Grid Rendering

The grid shall use virtualization or equivalent techniques to efficiently render large datasets.

---

## 3.13 Live Session Storage

### FR-060: Auto-Save Session

The system shall automatically save current working state during editing.

### FR-061: Browser Refresh Recovery

If the browser is refreshed, the system shall restore the latest saved session from local storage.

### FR-062: Temporary Workspace

The system shall maintain a temporary workspace containing:

- Master metadata
- Column configuration
- Working sheet rows
- Custom columns
- Lookup settings
- User edits

### FR-063: Optional Project State Save

The system may allow users to save and reload complete project state files.

### FR-064: Accidental Close Warning

The system shall prompt users before leaving the page if there are unsaved or active session changes.

---

## 3.14 Import Existing Generated Sheet

### FR-065: Import Previously Exported Sheet

The system shall allow users to import a previously exported working sheet.

### FR-066: Restore Structure

The system shall restore column structure, ordering, custom columns, and editable data where possible.

### FR-067: Continue Editing

After import, the user shall be able to continue editing and exporting the sheet.

---

## 3.15 Excel Export

### FR-068: Export to XLSX

The system shall export the final working sheet as an `.xlsx` file.

### FR-069: User-Defined Excel Filename

The user shall be able to enter the desired export filename.

Example:

`Attendance_June.xlsx`

### FR-070: Preserve Export Data

The exported Excel file shall preserve:

- All row data
- Column order
- Renamed columns
- Custom columns
- Unicode text
- Long text
- Mixed data types

### FR-071: Preserve Formatting

The exported Excel file should preserve basic formatting such as:

- Header style
- Column widths
- Text wrapping
- Date/number formatting where possible

### FR-072: No Data Truncation

The Excel export shall not truncate cell values.

---

## 3.16 PDF Export

### FR-073: Export to PDF

The system shall export the final working sheet as a `.pdf` file.

### FR-074: User-Defined PDF Filename

The user shall be able to enter the desired PDF export filename.

Example:

`Attendance_June.pdf`

### FR-075: Landscape Mode

The PDF export shall support landscape orientation.

### FR-076: Auto Page Breaks

The PDF export shall automatically handle page breaks for large tables.

### FR-077: Repeat Headers

The PDF export shall repeat table headers on new pages.

### FR-078: Column Alignment

The PDF export shall maintain proper column alignment.

### FR-079: Unicode Support

The PDF export shall support Unicode text, including Indian and international scripts where font support is available.

### FR-080: Large Table Support

The PDF export shall handle large working sheets without data loss.

---

## 3.17 Data Type and Content Handling

### FR-081: Unicode Support

The system shall support Unicode text including, but not limited to:

- Telugu
- Hindi
- Tamil
- Arabic
- Chinese

### FR-082: Very Long Text

The system shall support very long text values and wrap them correctly in grid, Excel export, and PDF export where practical.

### FR-083: Mixed Data Types

The system shall support the following data types:

- Text
- Number
- Date
- Currency
- Boolean

### FR-084: Formula Cells

The system shall provide an option to import formula cells as:

1. Calculated values; or
2. Preserved formulas, if implemented.

---

## 4. Non-Functional Requirements

## 4.1 Performance

### NFR-001: Lookup Response Time

Lookup-based auto-fill shall complete within 200ms for 10,000+ rows under normal usage conditions.

### NFR-002: Large Dataset Rendering

The application shall remain responsive when handling at least 10,000 rows and 50+ columns.

### NFR-003: Efficient Upload Processing

Excel upload and parsing shall provide progress or loading feedback for large files.

---

## 4.2 Security

### NFR-004: No Code Execution

The system shall not execute macros, scripts, formulas as code, or embedded objects from uploaded files.

### NFR-005: Upload Validation

Uploaded files shall be validated before processing.

### NFR-006: Content Sanitization

Cell values and file contents shall be sanitized before rendering to prevent script injection or unsafe HTML execution.

### NFR-007: Local-Only Core Storage

Core session data shall be stored locally in the browser and shall not depend on third-party storage platforms.

---

## 4.3 Reliability

### NFR-008: No Data Loss During Editing

The system shall preserve user edits during active editing and auto-save operations.

### NFR-009: No Data Loss During Export

The system shall export all visible working sheet data without loss.

### NFR-010: Recovery After Refresh

The system shall recover the latest locally saved session after browser refresh.

---

## 4.4 Usability

### NFR-011: Excel-Like Experience

The interface shall feel similar to Microsoft Excel, including keyboard navigation, editing, copy/paste, and row/column operations.

### NFR-012: Non-Technical User Friendly

The application shall use clear labels, guided steps, and understandable error messages.

### NFR-013: Responsive UI

The application shall be responsive and usable on common desktop and laptop screen sizes.

### NFR-014: No Unnecessary Page Refreshes

The application shall function as a modern single-page application with no unnecessary full-page refreshes.

---

## 4.5 Compatibility

### NFR-015: Browser Compatibility

The application should support current versions of:

- Google Chrome
- Microsoft Edge
- Mozilla Firefox

### NFR-016: Operating System Compatibility

The application shall be deployable on:

- Windows
- Linux

### NFR-017: Docker Compatibility

The application shall be Docker-compatible.

---

## 5. System Architecture

### 5.1 Recommended Architecture

The recommended implementation is a React frontend with a FastAPI backend.

```text
User Browser
   |
   |-- React UI
   |-- Excel-like Grid
   |-- LocalStorage / IndexedDB
   |
FastAPI Backend
   |-- Excel parser
   |-- Excel exporter
   |-- PDF exporter
   |-- Validation services
```

### 5.2 Frontend Responsibilities

The frontend shall handle:

- User interface
- File upload interaction
- Column configuration screens
- Excel-like grid rendering
- Real-time editing events
- Lookup triggering
- In-memory lookup index, if implemented client-side
- Local storage auto-save
- Export request initiation

### 5.3 Backend Responsibilities

The backend shall handle:

- Excel validation and parsing
- Header row extraction
- Optional formula handling
- Excel export generation
- PDF export generation
- Error normalization

### 5.4 Local Storage Requirement

The system shall use local browser storage for active session persistence. For large data, IndexedDB is recommended over localStorage because it supports larger structured datasets and better performance.

---

## 6. Data Model

### 6.1 Master Dataset Metadata

```json
{
  "fileName": "master.xlsx",
  "sheetName": "Sheet1",
  "headerRowIndex": 0,
  "columns": [
    {
      "id": "employee_id",
      "originalName": "Employee ID",
      "displayName": "Employee ID",
      "dataType": "text"
    }
  ]
}
```

### 6.2 Working Sheet Column Model

```json
{
  "id": "emp_code",
  "displayName": "Emp Code",
  "sourceColumn": "Employee ID",
  "isCustom": false,
  "order": 1
}
```

### 6.3 Custom Column Model

```json
{
  "id": "attendance",
  "displayName": "Attendance",
  "sourceColumn": null,
  "isCustom": true,
  "order": 4
}
```

### 6.4 Working Row Model

```json
{
  "rowId": "row_001",
  "values": {
    "emp_code": "1002",
    "name": "Alice",
    "department": "IT",
    "attendance": "Present"
  },
  "cellMeta": {
    "name": {
      "autoFilled": true,
      "manuallyEdited": false
    }
  },
  "status": "matched"
}
```

### 6.5 Lookup Index Model

```json
{
  "1002": [
    {
      "Employee ID": "1002",
      "Name": "Alice",
      "Department": "IT"
    }
  ]
}
```

---

## 7. User Workflow

### 7.1 Standard Workflow

1. User opens the application.
2. User uploads a master Excel file.
3. System validates and parses the file.
4. System detects headers.
5. User confirms or corrects headers.
6. User selects required columns.
7. User renames/reorders/removes columns if needed.
8. User adds custom columns.
9. User selects lookup column.
10. System warns about duplicates if applicable.
11. User generates working sheet.
12. User enters lookup values in the grid.
13. System auto-fills matching columns instantly.
14. User edits custom or required fields.
15. System auto-saves session locally.
16. User exports Excel and/or PDF.

---

## 8. Validation Rules

| Validation Area | Rule |
|---|---|
| File type | Only `.xlsx` and `.xls` accepted. |
| Empty file | Show `No data found.` |
| Corrupted file | Show readable error message. |
| Header row | Must contain at least one non-empty column name. |
| Duplicate headers | Auto-rename with suffix. |
| Custom columns | Duplicate custom names not allowed. |
| Lookup column | Must be selected before generation. |
| Lookup column source | Must exist in master columns. |
| Unsupported protected file | Show password-protected warning. |
| Not found lookup | Preserve entered value and highlight row. |

---

## 9. Edge Case Handling

| Edge Case | Expected Behavior |
|---|---|
| Empty Excel file | Display `No data found.` |
| Empty rows | Ignore automatically. |
| Empty columns | Allow removal. |
| Duplicate column names | Rename as `Name`, `Name_1`, `Name_2`. |
| Special characters | Preserve and render Unicode correctly. |
| Very long text | Wrap, do not truncate. |
| Mixed data types | Preserve text, number, date, currency, boolean where possible. |
| Formula cells | Import calculated values or preserve formulas based on user option. |
| Protected Excel files | Reject or request password if implemented. |
| Missing lookup column | Block sheet generation. |
| Duplicate custom column names | Prevent creation. |
| Accidental browser close | Prompt user before leaving. |

---

## 10. Export Requirements

### 10.1 Excel Export Acceptance Criteria

The Excel export shall be considered successful when:

- File downloads with the user-selected filename.
- All rows and columns are present.
- Column order matches the grid.
- Custom columns are included.
- Unicode text is preserved.
- Long text is not truncated.
- Basic formatting is preserved.

### 10.2 PDF Export Acceptance Criteria

The PDF export shall be considered successful when:

- File downloads with the user-selected filename.
- Table is readable in landscape mode.
- Headers repeat on each page.
- Page breaks are handled automatically.
- Columns remain aligned.
- Unicode text is preserved where font support allows.
- No visible data loss occurs.

---

## 11. Acceptance Criteria

The project shall be considered complete when a user can successfully:

1. Upload a valid master Excel file.
2. Detect or manually define headers.
3. Select required columns.
4. Reorder and rename selected columns.
5. Add custom columns.
6. Select a lookup column.
7. Generate an editable working sheet.
8. Enter a lookup value.
9. See matching master data auto-filled instantly.
10. Prevent accidental overwrite of manual edits.
11. Handle duplicate and missing lookup values safely.
12. Edit data in an Excel-like grid.
13. Preserve work after browser refresh.
14. Export the final data to Excel.
15. Export the final data to PDF.
16. Complete the workflow without data loss.

---

## 12. Suggested Development Phases

### Phase 1: Core Upload and Parsing

- File upload
- Excel validation
- Header detection
- Header correction
- Master data preview

### Phase 2: Column Configuration

- Column selection
- Rename/reorder/remove columns
- Custom column management
- Lookup column selection

### Phase 3: Live Grid and Auto-Fill

- Generate working sheet
- Excel-like grid integration
- Lookup indexing
- Real-time auto-fill
- Duplicate and not-found handling

### Phase 4: Session Persistence

- Local auto-save
- Browser refresh recovery
- Leave-page warning
- Existing generated sheet import

### Phase 5: Export

- Excel export
- PDF export
- Formatting and Unicode support
- Large-table testing

### Phase 6: Hardening and Deployment

- Performance optimization
- Edge-case handling
- Security validation
- Docker packaging
- Windows/Linux testing

---

## 13. Risks and Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| Very large Excel files may slow browser | High | Use backend parsing and grid virtualization. |
| localStorage size limits | High | Use IndexedDB for large session data. |
| PDF Unicode font issues | Medium | Bundle Unicode-compatible fonts. |
| Duplicate lookup ambiguity | Medium | Provide configurable handling modes. |
| Manual edit overwrite risk | High | Track cell metadata and require confirmation. |
| Formula preservation complexity | Medium | Start with calculated values, add formula mode later. |

---

## 14. Out of Scope for Initial Version

The following may be deferred unless explicitly required:

- Multi-user collaboration
- Cloud database storage
- Authentication and user accounts
- Real-time server sync
- Macro execution
- Advanced Excel formatting parity
- Complex formula recalculation engine
- Role-based permissions

---

## 15. Final Success Statement

The application succeeds when non-technical users can upload a master Excel file, configure a customized working sheet, type lookup values, receive instant auto-filled data, safely edit additional fields, preserve their work locally, and export clean Excel/PDF files with no data loss and an experience similar to Microsoft Excel.
