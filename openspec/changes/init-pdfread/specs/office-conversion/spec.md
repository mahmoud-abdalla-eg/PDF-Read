# Spec Delta

## Purpose

Enables bi-directional conversions between PDF documents and Microsoft Office formats (Word and Excel), accessible via context menus, inside Office ribbons, and within the app.

## ADDED Requirements

### Requirement: Converting Office Documents to PDF
The system SHALL convert Microsoft Word and Excel documents to PDF format utilizing local Microsoft Office COM Automation.

#### Scenario: Converting Word document to PDF
- **WHEN** a `.docx` or `.doc` file path is supplied for conversion
- **THEN** the system launches a headless Word instance, executes ExportAsFixedFormat to PDF, and stores the resulting PDF file

#### Scenario: Converting Excel spreadsheet to PDF
- **WHEN** an `.xlsx` or `.xls` file path is supplied for conversion
- **THEN** the system launches a headless Excel instance, executes ExportAsFixedFormat to PDF, and stores the resulting PDF file

### Requirement: Windows Explorer Context Menu Conversion
The system SHALL allow users to convert Office files directly from Windows Explorer via a right-click context menu.

#### Scenario: Converting file from Explorer right-click
- **WHEN** the user right-clicks a Word or Excel file in Explorer and selects "Convert to PDF with PdfRead"
- **THEN** the system converts the file in the background and opens the resulting PDF document in a new tab in PdfRead

### Requirement: Office Ribbon Export Integration
The system SHALL provide an Office Ribbon integration inside Microsoft Excel and Word that triggers a one-click export to PdfRead.

#### Scenario: Clicking Ribbon export button in Excel or Word
- **WHEN** the user clicks the "Export to PdfRead" ribbon action inside an open Excel workbook or Word document
- **THEN** the active file is exported to PDF and opened directly inside PdfRead

### Requirement: Exporting PDF to Excel Spreadsheet
The system SHALL extract tabular structures from a PDF document and export them into an editable Excel spreadsheet (`.xlsx`).

#### Scenario: Exporting table data to Excel
- **WHEN** the user initiates an export to Excel for the current PDF document
- **THEN** the system clusters text coordinates into rows and columns and generates a formatted `.xlsx` workbook

### Requirement: Exporting PDF to Word Document
The system SHALL convert text, headers, and media elements from a PDF document into an editable Word document (`.docx`).

#### Scenario: Exporting document to Word
- **WHEN** the user initiates an export to Word for the current PDF document
- **THEN** the system reconstructs paragraphs and sections into a valid `.docx` document
