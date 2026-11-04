# Spec Delta

## Purpose

Provides native Windows integration including red-themed brand identity, desktop launchers, file associations, and Explorer right-click context menu registration.

## ADDED Requirements

### Requirement: Application Identity and Red Logo Badge
The system SHALL feature a distinct red brand badge icon displaying "PDF" (or compact "P" for micro-resolutions) matching the MarkRead geometric badge aesthetic.

#### Scenario: Displaying application icon
- **WHEN** the application window, taskbar icon, or executable file is viewed in Windows
- **THEN** the system displays the high-DPI multi-resolution `.ico` featuring the red badge identity

### Requirement: Windows Explorer Context Menus
The system SHALL register context menu handlers in the current user Windows Registry for `.pdf` documents and Office files.

#### Scenario: Right-clicking a PDF file
- **WHEN** a user right-clicks a `.pdf` file in Windows Explorer
- **THEN** the menu displays "Open with PdfRead" and clicking it launches or passes the path to PdfRead

#### Scenario: Right-clicking an Office document
- **WHEN** a user right-clicks a `.docx` or `.xlsx` file in Windows Explorer
- **THEN** the menu displays "Convert to PDF with PdfRead" and clicking it starts conversion

### Requirement: Shell Scripts and Launch Pipeline
The system SHALL provide self-contained batch scripts for one-click launching, compiling, and shell registry management.

#### Scenario: Launching via batch launcher
- **WHEN** the user executes `Launch PdfRead.bat`
- **THEN** the script checks if the binary is compiled, invokes `dotnet build` if missing, and starts PdfRead
