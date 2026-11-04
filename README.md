# PdfRead

<div align="center">

[![.NET 8](https://img.shields.io/badge/.NET-8.0_WPF-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com/)
[![WebView2](https://img.shields.io/badge/Microsoft-WebView2-0078D4?logo=microsoftedge&logoColor=white)](https://developer.microsoft.com/en-us/microsoft-edge/webview2/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](./LICENSE)

**A high-performance, document-grade PDF reader and editor designed for readability, in-place non-shifting text editing, and seamless bi-directional Microsoft Office integration.**

[Features](#features) • [Quick Start](#quick-start) • [Shortcuts](#keyboard-shortcuts) • [Architecture](#project-structure) • [Shell Integration](#windows-explorer-integration)

</div>

---

## Overview

**PdfRead** is a native Windows PDF viewer and editor written in **C# .NET 8 WPF** with **Microsoft WebView2**, **PdfSharp**, and **PdfPig**. It provides book-grade reading typography, tabbed document navigation, in-place non-shifting text block editing, and bi-directional Microsoft Word & Excel conversion.

---

## Features

- **Document-Grade Reading Typography**: Clean, comfortable formatting with the *Zinc Dark Theme* (charcoal/zinc aesthetic without glare).
- **Tabs & Single-Instance IPC**: Opening multiple PDF files from Windows Explorer organizes them neatly in tabs within the same active window.
- **In-Place Text Editing**:
  - Click on any text block on a PDF page to edit text directly where it is without shifting adjacent page layout.
  - Interactive bounding box with resize handles and font controls.
  - Persists edits cleanly back into standard PDF streams.
- **Bi-Directional Microsoft Office Conversion**:
  - **Office to PDF ("The Other Way")**: Right-click any Word (`.docx`, `.doc`) or Excel (`.xlsx`, `.xls`) file in Explorer and select **"Convert to PDF with PdfRead"** to convert headlessly via Microsoft Office COM Automation (`ExportAsFixedFormat`).
  - **PDF to Excel (`.xlsx`)**: Extracts coordinate-aligned table grids and writes structured `.xlsx` workbooks via `ClosedXML`.
  - **PDF to Word (`.docx`)**: Converts PDF pages into editable `.docx` files using Word reflow automation.
- **Page Management**: Reorganize pages with 1-click 90° clockwise rotation and page deletion.
- **In-Page Search (`Ctrl + F`)**: Real-time keyword search with match navigation.
- **Drag & Drop**: Drop PDFs, Word documents, or Excel spreadsheets directly into the window to open or convert automatically.
- **Settings & Shortcut Management**:
  - 1-Click Update Desktop Icon with the red PDF badge.
  - 1-Click Update Start Menu & Taskbar shortcut for instant pinning.
  - 1-Click Master Update for all system shortcuts and icons.
  - Automatic Windows Startup option (instant background listener).
  - Windows 11 direct first-click context menu toggle (no "Show more options" needed).
- **Native Windows Explorer Integration**:
  - Right-click `.pdf` -> **"Open with PdfRead"**.
  - Right-click `.docx` / `.xlsx` -> **"Convert to PDF with PdfRead"**.
  - Custom red badge icon with "PDF" / "P" branding.

---

## Quick Start

### Running PdfRead
1. Double-click **`Launch PdfRead.bat`** (or run `build.bat` once).  
   *If the application hasn't been built yet, the launcher will automatically compile it for you using the .NET 8 SDK.*

### Windows Explorer Setup
Inside the [`scripts/`](./scripts/) folder:
- **`Register Context Menu.bat`**: Registers **"Open with PdfRead"** for PDFs and **"Convert to PDF with PdfRead"** for Word & Excel in current-user registry (no admin rights required).
- **`Unregister Context Menu.bat`**: Cleanly removes all context menus and registry entries.
- **`Install Office Addin.bat`**: Sets up Office Ribbon integration.

---

## Keyboard Shortcuts

| Shortcut | Action |
| :--- | :--- |
| <kbd>Ctrl</kbd> + <kbd>O</kbd> | Open PDF or Office document |
| <kbd>Ctrl</kbd> + <kbd>S</kbd> | Save document / apply in-place text edits |
| <kbd>Ctrl</kbd> + <kbd>T</kbd> | Open new tab |
| <kbd>Ctrl</kbd> + <kbd>W</kbd> | Close active tab |
| <kbd>Ctrl</kbd> + <kbd>F</kbd> | In-document search |

---

## Project Structure

```
pdfread repo folder/
├── Launch PdfRead.bat        # 1-click desktop launcher (with auto-build)
├── build.bat                 # 1-click build pipeline for Desktop
├── README.md                 # Project documentation
│
├── PdfRead/                  # Desktop WPF Project (.NET 8)
│   ├── Assets/               # Red badge app icons, viewer styling
│   ├── Models/               # Document tab, text block, settings models
│   ├── Services/             # PDF editing, Office conversion, and Shell integration
│   │   ├── PdfEditorService.cs
│   │   ├── OfficeConversionService.cs
│   │   ├── DocumentExportService.cs
│   │   └── ShellIntegration.cs
│   ├── MainWindow.xaml(.cs)  # Primary window with tab management & WebView2
│   └── PdfRead.csproj        # Project configuration
│
└── scripts/                  # Shell integration & Context Menu scripts
    ├── Register Context Menu.bat
    ├── Unregister Context Menu.bat
    └── Install Office Addin.bat
```
