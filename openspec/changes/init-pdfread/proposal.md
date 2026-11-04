# Proposal: Initial PdfRead Desktop Application

## Why

Users who appreciate MarkRead's fast, document-grade reading typography, tabbed interface, and native Windows integration currently lack a dedicated reader for PDF documents that provides the same clean, responsive experience without Markdown bloat. Furthermore, working with PDFs often requires two essential workflows that are fragmented across heavy or commercial tools: (1) quickly editing text directly in-place without shifting the surrounding layout, and (2) seamless bi-directional conversion between PDFs and Microsoft Office formats (Word and Excel), accessible directly from Windows Explorer and inside Office itself. 

`PdfRead` solves this by delivering a lightweight, native Windows PDF reader and editor built on .NET 8 WPF and WebView2, featuring MarkRead's Zinc aesthetic, in-place non-shifting text editing, markup annotations, and native bi-directional Word/Excel conversion.

## What Changes

* **Native WPF Desktop Shell**: Tabbed document interface (TDI), single-instance IPC (opening files from Explorer adds tabs to the existing window), and Zinc dark/light theme designed for long-form reading.
* **Red App Identity & Logo**: Custom red badge logo icon depicting "PDF" (or "P" badge fitting icon dimensions), matching MarkRead's aesthetic badge style, with desktop and shell launcher scripts.
* **High-Fidelity PDF Viewer**: WebView2 hosting PDF.js with continuous scrolling, zooming, page thumbnail sidebar, outline/bookmark navigation, and instant in-page text search (`Ctrl + F`).
* **Two-Level PDF Editing**:
  * **Level 1 (In-Place Text Editing)**: Interactive bounding-box text selection and modification directly on the page without shifting the surrounding layout, supporting font sizing, text replacement, and container resizing.
  * **Level 2 (Annotations & Markup)**: Highlighting, sticky notes, freehand drawing, stamps, text box overlays, and page-level organization (reorder, rotate, delete, insert pages).
* **Bi-Directional Office Conversion**:
  * **Office -> PDF ("The Other Way")**: Right-click context menu in Windows Explorer for `.docx` and `.xlsx` ("Convert to PDF with PdfRead") utilizing local Microsoft Office COM Automation (`ExportAsFixedFormat`), plus Excel and Word Ribbon export integration.
  * **PDF -> Office (Word & Excel)**: Exporting PDF tables into structured `.xlsx` spreadsheets and reflowable text/layouts into `.docx` documents.
* **Native Windows Integration**: Registry-based file associations (`.pdf`), right-click shell menus, and standalone launchers (`Launch PdfRead.bat`, `build.bat`).

## Capabilities

### New Capabilities
- `pdf-viewing`: High-performance PDF document rendering, navigation (thumbnails, bookmarks), zoom, tabs, single-instance IPC, in-page search, and Zinc reading themes.
- `pdf-editing`: Two-tier editing: in-place non-shifting text block editing with resize handles, and an annotation/markup layer with page manipulation.
- `office-conversion`: Bi-directional conversion between PDF and Office formats (Word/Excel to PDF via Explorer context menu and Office Ribbon; PDF to Word/Excel).
- `shell-integration`: Windows Explorer context menus, file associations, desktop shortcuts, and red application identity/icon assets.

### Modified Capabilities
*(None - Greenfield project)*

## Impact

* **New Codebase**: Initializes `PdfRead` in `E:\md file viewer\pdfread repo folder\PdfRead` targeting `net8.0-windows`.
* **Dependencies**: Microsoft.Web.WebView2, PDFsharp / PdfPig for PDF coordinate extraction and stream updates, ClosedXML for Excel generation, DocumentFormat.OpenXml for Word generation, and Microsoft Office Interop for background conversion.
* **Shell Scripts**: New standalone scripts in `scripts/` and root launchers (`Launch PdfRead.bat`, `build.bat`).
