# Design: PdfRead Architecture & Implementation

## Context

`PdfRead` is a native Windows PDF viewer, editor, and conversion suite. It builds upon the proven desktop patterns of `MarkRead` (C# .NET 8 WPF, WebView2, named pipe IPC, Zinc reading theme, and registry shell integrations) while focusing entirely on PDF documents, in-place non-shifting text editing, and bi-directional Microsoft Office conversions. See `proposal.md` for motivation and background.

## Goals / Non-Goals

**Goals:**
* Deliver a lightweight, tabbed PDF reader in .NET 8 WPF with Zinc dark/light themes and single-instance IPC.
* Provide interactive in-place text editing with fixed positioning, resize handles, and stream redaction/replacement without layout shifting.
* Provide an annotation and markup layer (highlights, notes, ink, stamps) and page reorganizing.
* Implement bi-directional Office conversion:
  * Word and Excel to PDF via Explorer right-click context menu and Office Ribbon integration.
  * PDF to Word (`.docx`) and Excel (`.xlsx`) via structured coordinate and table extraction.
* Generate a red brand badge identity ("PDF" / "P") and Windows shell scripts.

**Non-Goals:**
* Markdown rendering or editing (MarkRead handles Markdown; PdfRead is exclusively for PDFs and document conversion).
* Cloud/online conversion APIs (all conversions run 100% locally and offline).

## Decisions

### 1. Viewer Host: .NET 8 WPF + Microsoft WebView2 hosting PDF.js
* **Decision**: Host Mozilla PDF.js inside a WPF `WebView2` control with a two-way C# <-> JavaScript JSON-RPC bridge.
* **Rationale**: PDF.js provides complete rendering parity across complex PDF specifications, supports continuous scrolling, thumbnails, and bookmark navigation, and allows complete CSS customization for the Zinc dark theme. It also natively exposes text layer bounding coordinates.
* **Alternatives Considered**:
  * *PdfiumViewer*: Lower memory overhead, but lacks modern HTML/CSS theming, requires custom WPF controls for text selection and search highlights, and complicates annotation overlays.

### 2. In-Place Text Editing via Bounding Box Overlays & Stream Patching
* **Decision**: Implement a two-phase in-place text editing system:
  1. *Interactive Overlay*: When Text Edit Mode is active, clicking on any text span creates an absolute-positioned overlay text box matching the exact coordinates `(x, y, w, h)`, font family, font size, and color. The box includes resize handles. The user modifies text or resizes without affecting surrounding document layout.
  2. *Stream Mutation*: On document save, the backend PDF engine (via `PdfSharp` / `PdfPig`) calculates the bounding box, draws a clean background redaction patch over the original vector glyphs, and renders the replacement text at the new dimensions using matching typography.
* **Rationale**: Raw PDF content stream editing on subsetted fonts frequently causes font corruption. Coordinate-based bounding-box patching guarantees stability and layout fidelity.
* **Alternatives Considered**:
  * *Global Text Reflow*: Breaks multi-column documents, tables, and graphic layouts.

### 3. Word/Excel to PDF Conversion via Local Office COM Interop
* **Decision**: Leverage local Microsoft Office Professional Plus 2019 COM Automation (`Word.Application` and `Excel.Application` invoking `ExportAsFixedFormat`).
* **Rationale**: Word and Excel COM automation provides 100% pixel-perfect conversion to PDF identical to Adobe Acrobat, handles macro-enabled and complex documents, and requires no external third-party commercial licenses.
* **Alternatives Considered**:
  * *LibreOffice Headless*: Requires a large 500MB separate installation.
  * *Open-source rendering engines*: Poor fidelity for complex Excel formulas and Word layouts.

### 4. Windows Explorer Shell & Office Ribbon Integration
* **Decision**:
  * *Explorer Right-Click*: Register `Software\Classes\SystemFileAssociations\.docx\shell\Convert to PDF with PdfRead` and `.xlsx\shell\Convert to PDF with PdfRead` in HKCU. Invokes `PdfRead.exe --convert-to-pdf "%1"`.
  * *Office Ribbon*: Provide an installable Office Ribbon customization / macro add-in template that places an "Export to PdfRead" button on the Word and Excel ribbons.
* **Rationale**: Direct integration matches user habits from Adobe Acrobat and WPS Office without requiring administrative privileges.

### 5. PDF to Excel/Word Export Pipeline
* **Decision**:
  * *To Excel (`.xlsx`)*: Use `PdfPig` to extract text coordinates, run spatial clustering algorithms to identify tabular rows and column boundaries, and write formatted spreadsheets using `ClosedXML`.
  * *To Word (`.docx`)*: Use Word's built-in PDF reflow engine via Word COM (`Documents.Open(pdfPath)`) to produce an editable `.docx`, with fallback to structured OpenXml generation.
* **Rationale**: Leverages the best available local engines to produce clean, editable documents.

## Risks / Trade-offs

* **[Risk: Office COM process lingering in background]** → **Mitigation**: Implement strict `try...finally` resource cleanup with explicit `Marshal.ReleaseComObject` and process termination safeguards.
* **[Risk: Text edit overflow beyond bounding box]** → **Mitigation**: Text overlay supports resize handles and font auto-scaling, allowing the user to expand the box or scale down text to prevent collision with adjacent elements.
* **[Risk: Large PDF memory footprint in WebView2]** → **Mitigation**: Enable PDF.js page virtualization (destroying off-screen rendered canvas viewports while maintaining text layer anchors).

## Migration Plan

1. Standalone launcher `Launch PdfRead.bat` auto-compiles with .NET 8 SDK if binaries are missing.
2. `scripts/Register Context Menu.bat` sets up HKCU shell keys without requiring administrator rights.
3. `scripts/Unregister Context Menu.bat` cleanly deletes registered shell keys.
