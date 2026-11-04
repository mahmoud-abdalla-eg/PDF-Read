# Tasks: Initialize PdfRead Desktop Application

## 1. Project Scaffolding & Brand Identity

- [x] 1.1 Create the .NET 8 WPF project structure (`PdfRead.csproj`, `App.xaml`, directory layout) and verify compilation with `dotnet build`
- [x] 1.2 Design and generate the red badge application icon (`app.ico`) displaying "PDF" (or compact "P" for micro-resolutions) matching the MarkRead geometric badge aesthetic, and verify assets in `Assets/`
- [x] 1.3 Create standalone Windows launcher and build batch scripts (`Launch PdfRead.bat`, `build.bat`) and verify automated build and launch behavior

## 2. Core Viewer & UI Shell

- [x] 2.1 Integrate Microsoft WebView2 with bundled PDF.js viewer assets, reader script bridge, and Zinc dark/light theme stylesheets, and verify viewer initialization
- [x] 2.2 Implement Tabbed Document Interface (TDI) and tab management in `MainWindow.xaml(.cs)` and verify multiple tabs can be opened and closed independently
- [x] 2.3 Implement single-instance NamedPipe IPC in `App.xaml.cs` to route external file openings into tabs of the active window, and verify new file arguments attach to active instance
- [x] 2.4 Implement in-page text search (`Ctrl + F`), zoom presets, continuous scroll, thumbnail sidebar, and bookmark outline navigation, and verify navigation controls

## 3. In-Place Text Editing & Annotations

- [x] 3.1 Implement text layer coordinate extraction through the PDF.js bridge to detect text block bounding boxes `(x, y, w, h)`, font family, size, and baseline
- [x] 3.2 Build the interactive in-place text editor overlay with resize handles, font styling controls, and text input without shifting surrounding page layout
- [x] 3.3 Implement PDF stream patcher service to redact original vector glyph regions and render replacement text into the PDF output on document save
- [x] 3.4 Implement markup annotations (highlights, sticky notes, freehand ink, stamps) and page operations (rotate, reorder, delete), and verify changes persist to file

## 4. Bi-Directional Office Conversions & Shell Integration

- [x] 4.1 Implement `OfficeConversionService` utilizing local Microsoft Word and Excel COM Automation (`ExportAsFixedFormat`) to convert `.docx` and `.xlsx` to PDF headlessly
- [x] 4.2 Implement Explorer right-click shell context menus ("Convert to PDF with PdfRead" and "Open with PdfRead") in `ShellIntegration.cs` and verify registration scripts
- [x] 4.3 Implement Office Ribbon export integration for Microsoft Excel and Word to enable one-click export into PdfRead
- [x] 4.4 Implement `DocumentExportService` to extract PDF tables into structured `.xlsx` spreadsheets (via ClosedXML) and document content into `.docx` files

## 5. End-to-End Verification & Documentation

- [x] 5.1 Perform comprehensive end-to-end testing of PDF reading, in-place text block editing, annotations, and Word/Excel bi-directional conversions
- [x] 5.2 Create project documentation in `README.md` detailing architecture, keyboard shortcuts, build instructions, and shell integration steps
