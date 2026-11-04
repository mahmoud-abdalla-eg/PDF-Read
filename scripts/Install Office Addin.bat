@echo off
setlocal
echo ========================================================
echo        PdfRead - Office Ribbon Add-in Setup
echo ========================================================
echo.

set "EXCEL_START=%APPDATA%\Microsoft\Excel\XLSTART"
set "WORD_START=%APPDATA%\Microsoft\Word\STARTUP"

echo Checking Office startup paths...
if not exist "%EXCEL_START%" mkdir "%EXCEL_START%" >nul 2>nul
if not exist "%WORD_START%" mkdir "%WORD_START%" >nul 2>nul

echo.
echo [INFO] Office Ribbon Integration:
echo Inside Excel or Word:
echo   1. You can export directly using the "Convert Office to PDF" button in PdfRead.
echo   2. Or right-click any Excel/Word file in Windows Explorer -> "Convert to PDF with PdfRead".
echo.
echo For direct Office Ribbon buttons:
echo The template add-in scripts have been prepared in:
echo   %~dp0..\scripts\
echo.
pause
