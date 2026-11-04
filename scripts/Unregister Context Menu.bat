@echo off
setlocal
echo ========================================================
echo       PdfRead - Unregister Explorer Context Menus
echo ========================================================
echo.

set "EXE=%~dp0..\PdfRead\bin\Release\net8.0-windows\win-x64\publish\PdfRead.exe"
if not exist "%EXE%" (
    set "EXE=%~dp0..\PdfRead\bin\Debug\net8.0-windows\PdfRead.exe"
)

if exist "%EXE%" (
    "%EXE%" --unregister
) else (
    reg delete "HKCU\Software\Classes\SystemFileAssociations\.pdf\shell\Open with PdfRead" /f >nul 2>nul
    reg delete "HKCU\Software\Classes\SystemFileAssociations\.docx\shell\Convert to PDF with PdfRead" /f >nul 2>nul
    reg delete "HKCU\Software\Classes\SystemFileAssociations\.xlsx\shell\Convert to PDF with PdfRead" /f >nul 2>nul
    reg delete "HKCU\Software\Classes\PdfRead.Document" /f >nul 2>nul
)

echo.
echo [SUCCESS] Windows Explorer context menus removed cleanly.
echo.
pause
