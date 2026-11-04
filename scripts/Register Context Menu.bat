@echo off
setlocal
echo ========================================================
echo        PdfRead - Register Explorer Context Menus
echo ========================================================
echo.

set "EXE=%~dp0..\PdfRead\bin\Release\net8.0-windows\win-x64\publish\PdfRead.exe"
if not exist "%EXE%" (
    set "EXE=%~dp0..\PdfRead\bin\Debug\net8.0-windows\PdfRead.exe"
)
if not exist "%EXE%" (
    echo [INFO] PdfRead.exe not found. Building now...
    call "%~dp0..\build.bat"
    set "EXE=%~dp0..\PdfRead\bin\Release\net8.0-windows\win-x64\publish\PdfRead.exe"
)

if not exist "%EXE%" (
    echo [ERROR] Could not find PdfRead.exe. Build failed.
    pause
    exit /b 1
)

echo Registering context menus and desktop/taskbar shortcuts...
"%EXE%" --register
"%EXE%" --shortcut

echo.
echo [SUCCESS] Windows Explorer integration complete:
echo   - Right-click .pdf   -> "Open with PdfRead"
echo   - Right-click .docx  -> "Convert to PDF with PdfRead"
echo   - Right-click .xlsx  -> "Convert to PDF with PdfRead"
echo   - Red badge shortcut created on Desktop and Start Menu (ready for Taskbar pinning)
echo.
pause
