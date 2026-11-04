@echo off
setlocal
set "EXE=%~dp0PdfRead\bin\Release\net8.0-windows\win-x64\publish\PdfRead.exe"

if not exist "%EXE%" (
    set "EXE=%~dp0PdfRead\bin\Release\net8.0-windows\PdfRead.exe"
)
if not exist "%EXE%" (
    set "EXE=%~dp0PdfRead\bin\Debug\net8.0-windows\PdfRead.exe"
)
if not exist "%EXE%" (
    echo [INFO] PdfRead executable not found. Compiling with build.bat...
    call "%~dp0build.bat"
    set "EXE=%~dp0PdfRead\bin\Release\net8.0-windows\win-x64\publish\PdfRead.exe"
)

if exist "%EXE%" (
    start "" "%EXE%" %*
) else (
    echo [ERROR] Could not find or build PdfRead.exe.
    echo Please make sure .NET 8 SDK is installed.
    pause
)
