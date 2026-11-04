@echo off
setlocal
echo ========================================================
echo               PdfRead - Build Pipeline
echo ========================================================
echo.

where dotnet >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERROR] .NET SDK is not found on your system PATH!
    echo Please install the .NET 8 SDK from: https://dotnet.microsoft.com/download/dotnet/8.0
    pause
    exit /b 1
)

echo Building PdfRead Desktop (.NET 8 WPF x64)...
dotnet publish "%~dp0PdfRead\PdfRead.csproj" -c Release -r win-x64 --self-contained false -o "%~dp0PdfRead\bin\Release\net8.0-windows\win-x64\publish"

if %ERRORLEVEL% neq 0 (
    echo [ERROR] Desktop build failed!
    pause
    exit /b 1
)

echo.
echo ========================================================
echo   Build Completed Successfully!
echo ========================================================
echo.
echo You can now run:
echo   - "Launch PdfRead.bat" (Desktop Viewer & Editor)
echo.
pause
