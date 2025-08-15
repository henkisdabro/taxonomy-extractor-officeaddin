@echo off
:: IPG Taxonomy Extractor - Simple Desktop Installation
:: Double-click this file to install the add-in

echo.
echo ╔═══════════════════════════════════════════════════════════════════════════════╗
echo ║                    IPG TAXONOMY EXTRACTOR - DESKTOP INSTALLER                ║
echo ╚═══════════════════════════════════════════════════════════════════════════════╝
echo.

:: Check if PowerShell is available
powershell -Command "Write-Host 'PowerShell detected successfully'" >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ ERROR: PowerShell not found or not accessible
    echo.
    echo This installer requires PowerShell to run.
    echo PowerShell is included with Windows 7 SP1 and later.
    echo.
    pause
    exit /b 1
)

:: Check if install script exists
if not exist "install-excel-desktop.ps1" (
    echo ❌ ERROR: install-excel-desktop.ps1 not found
    echo.
    echo Please ensure this batch file is in the same folder as:
    echo   - install-excel-desktop.ps1
    echo.
    pause
    exit /b 1
)

echo ⚡ Starting PowerShell installation script...
echo.
echo ℹ️  Note: You may see a security prompt - this is normal.
echo    Choose "Run once" or "Yes" to proceed with installation.
echo.

:: Set execution policy and run the PowerShell script
powershell -ExecutionPolicy Bypass -File "install-excel-desktop.ps1"

:: Check the result
if %errorlevel% equ 0 (
    echo.
    echo ✅ Installation process completed successfully!
) else (
    echo.
    echo ⚠️  Installation process completed with warnings or errors.
    echo    Check the output above for details.
)

echo.
echo Press any key to close this window...
pause >nul