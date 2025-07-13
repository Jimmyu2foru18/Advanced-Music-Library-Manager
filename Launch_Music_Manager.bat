@echo off
REM Music Library Manager Launcher
REM This batch file launches the Music Library Manager GUI

echo ========================================
echo    Advanced Music Library Manager
echo ========================================
echo.
echo Starting the Music Library Manager GUI...
echo.

REM Check if PowerShell is available
powershell -Command "Write-Host 'PowerShell is available'" >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell is not available or not in PATH
    echo Please ensure PowerShell is installed and accessible
    pause
    exit /b 1
)

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"

REM Check if the GUI script exists
if not exist "%SCRIPT_DIR%MusicLibraryGUI.ps1" (
    echo ERROR: MusicLibraryGUI.ps1 not found in %SCRIPT_DIR%
    echo Please ensure all files are in the same directory
    pause
    exit /b 1
)

REM Launch the GUI application
echo Launching GUI application...
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT_DIR%MusicLibraryGUI.ps1"

REM Check if there was an error
if errorlevel 1 (
    echo.
    echo ERROR: Failed to launch the Music Library Manager
    echo This might be due to:
    echo - PowerShell execution policy restrictions
    echo - Missing dependencies
    echo - Corrupted script files
    echo.
    echo Try running this command manually:
    echo powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%MusicLibraryGUI.ps1"
    echo.
    pause
)

exit /b 0