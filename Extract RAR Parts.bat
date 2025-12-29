@echo off
setlocal

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"

REM Check if a file was dragged onto this batch file
if "%~1"=="" (
    REM No file dragged - run PowerShell script with file picker
    powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_DIR%extract-and-delete.ps1"
) else (
    REM File was dragged - pass it to the PowerShell script
    powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_DIR%extract-and-delete.ps1" "%~1"
)

REM Keep window open if there was an error
if errorlevel 1 (
    echo.
    echo Press any key to close...
    pause >nul
)

endlocal