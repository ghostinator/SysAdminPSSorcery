@echo off
title SharePoint/OneDrive Mount Tool
echo Starting SharePoint/OneDrive Mount Tool...
echo.

REM Check if PowerShell is available
powershell -Command "exit" >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell is not available or not in PATH
    echo Please ensure PowerShell is installed and accessible
    pause
    exit /b 1
)

REM Run the GUI script
powershell -ExecutionPolicy Bypass -File "%~dp0SharePointOneDriveGUI.ps1"

REM Check if the script ran successfully
if errorlevel 1 (
    echo.
    echo The script encountered an error.
    echo Please check the error messages above.
    pause
)