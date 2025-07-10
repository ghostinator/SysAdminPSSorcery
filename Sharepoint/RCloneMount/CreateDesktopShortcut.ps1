#Requires -Version 5.1

<#
.SYNOPSIS
    Creates a desktop shortcut for the SharePoint/OneDrive Mount Tool
.DESCRIPTION
    This script creates a desktop shortcut that launches the SharePoint/OneDrive GUI tool
.EXAMPLE
    .\CreateDesktopShortcut.ps1
#>

try {
    # Get the current script directory
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    
    # Define paths
    $targetPath = Join-Path $scriptDir "LaunchGUI.bat"
    $iconPath = Join-Path $scriptDir "SharePointOneDriveGUI.ps1"
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath "SharePoint OneDrive Mount Tool.lnk"
    
    # Check if the target file exists
    if (-not (Test-Path $targetPath)) {
        Write-Error "LaunchGUI.bat not found in the script directory: $scriptDir"
        Write-Host "Please ensure all files are in the same directory." -ForegroundColor Yellow
        exit 1
    }
    
    # Create the shortcut
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = $targetPath
    $Shortcut.WorkingDirectory = $scriptDir
    $Shortcut.Description = "SharePoint/OneDrive Mount Tool - Easy mounting of Microsoft 365 folders"
    $Shortcut.WindowStyle = 1  # Normal window
    
    # Try to set an icon (use PowerShell icon as fallback)
    try {
        $Shortcut.IconLocation = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe,0"
    }
    catch {
        # Fallback to default icon
    }
    
    # Save the shortcut
    $Shortcut.Save()
    
    # Release COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
    
    Write-Host "Desktop shortcut created successfully!" -ForegroundColor Green
    Write-Host "Shortcut location: $shortcutPath" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "You can now double-click the shortcut on your desktop to launch the SharePoint/OneDrive Mount Tool." -ForegroundColor Yellow
}
catch {
    Write-Error "Failed to create desktop shortcut: $($_.Exception.Message)"
    Write-Host "You can manually create a shortcut to LaunchGUI.bat if needed." -ForegroundColor Yellow
    exit 1
}