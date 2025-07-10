#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Removes OneDrive from startup locations for all existing user profiles
.DESCRIPTION
    This script removes OneDrive startup entries from both registry Run keys and startup folders
    for all existing user profiles on the system. Must be run as Administrator.
.NOTES
    Run as Administrator to access all user profiles and registry hives
#>

[CmdletBinding()]
param()

# Configuration
$ErrorActionPreference = "Continue"
$LogPath = "$env:TEMP"
$LogFile = "$LogPath\OneDriveStartupRemoval_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

function Write-Log {
    param([string]$Message)
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Time $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
    Write-Host $Message
}

Write-Log "Starting OneDrive startup removal for all existing users..."

# Get all user profiles with NTUSER.DAT files (existing users)
$userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object { 
    Test-Path "$($_.FullName)\NTUSER.DAT" -and 
    $_.Name -notin @('Default', 'Public', 'All Users', 'Default User')
}

Write-Log "Found $($userProfiles.Count) user profiles to process"

$totalRemoved = 0

foreach ($profile in $userProfiles) {
    $userName = $profile.Name
    $userPath = $profile.FullName
    Write-Log "Processing user: $userName"
    
    # --- A. Remove from Registry (HKCU Run Key) ---
    try {
        $ntUserDat = "$userPath\NTUSER.DAT"
        $tempHive = "HKU\TempHive_$userName"
        
        # Load the user's registry hive
        $loadResult = reg load $tempHive $ntUserDat 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Loaded registry hive for $userName"
            
            # Check if OneDrive entry exists in Run key
            $runKeyPath = "Registry::$tempHive\Software\Microsoft\Windows\CurrentVersion\Run"
            if (Test-Path $runKeyPath) {
                $runKey = Get-ItemProperty -Path $runKeyPath -ErrorAction SilentlyContinue
                if ($runKey -and $runKey.PSObject.Properties.Name -contains "OneDrive") {
                    Remove-ItemProperty -Path $runKeyPath -Name "OneDrive" -ErrorAction Stop
                    Write-Log "Removed OneDrive from Run registry key for $userName"
                    $totalRemoved++
                } else {
                    Write-Log "OneDrive not found in Run registry key for $userName"
                }
            }
            
            # Unload the registry hive
            reg unload $tempHive | Out-Null
            Write-Log "Unloaded registry hive for $userName"
        } else {
            Write-Log "Failed to load registry hive for $userName`: $loadResult"
        }
    } catch {
        Write-Log "Error processing registry for $userName`: $($_.Exception.Message)"
        # Attempt to unload hive if it was loaded
        reg unload $tempHive 2>&1 | Out-Null
    }
    
    # --- B. Remove from Startup Folders ---
    
    # Remove from user's personal startup folder
    $userStartupFolder = "$userPath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    $userOneDriveLnk = Join-Path $userStartupFolder "OneDrive.lnk"
    
    if (Test-Path $userOneDriveLnk) {
        try {
            Remove-Item $userOneDriveLnk -Force
            Write-Log "Removed OneDrive shortcut from startup folder for $userName"
            $totalRemoved++
        } catch {
            Write-Log "Failed to remove OneDrive shortcut from startup folder for $userName`: $($_.Exception.Message)"
        }
    } else {
        Write-Log "OneDrive shortcut not found in startup folder for $userName"
    }
}

# Remove from All Users Startup Folder
$allUsersStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
$allUsersOneDriveLnk = Join-Path $allUsersStartup "OneDrive.lnk"

if (Test-Path $allUsersOneDriveLnk) {
    try {
        Remove-Item $allUsersOneDriveLnk -Force
        Write-Log "Removed OneDrive shortcut from All Users startup folder"
        $totalRemoved++
    } catch {
        Write-Log "Failed to remove OneDrive shortcut from All Users startup folder: $($_.Exception.Message)"
    }
} else {
    Write-Log "OneDrive shortcut not found in All Users startup folder"
}

# Summary
Write-Log "OneDrive startup removal completed for all existing users."
Write-Log "Total startup entries removed: $totalRemoved"
Write-Log "Log file saved to: $LogFile"

if ($totalRemoved -gt 0) {
    Write-Host "Successfully removed $totalRemoved OneDrive startup entries. See log: $LogFile" -ForegroundColor Green
    exit 0
} else {
    Write-Host "No OneDrive startup entries found to remove. See log: $LogFile" -ForegroundColor Yellow
    exit 0
}
