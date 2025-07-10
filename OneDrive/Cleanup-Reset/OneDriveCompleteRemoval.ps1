#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Complete OneDrive removal script with comprehensive logging
.DESCRIPTION
    This script performs all OneDrive removal functions:
    - Uninstalls OneDrive system-wide
    - Removes OneDrive from startup for all existing users
    - Prevents OneDrive setup for new users
    - Optionally removes OneDrive data and credentials
.NOTES
    Must be run as Administrator
    Designed for deployment via Microsoft Intune or local execution
#>

[CmdletBinding()]
param(
    [switch]$RemoveUserData = $false  # Set to $true to also remove user OneDrive folders
)

# Configuration
$ErrorActionPreference = "Continue"
$LogPath = "$env:TEMP"
$LogFile = "$LogPath\OneDriveCompleteRemoval_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

# Initialize logging
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] [$Level] $Message"
    
    # Write to log file
    $LogEntry | Out-File -FilePath $LogFile -Append -Encoding UTF8
    
    # Write to console based on level
    switch ($Level) {
        "Error" { Write-Host $Message -ForegroundColor Red }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Success" { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message -ForegroundColor White }
    }
}

# Statistics tracking
$Stats = @{
    UninstallAttempts = 0
    UninstallSuccesses = 0
    RegistryEntriesRemoved = 0
    StartupShortcutsRemoved = 0
    UserDataFoldersRemoved = 0
    UsersProcessed = 0
    Errors = 0
}

Write-Log "=== OneDrive Complete Removal Script Started ===" -Level "Info"
Write-Log "Remove user  $RemoveUserData" -Level "Info"

# ===== STEP 1: UNINSTALL ONEDRIVE SYSTEM-WIDE =====
Write-Log "Step 1: Uninstalling OneDrive system-wide..." -Level "Info"

$oneDriveSetupPaths = @(
    "$env:SystemRoot\System32\OneDriveSetup.exe",
    "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
)

foreach ($setupPath in $oneDriveSetupPaths) {
    $Stats.UninstallAttempts++
    Write-Log "Attempting uninstall with: $setupPath" -Level "Info"
    
    if (Test-Path $setupPath) {
        try {
            Start-Process -FilePath $setupPath -ArgumentList "/uninstall" -Wait -NoNewWindow
            Write-Log "Successfully executed uninstall command: $setupPath" -Level "Success"
            $Stats.UninstallSuccesses++
        } catch {
            Write-Log "Failed to execute uninstall command: $setupPath - $($_.Exception.Message)" -Level "Error"
            $Stats.Errors++
        }
    } else {
        Write-Log "OneDrive setup not found at: $setupPath" -Level "Warning"
    }
}

# ===== STEP 2: REMOVE ONEDRIVE FROM STARTUP FOR ALL EXISTING USERS =====
Write-Log "Step 2: Removing OneDrive from startup for all existing users..." -Level "Info"

# Get all user profiles with NTUSER.DAT files (existing users)
$userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object { 
    Test-Path "$($_.FullName)\NTUSER.DAT" -and 
    $_.Name -notin @('Default', 'Public', 'All Users', 'Default User')
}

Write-Log "Found $($userProfiles.Count) user profiles to process" -Level "Info"

foreach ($profile in $userProfiles) {
    $userName = $profile.Name
    $userPath = $profile.FullName
    $Stats.UsersProcessed++
    Write-Log "Processing user: $userName" -Level "Info"
    
    # Remove from Registry (HKCU Run Key)
    try {
        $ntUserDat = "$userPath\NTUSER.DAT"
        $tempHive = "HKU\TempHive_$userName"
        
        # Load the user's registry hive
        $loadResult = reg load $tempHive $ntUserDat 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Loaded registry hive for $userName" -Level "Info"
            
            # Check if OneDrive entry exists in Run key
            $runKeyPath = "Registry::$tempHive\Software\Microsoft\Windows\CurrentVersion\Run"
            if (Test-Path $runKeyPath) {
                $runKey = Get-ItemProperty -Path $runKeyPath -ErrorAction SilentlyContinue
                if ($runKey -and $runKey.PSObject.Properties.Name -contains "OneDrive") {
                    Remove-ItemProperty -Path $runKeyPath -Name "OneDrive" -ErrorAction Stop
                    Write-Log "Removed OneDrive from Run registry key for $userName" -Level "Success"
                    $Stats.RegistryEntriesRemoved++
                } else {
                    Write-Log "OneDrive not found in Run registry key for $userName" -Level "Info"
                }
            }
            
            # Unload the registry hive
            reg unload $tempHive | Out-Null
            Write-Log "Unloaded registry hive for $userName" -Level "Info"
        } else {
            Write-Log "Failed to load registry hive for $userName" -Level "Error"
            $Stats.Errors++
        }
    } catch {
        Write-Log "Error processing registry for $userName`: $($_.Exception.Message)" -Level "Error"
        $Stats.Errors++
        # Attempt to unload hive if it was loaded
        reg unload $tempHive 2>&1 | Out-Null
    }
    
    # Remove from user's personal startup folder
    $userStartupFolder = "$userPath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    $userOneDriveLnk = Join-Path $userStartupFolder "OneDrive.lnk"
    
    if (Test-Path $userOneDriveLnk) {
        try {
            Remove-Item $userOneDriveLnk -Force
            Write-Log "Removed OneDrive shortcut from startup folder for $userName" -Level "Success"
            $Stats.StartupShortcutsRemoved++
        } catch {
            Write-Log "Failed to remove OneDrive shortcut from startup folder for $userName`: $($_.Exception.Message)" -Level "Error"
            $Stats.Errors++
        }
    } else {
        Write-Log "OneDrive shortcut not found in startup folder for $userName" -Level "Info"
    }
}

# Remove from All Users Startup Folder
$allUsersStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
$allUsersOneDriveLnk = Join-Path $allUsersStartup "OneDrive.lnk"

if (Test-Path $allUsersOneDriveLnk) {
    try {
        Remove-Item $allUsersOneDriveLnk -Force
        Write-Log "Removed OneDrive shortcut from All Users startup folder" -Level "Success"
        $Stats.StartupShortcutsRemoved++
    } catch {
        Write-Log "Failed to remove OneDrive shortcut from All Users startup folder: $($_.Exception.Message)" -Level "Error"
        $Stats.Errors++
    }
} else {
    Write-Log "OneDrive shortcut not found in All Users startup folder" -Level "Info"
}

# ===== STEP 3: PREVENT ONEDRIVE SETUP FOR NEW USERS =====
Write-Log "Step 3: Preventing OneDrive setup for new users..." -Level "Info"

try {
    $defaultUserHive = "hklm\Default_profile"
    $defaultUserDat = "C:\Users\Default\NTUSER.DAT"
    
    if (Test-Path $defaultUserDat) {
        # Load default user hive
        $loadResult = reg load $defaultUserHive $defaultUserDat 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Loaded Default user registry hive" -Level "Info"
            
            # Remove OneDriveSetup from Run key
            $deleteResult = reg delete "$defaultUserHive\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "OneDriveSetup" /f 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Removed OneDriveSetup from Default user Run key" -Level "Success"
                $Stats.RegistryEntriesRemoved++
            } else {
                Write-Log "OneDriveSetup not found in Default user Run key (may already be removed)" -Level "Info"
            }
            
            # Unload default user hive
            reg unload $defaultUserHive | Out-Null
            Write-Log "Unloaded Default user registry hive" -Level "Info"
        } else {
            Write-Log "Failed to load Default user registry hive" -Level "Error"
            $Stats.Errors++
        }
    } else {
        Write-Log "Default user NTUSER.DAT not found" -Level "Warning"
    }
} catch {
    Write-Log "Error preventing OneDrive setup for new users: $($_.Exception.Message)" -Level "Error"
    $Stats.Errors++
}

# ===== STEP 4: REMOVE ONEDRIVE DATA AND CREDENTIALS (OPTIONAL) =====
if ($RemoveUserData) {
    Write-Log "Step 4: Removing OneDrive data and credentials..." -Level "Info"
    
    foreach ($profile in $userProfiles) {
        $userName = $profile.Name
        $userPath = $profile.FullName
        Write-Log "Removing OneDrive data for user: $userName" -Level "Info"
        
        $oneDriveFolders = @(
            "$userPath\OneDrive",
            "$userPath\AppData\Local\Microsoft\OneDrive",
            "$userPath\AppData\Roaming\Microsoft\OneDrive"
        )
        
        foreach ($folder in $oneDriveFolders) {
            if (Test-Path $folder) {
                try {
                    Remove-Item $folder -Recurse -Force
                    Write-Log "Removed OneDrive folder: $folder" -Level "Success"
                    $Stats.UserDataFoldersRemoved++
                } catch {
                    Write-Log "Failed to remove OneDrive folder: $folder - $($_.Exception.Message)" -Level "Error"
                    $Stats.Errors++
                }
            } else {
                Write-Log "OneDrive folder not found: $folder" -Level "Info"
            }
        }
    }
} else {
    Write-Log "Step 4: Skipping OneDrive data removal (RemoveUserData = false)" -Level "Info"
}

# ===== FINAL SUMMARY =====
Write-Log "=== OneDrive Complete Removal Script Completed ===" -Level "Info"
Write-Log "SUMMARY STATISTICS:" -Level "Info"
Write-Log "- Uninstall attempts: $($Stats.UninstallAttempts)" -Level "Info"
Write-Log "- Uninstall successes: $($Stats.UninstallSuccesses)" -Level "Info"
Write-Log "- Registry entries removed: $($Stats.RegistryEntriesRemoved)" -Level "Info"
Write-Log "- Startup shortcuts removed: $($Stats.StartupShortcutsRemoved)" -Level "Info"
Write-Log "- User data folders removed: $($Stats.UserDataFoldersRemoved)" -Level "Info"
Write-Log "- Users processed: $($Stats.UsersProcessed)" -Level "Info"
Write-Log "- Errors encountered: $($Stats.Errors)" -Level "Info"
Write-Log "Log file saved to: $LogFile" -Level "Info"

# Exit with appropriate code
if ($Stats.Errors -eq 0) {
    Write-Log "Script completed successfully with no errors" -Level "Success"
    exit 0
} else {
    Write-Log "Script completed with $($Stats.Errors) errors - check log for details" -Level "Warning"
    exit 1
}
