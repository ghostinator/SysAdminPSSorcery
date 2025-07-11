# OneDrive Removal and Startup Cleanup: README

This repository provides scripts and instructions for **completely removing Microsoft OneDrive** from Windows systems, including uninstalling the application, cleaning up startup entries for all users, and preventing OneDrive from launching or reinstalling.

## Table of Contents

- [Overview](#overview)
- [OneDrive Complete Removal Script (All Actions Enabled by Default)](#onedrive-complete-removal-script-all-actions-enabled-by-default)
- [Removal Options](#removal-options)
  - [1. Uninstall OneDrive System-Wide](#1-uninstall-onedrive-system-wide)
  - [2. Remove OneDrive from Startup for All Users](#2-remove-onedrive-from-startup-for-all-users)
  - [3. Prevent OneDrive Setup for New Users](#3-prevent-onedrive-setup-for-new-users)
  - [4. Remove OneDrive Data and Credentials (Optional)](#4-remove-onedrive-data-and-credentials-optional)
- [Sample Scripts](#sample-scripts)
- [Best Practices](#best-practices)
- [References](#references)

## Overview

Microsoft OneDrive is integrated into Windows 10 and 11, but many organizations and users prefer to remove it for privacy, compliance, or workflow reasons. This guide covers all major removal and cleanup options, with sample PowerShell scripts for automation.


# OneDrive Complete Removal Script (All Actions Enabled by Default)

## Overview

This PowerShell script provides a **comprehensive, modular, and enterprise-ready solution** for removing Microsoft OneDrive from Windows systems. All removal and cleanup actions are enabled by default, and every step is logged in detail for auditing and troubleshooting. The script is suitable for manual use, remote execution, and deployment via Microsoft Intune or other management tools.

## Features

- **All actions enabled by default**—no arguments required for full removal.
- **Modular switches**—override any action by passing the switch as `$false`.
- **Enhanced logging**—all operations, errors, and warnings are logged to a timestamped file in the system temp directory.
- **Enterprise-ready**—handles all user profiles, system-wide registry, and policy enforcement.
- **Safe and robust**—detailed error handling, file lock detection, and statistics reporting.

## Actions Performed

| Action                       | Description                                                                                 | Switch (default: enabled)      |
|------------------------------|--------------------------------------------------------------------------------------------|-------------------------------|
| Stop OneDrive Processes      | Terminates all running OneDrive processes                                                  | `-StopProcesses`              |
| Uninstall OneDrive           | Removes OneDrive using system uninstallers (classic/legacy)                                | `-UninstallOneDrive`          |
| Remove Modern App            | Uninstalls the OneDrive Appx package (if present)                                          | `-RemoveModernApp`            |
| Remove Startup Entries       | Removes OneDrive from all user and system startup locations                                | `-RemoveStartupEntries`       |
| Remove User Data             | Deletes OneDrive folders and cached credentials for all users and system locations         | `-RemoveUserData`             |
| Registry Cleanup             | Removes OneDrive integration from Windows Explorer and related registry keys               | `-RegistryCleanup`            |
| Prevent Setup for New Users  | Blocks OneDrive setup for new user accounts by cleaning the Default user registry hive     | `-PreventSetupNewUsers`       |
| Policy Block                 | Sets registry policy to block OneDrive reinstallation and usage                            | `-PolicyBlock`                |
| Restart Explorer             | Restarts Windows Explorer to apply changes (can be skipped for non-disruptive remote use)  | `-RestartExplorer`            |

## Usage

### **Run All Actions (Default)**
```powershell
.\OneDriveCompleteRemoval.ps1
```

### **Skip Any Action**
To skip an action, pass the switch as `$false`:
```powershell
.\OneDriveCompleteRemoval.ps1 -RestartExplorer:$false
```
You can combine multiple overrides as needed.

## Logging & Reporting

- **Log file location:**  
  `%TEMP%\OneDriveCombinedRemoval_YYYY-MM-DD_HH-mm-ss.log`
- **Log contents:**  
  - All actions, errors, warnings, and debug information
  - Exception details, stack traces, and locked file lists
  - Summary statistics for all operations

- **Exit codes:**  
  - `0` = Success (no errors)
  - `1` = Errors encountered (see log for details)

## Remote Execution & User Experience

- **Non-disruptive by default** (except for `-RestartExplorer`, which restarts Windows Explorer and may briefly interrupt the user).
- **Recommended for remote/Intune use:**  
  Omit `-RestartExplorer` for a seamless user experience.
- **All other actions** run in the background and do not require user interaction.

## System Requirements

- **Windows 10/11**
- **PowerShell 5.1 or later**
- **Administrator privileges**

## Best Practices

- **Test in a non-production environment** before wide deployment.
- **Review the log file** after execution for any errors or locked files.
- **Schedule during maintenance windows** if using `-RestartExplorer`.
- **Combine with Group Policy** for additional OneDrive restrictions if needed.

## Example: Intune Deployment

1. Upload the script to Intune as a device-context PowerShell script.
2. Assign to target device groups.
3. Monitor deployment status via exit codes and log files.

## Important Notes

- **`-RemoveUserData` permanently deletes OneDrive folders and cached data.**
- **Some files may remain if locked by other processes.** The script logs all such files for manual follow-up.
- **A system restart is recommended** after running the script to finalize all changes.

## Support

For troubleshooting, review the generated log file for detailed error and operation information.  
If you encounter persistent issues, ensure the script is run as Administrator and that all OneDrive processes are stopped before data removal.



## Manual Removal Options

### 1. Uninstall OneDrive System-Wide

- **Windows 10/11:** OneDrive can be uninstalled via Settings or with a command-line script.
- **Command:**
  - For 64-bit Windows:  
    `C:\Windows\System32\OneDriveSetup.exe /uninstall`
  - For 32-bit Windows:  
    `C:\Windows\SysWOW64\OneDriveSetup.exe /uninstall`
- **Note:** This must be run as an administrator.

### 2. Remove OneDrive from Startup for All Users

- **Why:** Even after uninstalling, startup entries may remain in user profiles, causing errors or reinstallation.
- **What to Remove:**
  - Registry: `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` → `OneDrive`
  - Startup Folders:  
    - `C:\Users\<username>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\OneDrive.lnk`
    - `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\OneDrive.lnk`
- **How:** Use a PowerShell script to iterate all user profiles and remove these entries.

### 3. Prevent OneDrive Setup for New Users

- **Why:** By default, new user profiles may auto-install OneDrive.
- **How:** Remove the `OneDriveSetup` entry from the Default user registry hive:
  - Load `C:\Users\Default\NTUSER.DAT` and delete the `OneDriveSetup` value from `SOFTWARE\Microsoft\Windows\CurrentVersion\Run`.

### 4. Remove OneDrive Data and Credentials (Optional)

- **Why:** For privacy or to ensure no residual data remains.
- **How:** Delete OneDrive folders and cached credentials from each user profile.

## Sample Scripts

### A. Uninstall OneDrive (System-Wide)

```powershell
# Uninstall OneDrive for all users (run as Administrator)
$paths = @(
    "$env:SystemRoot\System32\OneDriveSetup.exe",
    "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
)
foreach ($exe in $paths) {
    if (Test-Path $exe) {
        Start-Process -FilePath $exe -ArgumentList "/uninstall" -Wait
    }
}
```

### B. Remove OneDrive from Startup for All Existing Users

```powershell
#Requires -RunAsAdministrator
# Removes OneDrive from Run registry and Startup folders for all user profiles

$userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
    Test-Path "$($_.FullName)\NTUSER.DAT" -and
    $_.Name -notin @('Default', 'Public', 'All Users', 'Default User')
}

foreach ($profile in $userProfiles) {
    $userName = $profile.Name
    $userPath = $profile.FullName
    $ntUserDat = "$userPath\NTUSER.DAT"
    $tempHive = "HKU\TempHive_$userName"
    reg load $tempHive $ntUserDat | Out-Null

    $runKeyPath = "Registry::$tempHive\Software\Microsoft\Windows\CurrentVersion\Run"
    if (Test-Path $runKeyPath) {
        Remove-ItemProperty -Path $runKeyPath -Name "OneDrive" -ErrorAction SilentlyContinue
    }
    reg unload $tempHive | Out-Null

    $startupFolder = "$userPath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    $oneDriveLnk = Join-Path $startupFolder "OneDrive.lnk"
    if (Test-Path $oneDriveLnk) {
        Remove-Item $oneDriveLnk -Force
    }
}

# Remove from All Users Startup folder
$allUsersStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\OneDrive.lnk"
if (Test-Path $allUsersStartup) {
    Remove-Item $allUsersStartup -Force
}
```

### C. Prevent OneDrive Setup for New Users

```powershell
# Prevent OneDrive from auto-installing for new users (run as Administrator)
reg load "hklm\Default_profile" "C:\Users\Default\NTUSER.DAT"
reg delete "hklm\Default_profile\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "OneDriveSetup" /f
reg unload "hklm\Default_profile"
```

### D. Remove OneDrive Data and Credentials (Optional)

```powershell
# Remove OneDrive folders and cached credentials for all users
$userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
    $_.Name -notin @('Default', 'Public', 'All Users', 'Default User')
}
foreach ($profile in $userProfiles) {
    $userPath = $profile.FullName
    Remove-Item "$userPath\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$userPath\AppData\Local\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$userPath\AppData\Roaming\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
}
```

## Best Practices

- **Run all scripts as Administrator** to ensure access to all user profiles and system locations.
- **Test on a non-production machine** before wide deployment.
- **Combine these scripts** as needed for your environment and requirements.
- **Document changes** for compliance and troubleshooting.


**Disclaimer:**  
These scripts are provided as-is. Use at your own risk and always test in a safe environment before deploying to production systems.
