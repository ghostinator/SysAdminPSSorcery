# OneDrive Removal and Startup Cleanup: README

This repository provides scripts and instructions for **completely removing Microsoft OneDrive** from Windows systems, including uninstalling the application, cleaning up startup entries for all users, and preventing OneDrive from launching or reinstalling.

## Table of Contents

- [Overview](#overview)
- [Complete OneDrive Removal Script with Logging](#complete-onedrive-removal-script-with-logging)
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


# Complete OneDrive Removal Script with Logging

## Overview

This PowerShell script provides **comprehensive OneDrive removal** capabilities for Windows systems with **enhanced logging and error tracking**. The script is modular, allowing administrators to select specific removal actions via command-line switches, making it ideal for both manual execution and enterprise deployment via Microsoft Intune.

## Key Features

### **Modular Design**
- **Selectable actions** via command-line switches - run only what you need
- **Comprehensive coverage** of all OneDrive components and integration points
- **Enterprise-ready** with proper exit codes and detailed logging

### **Enhanced Logging & Error Tracking**
- **Complete error capture** - all errors, warnings, and operations logged to file
- **Detailed error information** including exception details, stack traces, and affected files
- **File lock detection** - identifies which specific files cannot be deleted and why
- **Statistics tracking** - comprehensive reporting of all operations performed
- **Debug-level logging** for troubleshooting complex issues

### **Comprehensive Removal Options**
- **Process termination** - Stops all running OneDrive processes
- **Application uninstall** - Removes both legacy and modern OneDrive applications
- **Startup cleanup** - Removes OneDrive from all user startup locations
- **Data removal** - Deletes OneDrive folders and cached credentials
- **Registry cleanup** - Removes OneDrive integration from Windows Explorer
- **Policy enforcement** - Prevents OneDrive reinstallation and usage
- **New user prevention** - Blocks OneDrive setup for future user accounts

## Available Actions

| Switch | Action | Description |
|--------|--------|-------------|
| `-StopProcesses` | Stop OneDrive Processes | Terminates all running OneDrive processes |
| `-UninstallOneDrive` | Uninstall Application | Removes OneDrive using system uninstallers |
| `-RemoveStartupEntries` | Startup Cleanup | Removes OneDrive from all user startup locations |
| `-RemoveModernApp` | Modern App Removal | Uninstalls OneDrive Appx package if present |
| `-RemoveUserData` | Data Removal | Deletes OneDrive folders and cached data |
| `-RegistryCleanup` | Registry Cleanup | Removes OneDrive from Windows Explorer integration |
| `-PreventSetupNewUsers` | Block New Users | Prevents OneDrive setup for new user accounts |
| `-PolicyBlock` | Policy Enforcement | Sets registry policy to block OneDrive usage |
| `-RestartExplorer` | Explorer Restart | Restarts Windows Explorer to apply changes |

## Usage Examples

### **Complete Removal (All Actions)**
```powershell
.\OneDriveCompleteRemoval.ps1 -StopProcesses -UninstallOneDrive -RemoveStartupEntries -RemoveModernApp -RemoveUserData -RegistryCleanup -PreventSetupNewUsers -PolicyBlock -RestartExplorer
```

### **Basic Cleanup (No Data Removal)**
```powershell
.\OneDriveCompleteRemoval.ps1 -StopProcesses -UninstallOneDrive -RemoveStartupEntries -PolicyBlock
```

### **Startup Cleanup Only**
```powershell
.\OneDriveCompleteRemoval.ps1 -RemoveStartupEntries
```

### **Policy Block Only**
```powershell
.\OneDriveCompleteRemoval.ps1 -PolicyBlock
```

## Remote Execution Considerations

### **User Interruption**
When running this script remotely (via Intune, RMM tools, or remote PowerShell), **most actions will not interrupt the user's current session**:

- **Process termination** may briefly affect users actively using OneDrive
- **Registry changes** apply immediately but don't disrupt current applications
- **File/folder removal** operates in the background
- **Explorer restart** (`-RestartExplorer`) **will briefly interrupt the user** by restarting Windows Explorer

### **Recommendations for Remote Deployment**
- **Exclude `-RestartExplorer`** for non-disruptive remote execution
- **Schedule during maintenance windows** if using `-RestartExplorer`
- **Use `-StopProcesses`** to minimize file lock issues
- **Monitor via logs** rather than console output for remote execution

### **Non-Disruptive Remote Command**
```powershell
.\OneDriveCompleteRemoval.ps1 -StopProcesses -UninstallOneDrive -RemoveStartupEntries -RemoveModernApp -RemoveUserData -RegistryCleanup -PreventSetupNewUsers -PolicyBlock
```
*Note: Excludes `-RestartExplorer` to avoid user interruption*

## System Requirements

- **Windows 10/11** (tested on current versions)
- **Administrator privileges** (required for all operations)
- **PowerShell 5.1 or later**

## Deployment Methods

### **Microsoft Intune**
1. Upload script to **Devices > Scripts and remediations > Platform scripts**
2. Configure to **run as Administrator** (device context)
3. Set **Run script in 64-bit PowerShell host**: Yes
4. Assign to target device groups
5. Monitor deployment status via exit codes

### **Manual Execution**
1. Run PowerShell as Administrator
2. Execute with desired switches
3. Review log file for detailed results

### **Remote Management Tools**
- Compatible with most RMM platforms
- Use device/system context for full functionality
- Monitor via log files rather than console output

## Logging and Troubleshooting

### **Log File Location**
- **Path**: `%TEMP%\OneDriveCombinedRemoval_YYYY-MM-DD_HH-mm-ss.log`
- **Content**: Complete operation log with timestamps, error details, and statistics
- **Format**: Human-readable text with structured entries

### **Log Information Includes**
- **Detailed error messages** with exception information
- **File lock detection** and affected file lists
- **Registry operation results** with specific keys/values
- **Process termination details** with PIDs
- **Folder sizes** before deletion
- **Complete statistics** of all operations performed

### **Exit Codes**
- **0**: Success (no errors encountered)
- **1**: Failure (errors occurred - check log for details)

## Important Notes

### **Data Safety**
- **`-RemoveUserData` permanently deletes OneDrive folders** - use with caution
- **Always test in non-production environment** before wide deployment
- **Review log files** for any unexpected errors or locked files

### **Enterprise Considerations**
- **Combines well with Group Policy** for additional OneDrive restrictions
- **Compatible with Windows Update policies** and other system management
- **Supports both domain-joined and Azure AD-joined devices**

### **Limitations**
- **Some files may remain locked** if OneDrive processes cannot be terminated
- **Registry changes may require reboot** for complete effect
- **New OneDrive installations** may override some settings if not blocked by policy

## Support and Troubleshooting

### **Common Issues**
- **Access denied errors**: Ensure script runs as Administrator
- **File lock errors**: Stop all OneDrive processes before data removal
- **Registry errors**: Check for existing Group Policy conflicts

### **Best Practices**
- **Run complete removal** during maintenance windows
- **Test on pilot group** before organization-wide deployment
- **Keep log files** for compliance and troubleshooting
- **Combine with policy enforcement** for permanent OneDrive blocking


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
