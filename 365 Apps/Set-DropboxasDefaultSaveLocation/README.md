
# OneDrive Removal & Dropbox Integration Script

This PowerShell script removes OneDrive from Windows systems and configures Dropbox as the default cloud storage provider for Microsoft Office applications.

## 📋 Overview

The script performs the following operations:

1. Re-enables Office cloud storage features that may have been disabled by policies
2. Uninstalls OneDrive (both 32-bit and 64-bit versions)
3. Cleans up OneDrive leftover files and registry entries
4. Configures Dropbox as a cloud storage provider in Office for all users
5. Logs all operations to a detailed log file

## 🎯 Use Cases

- **Enterprise environments** transitioning from OneDrive to Dropbox
- **IT administrators** needing to standardize on Dropbox for cloud storage
- **Organizations** with specific compliance requirements for cloud storage providers
- **Deployment scenarios** where consistent cloud storage configuration is required across multiple machines


## ⚡ Features

- **Comprehensive OneDrive removal**: Handles both installation cleanup and leftover files
- **Multi-user support**: Configures Dropbox for all existing users and future users
- **Registry hive management**: Safely loads/unloads user registry hives for offline configuration
- **Policy reset**: Removes Group Policy restrictions on cloud storage providers
- **Detailed logging**: Creates timestamped logs of all operations
- **Error handling**: Robust error handling with detailed error logging


## 🔧 Requirements

- **PowerShell 5.1** or later
- **Administrator privileges** (required for system-wide changes)
- **Windows 10/11** or Windows Server 2016+
- **Microsoft Office** (2016 or later) installed on target systems


## 📁 File Structure

```
C:\temp\
└── dropboxupdates.log    # Detailed operation log (created automatically)
```


## 🚀 Usage

### Basic Execution

```powershell
# Run as Administrator
.\Set-DropboxDefaultSave.ps1
```


### Via Intune/MDM

This script is designed to work with Microsoft Intune or other MDM solutions:

1. Upload the script to your MDM platform
2. Deploy as a PowerShell script policy
3. Set to run with administrator privileges

### Manual Deployment

```powershell
# Copy script to target machine
# Run from elevated PowerShell session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
.\Set-DropboxDefaultSave.ps1
```


## 📊 What Gets Modified

### Registry Changes

- **Removes**: `HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\CloudStorage`
- **Removes**: `HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\CloudStorage`
- **Sets**: `UseOnlineContent = 2` in Office Internet policies
- **Adds**: Dropbox cloud storage provider entries for all users


### Files \& Folders Removed

- `%LOCALAPPDATA%\Microsoft\OneDrive`
- `%USERPROFILE%\OneDrive`
- `%PROGRAMDATA%\Microsoft OneDrive`
- `%PROGRAMFILES%\Microsoft OneDrive`
- `%PROGRAMFILES(x86)%\Microsoft OneDrive`


### System Changes

- Uninstalls OneDrive via setup executables
- Removes OneDrive scheduled tasks
- Clears OneDrive Group Policy settings


## 📝 Log File Details

The script creates a comprehensive log at `C:\temp\dropboxupdates.log` containing:

- Timestamp for each operation
- System information (computer name, user, domain)
- Detailed registry modifications
- File/folder operations
- Error messages and stack traces
- Success/failure status for each operation


### Sample Log Output

```
[2025-01-15 10:30:15] === SCRIPT STARTED: Set-DropboxDefaultSave.ps1 ===
[2025-01-15 10:30:15] System Info: Computer: WORKSTATION01, User: admin, Domain: COMPANY
[2025-01-15 10:30:16] ==> Re-enabling Office Cloud Storage integration…
[2025-01-15 10:30:16]     Set UseOnlineContent=2 at HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Internet
[2025-01-15 10:30:17] ==> Uninstalling OneDrive if present…
[2025-01-15 10:30:18]     Running: C:\Windows\System32\OneDriveSetup.exe /uninstall
```


## ⚠️ Important Notes

### Before Running

- **Backup user data**: Ensure OneDrive files are backed up or synced elsewhere
- **Test in lab environment**: Validate the script works in your specific environment
- **User communication**: Notify users about the change to avoid confusion


### Limitations

- **Active OneDrive sync**: The script doesn't handle files currently syncing
- **User notification**: Users aren't automatically notified of the change
- **Dropbox installation**: The script assumes Dropbox is already installed
- **Office versions**: Primarily tested with Office 2016-2021 and Microsoft 365


### Security Considerations

- Requires administrator privileges
- Modifies system-wide registry settings
- Loads/unloads user registry hives
- Should be tested before production deployment


## 🔍 Troubleshooting

### Common Issues

1. **"Access Denied" errors**: Ensure script runs as Administrator
2. **Registry hive load failures**: Check if user profiles are corrupted
3. **OneDrive still appears**: Clear Office cache and restart Office applications
4. **Log file missing**: Verify C:\temp directory permissions

### Verification Steps

```powershell
# Check if OneDrive is uninstalled
Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*OneDrive*"}

# Verify Dropbox registry entries
Get-ChildItem "Registry::HKEY_USERS\*\Software\Microsoft\Office\Common\Cloud Storage" -ErrorAction SilentlyContinue
```


## 📞 Support
It's a script on github,

**⚠️ Disclaimer**: This script makes significant system changes. Always test thoroughly in a non-production environment before deploying to production systems.

