# Dropbox Office Integration Script

## Overview
This PowerShell script automates the configuration of Microsoft Office to use Dropbox as the default cloud storage provider. It performs the following tasks:
- Uninstalls OneDrive and removes leftover files
- Re-enables Office cloud storage features that may have been disabled by Group Policy
- Configures Dropbox as a cloud storage provider for Office applications
- Applies these settings to all existing user profiles and the Default User profile
- Logs all actions and can send email notifications upon completion

## Prerequisites
- Windows 10/11
- PowerShell 5.1 or higher
- Administrative privileges
- Microsoft Office 365 installed
- Dropbox installed or planned installation

## Deployment Options
### Manual Execution
1. Download the script to the target computer
2. Open PowerShell as Administrator
3. Set execution policy: `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process`
4. Navigate to the script location: `cd C:\path\to\script`
5. Run the script: `.\DropboxOfficeIntegration.ps1`

### Microsoft Intune Deployment
1. In the Intune admin center, go to **Devices** > **Windows** > **PowerShell scripts**
2. Click **+ Add** to create a new script
3. Upload the script file and configure the following settings:
   - Run script in 64-bit PowerShell: **Yes**
   - Run script with administrative privileges: **Yes**
   - Enforce script signature check: **No**
4. Assign the script to the appropriate device groups
5. The script will run automatically on targeted devices

## Script Structure
The script consists of several key functions:

- `ReEnable-OfficeCloudFeatures`: Removes Group Policy restrictions on Office cloud storage
- `Uninstall-OneDrive`: Uninstalls the OneDrive application
- `Cleanup-OneDriveLeftovers`: Removes OneDrive remnants from the system
- `Add-DropboxCloudStorageKey`: Adds Dropbox registry entries for Office integration
- `Configure-For-AllUsers`: Applies settings to all existing user profiles
- `Configure-DefaultUser`: Applies settings to the Default User profile
- `Write-Log`: Handles logging to file and console
- `Send-CompletionEmail`: Sends email notification with execution results

## Configuration
### Email Notifications
To enable email notifications, update the following variables in the script:
```powershell
$emailConfig = @{
    To         = "recepient@domain.com
    From       = "sender@yourdomain.com"
    Subject    = "Dropbox Office Integration - $env:COMPUTERNAME"
    SMTPServer = "smtp.yourserver.yeah"
    Port       = 587
    UseSSL     = $true
}
```
### Logging
By default, the script logs to `C:\temp\dropboxupdates.log`. You can modify the log path by changing the `$logFile` variable.

## Troubleshooting
- **Registry Access Errors**: Ensure the script is running with administrative privileges
- **Email Sending Failures**: Verify SMTP server settings and credentials
- **OneDrive Uninstall Issues**: Check if OneDrive is in use or locked by another process
- **Registry Modification Errors**: Ensure the user profiles are not corrupted

## Security Considerations
- The script contains email credentials in plain text. Consider using a secure credential store in production.
- Registry modifications are performed system-wide. Test thoroughly before deploying to production.

## License
This script is provided as-is with no warranty. Use at your own risk.

## Author
Created by: Brandon Cook
Last Updated: July 22, 2025
