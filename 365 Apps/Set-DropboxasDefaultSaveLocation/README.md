# Intune PowerShell Script: Set Default Save Location for Office and Documents

## Overview

This script configures the **default save location** for the Windows Documents folder and Microsoft Office applications (Word, Excel, PowerPoint) to the user's Dropbox folder. It is intended for deployment via **Microsoft Intune** using the PowerShell scripts feature.

> **Note:**  
> This script only modifies registry keys related to the default save location. No other Office or Windows registry settings are changed.

---

## How It Works

- **Windows Documents folder**: Sets the default path to `%USERPROFILE%\Dropbox`.
- **Word**: Sets the default save path to Dropbox.
- **Excel**: Sets the default save path to Dropbox.
- **PowerPoint**: Sets the default save path to Dropbox.

The script checks current values before making changes and logs all actions to a file in `C:\temp\`.

---

## Deployment Instructions

1. **Upload the script to Intune**  
   - Go to **Intune admin center** > **Devices** > **Scripts and remediations** > **Platform scripts** > **Add** > **Windows 10 and later**.
   - Name and describe the script policy.
   - Upload the script file.

2. **Configure script settings**  
   - **Run this script using the logged on credentials**: Yes (recommended for per-user settings).
   - **Enforce script signature check**: As required by your organization. Most likely no
   - **Run script in 64-bit PowerShell host**: Yes (recommended).

3. **Assign the script** to the appropriate user or device groups.

4. **Monitor deployment** in the Intune portal for success or error reporting.

---

## Important Notes

- The script should run in **user context** to access the correct registry hive (`HKCU`).
- The Dropbox folder must exist in the user's profile. The script creates it if missing.
- Users may need to restart Office applications for changes to take effect.
- The script is **idempotent**: it only updates settings if needed.

---

## Privacy and Security

- The script does **not** collect or transmit any user data.
- No sensitive information is stored or processed.
- Always review and test scripts in a non-production environment before wide deployment.

---

## Support

For questions or issues, contact your IT administrator or refer to the official [Intune PowerShell documentation](https://learn.microsoft.com/en-us/intune/intune-service/apps/powershell-scripts)[2].

