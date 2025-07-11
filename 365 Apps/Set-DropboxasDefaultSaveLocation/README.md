# Intune PowerShell Script: Set Default Save Location for Office and Documents

## Overview

This script configures the **default save location** for the Windows Documents folder and Microsoft Office applications (Word, Excel, PowerPoint) to the user's Dropbox folder.

It also sets a key policy to make Office applications **"Save to Computer by default"** instead of prioritizing cloud locations like OneDrive. This ensures the local path setting is honored. It is intended for deployment via **Microsoft Intune**.

> **Note:**
> This script is designed to run in the user's context and only modifies specific registry keys. No other Office or Windows settings are changed.

---

## How It Works

The script performs the following actions:

-   **Sets Office to Prefer Local Saves**: Modifies the registry to enable the "Save to Computer by default" setting across the Office suite. This is a prerequisite for ensuring the other changes take effect as expected.
-   **Redirects Windows Documents folder**: Reliably sets the default path to `%USERPROFILE%\Dropbox` by updating both the modern and legacy registry keys for maximum compatibility.
-   **Configures Office Applications**:
    -   **Word**: Sets the "Default local file location" to the Dropbox path.
    -   **Excel**: Sets the "Default local file location" to the Dropbox path.
    -   **PowerPoint**: Sets the default save path to the Dropbox path.

The script is idempotent, meaning it checks current values before making any changes and only acts if necessary. All actions are logged to a file in `C:\temp\`.

---

## Deployment Instructions

1.  **Upload the script to Intune**
    -   Go to **Intune admin center** > **Devices** > **Scripts and remediations** > **Platform scripts** > **Add** > **Windows 10 and later**.
    -   Name and describe the script policy.
    -   Upload the script file.

2.  **Configure script settings**
    -   **Run this script using the logged on credentials**: **Yes** (This is required for per-user settings).
    -   **Enforce script signature check**: No (Unless required by your organization).
    -   **Run script in 64-bit PowerShell host**: **Yes**.

3.  **Assign the script** to the appropriate user or device groups.

4.  **Monitor deployment** in the Intune portal for success or error reporting.

---

## Important Notes

-   The script **must** run in the **user context** to access the correct registry hive (`HKCU`).
-   Users may need to restart Office applications for the save path changes to take effect. For the Documents folder redirection, a **sign-out and sign-in** may be required for the change to be reflected everywhere in Windows.
-   The script creates the Dropbox folder (`%USERPROFILE%\Dropbox`) if it is missing. This assumes a standard Dropbox installation path.

---

## Privacy and Security

-   The script does **not** collect or transmit any user data.
-   No sensitive information is stored or processed.
-   Always review and test scripts in a non-production environment before wide deployment.

---

## Support

For questions or issues, contact your IT administrator or refer to the official [Intune PowerShell documentation](https://learn.microsoft.com/mem/intune/apps/intune-management-extension).