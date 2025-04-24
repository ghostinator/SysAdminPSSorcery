# M365 Primary Calendar Permissions Manager GUI

**Author:** Brandon Cook 
**Copyright:** (c) 2025 Brandon Cook
**GitHub:** https://github.com/ghostinator/SysAdminPSSorcery
**Script Name:** `Manage-PrimaryCalendarPermissionsGui.ps1` 
**Version:** 1.9

## Description

This PowerShell script provides a Windows Presentation Foundation (WPF) graphical user interface (GUI) for administrators to manage delegate permissions specifically for users' **primary** calendar folders within a Microsoft 365 / Exchange Online environment.

It simplifies viewing, adding, modifying, and removing permissions (like Viewer, Editor, Author) that a delegate user has on another user's main calendar.

![alt text](<CleanShot 2025-04-24 at 14.35.54.png>)

## Features

* **Connect to Exchange Online:** Securely connects using the modern ExchangeOnlineManagement module (v3+). Handles modern authentication, automatically attempting browser-based login if integrated Windows authentication (WAM) fails or causes issues.
* **Module Auto-Handling:** Checks if the required `ExchangeOnlineManagement` PowerShell module is installed and attempts to install it for the current user if missing.
* **Mailbox Browser:** Includes a pop-up window to search for or load all mailboxes (with performance warnings) to easily select the target user.
* **Alias Resolution:** If an email address typed into the "Target Mailbox" field isn't a primary address, the script attempts to resolve it as an alias.
* **Calendar Listing:** Lists all calendar-type folders found in the target user's mailbox.
* **Primary Calendar Permissions:** Allows viewing, adding, setting, and removing delegate permissions *only* on the folder identified as the primary `\Calendar`.
* **Clear Logging:** Actions, successes, and errors are logged to the "Results" text box within the GUI.

## Prerequisites

1.  **Windows PowerShell:** Version 5.1 or later.
2.  **Windows Environment:** Designed to run on Windows with .NET Framework (for WPF).
3.  **Run as Administrator:** The script requires elevated privileges to connect to Exchange Online and manage permissions across mailboxes. You must launch PowerShell "As Administrator".
4.  **Microsoft 365 Permissions:** The account running the script needs sufficient administrative permissions in Microsoft 365 to manage mailbox folder permissions (e.g., Exchange Administrator, Global Administrator, or a suitable custom role).
5.  **ExchangeOnlineManagement Module:** The script requires this PowerShell module. It will attempt to automatically install it from the PowerShell Gallery if it's not detected (requires internet connectivity and appropriate execution policy for installation).
6.  **PowerShell Execution Policy:** Your system's PowerShell execution policy must allow running local scripts. See Setup below.

## Limitations

* **Primary Calendar Only:** Due to observed limitations and inconsistencies with Exchange Online cmdlets when identifying non-primary calendar folders reliably via script, this tool **only supports managing permissions on the primary `\Calendar` folder**. Other calendar folders are listed for informational purposes but are not actionable through this tool. Attempts to select a non-primary calendar for permission changes will result in an informative message.

## Setup and Usage

### 1. PowerShell Execution Policy

PowerShell's execution policy can prevent scripts from running. You need to ensure it's set to allow local scripts for your user account.

* **Check Current Policy:** Open PowerShell (regular, not necessarily as admin for this check) and run:
    ```powershell
    Get-ExecutionPolicy -Scope CurrentUser
    ```
* **Set Policy (If Needed):** If the policy is `Restricted` or `AllSigned`, you'll need to change it to `RemoteSigned` or `Unrestricted` for the current user. `RemoteSigned` is generally recommended as it allows local scripts but requires downloaded scripts to be signed. Run PowerShell **as Administrator** to execute the following command:
    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    ```
    * Answer 'Y' or 'A' if prompted for confirmation.
    * **Security Note:** Understand the implications of changing the execution policy. `RemoteSigned` provides a balance for running your own scripts while adding some protection against untrusted downloaded scripts.

### 2. Run as Administrator

* Right-click the PowerShell icon or shortcut.
* Select "Run as administrator".
* Accept the User Account Control (UAC) prompt if it appears.

### 3. Execute the Script

* Navigate to the directory where you saved the script file using the `cd` command. For example:
    ```powershell
    cd "C:\Scripts"
    ```
* Execute the script:
    ```powershell
    .\Manage-PrimaryCalendarPermissionsGui.ps1
    ```
    *(Replace the filename if you saved it differently).*
* The script will first check for the required module and attempt installation if needed (you might see progress in the console).

### 4. Using the GUI

1.  **Connect:** Click the "Connect/Disconnect" button. Authenticate using the browser window that appears (due to the `-DisableWAM` flag ensuring browser auth). The status should change to "Connected".
2.  **Target Mailbox:**
    * Type the email address or alias of the user *whose calendar you want to manage* into the "Target Mailbox (Owner)" box.
    * Alternatively, click the "..." (Browse) button to open the Mailbox Browser pop-up. Search (min 3 chars) or use "Load All" (with caution), select a mailbox, and click "Select".
3.  **List Calendars:** Click the "List Calendars" button. The script will resolve the target mailbox (checking aliases if needed) and list the found calendar folders below. The primary `\Calendar` folder should be listed first and selected by default.
4.  **Select Primary Calendar:** Ensure the `\Calendar` entry is selected in the list view. (Actions on other calendars are disabled).
5.  **Delegate User:** Enter the email address of the user *whose permissions you want to view or change* in the "Delegate User" box.
6.  **Permission Level:** Select the desired access right (e.g., Editor, Reviewer, Author) from the dropdown for Add/Set actions. This is ignored for View/Remove actions.
7.  **Action Buttons:**
    * **View Permissions:** Shows current permissions for the selected delegate on the primary calendar.
    * **Add / Set Permission:** Adds the selected permission for the delegate user or modifies their existing permission to the selected level.
    * **Remove Permission:** Removes all permissions for the specified delegate user from the primary calendar (requires confirmation).
8.  **Results:** Monitor the "Results" box for status messages and errors.
9.  **Disconnect/Close:** Click "Disconnect" or "Close" when finished. Closing the window also attempts to disconnect the session.

## Disclaimer

This script is provided as-is. Always test thoroughly in a non-production environment before using it on production systems. The author is not responsible for any unintended consequences or data loss resulting from the use of this script. Ensure you have appropriate backups and understand the permissions you are granting or revoking.