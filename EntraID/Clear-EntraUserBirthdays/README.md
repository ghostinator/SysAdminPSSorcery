# Clear-EntraUserBirthdays.ps1

**Author:** Brandon Cook (brandon@ghostinator.co)
**Date:** May 7, 2025
**Version:** 1.1
## Purpose

This PowerShell script is designed to clear the 'Birthday' attribute for all users in a Microsoft Entra ID tenant. It iterates through users, checks if their birthday is set, and if so, updates the `Birthday` attribute to `null`.

**USE WITH EXTREME CAUTION. THIS IS A DESTRUCTIVE OPERATION.**

## Prerequisites

1.  **PowerShell:** Version 5.1 or later.
2.  **Microsoft Graph PowerShell SDK:** The `Microsoft.Graph.Users` module must be installed.
    ```powershell
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
    ```
3.  **Permissions:** The account used to run this script requires the following Microsoft Graph API permissions. You will be prompted to authenticate when the script connects.
    * `User.ReadWrite.All`: To read user information and update user objects (clear the birthday).
    * `Sites.ReadWrite.All`: Required because the 'Birthday' attribute in the target environment was found to be linked with SharePoint User Profiles, and updating it via Graph API may require permissions to write to underlying site data.
    Admin consent may be required for these permissions in your tenant.

## How to Use

1.  **Save the Script:** Save the script content as `Clear-EntraUserBirthdays.ps1` in a directory on your computer.
2.  **Backup Birthday Data (CRITICAL):** Before running this script, ensure you have a backup of any existing birthday information you might want to restore later. Use the previously generated script `Get-UserBirthdayAndExtAttr2.ps1` (or a similar one) to export current birthday data.
3.  **Open PowerShell:** Launch a PowerShell console.
4.  **Navigate to Script Directory:** Use `cd` to navigate to the directory where you saved the script.
    ```powershell
    cd C:\Path\To\Your\Scripts
    ```
5.  **Run the Script:**
    ```powershell
    .\Clear-EntraUserBirthdays.ps1
    ```
6.  **Confirmation:** The script will display a prominent warning and ask for explicit confirmation (you must type 'YES') before it proceeds with any changes.
7.  **Testing (Highly Recommended):**
    Before running on all users, modify the script to target a specific test user or a small group of test users. You can change the line:
    `$allUsers = Get-MgUser -All -Property "Id,UserPrincipalName" -ErrorAction Stop`
    to something like:
    `$testUserUPNs = "testuser1@yourdomain.com", "testuser2@yourdomain.com"`
    `$allUsers = $testUserUPNs | ForEach-Object { Get-MgUser -UserId $_ -Property "Id,UserPrincipalName" -ErrorAction SilentlyContinue } | Where-Object {$_ -ne $null}`

## Script Logic Overview

1.  **Warning & Confirmation:** Prompts the user to confirm the destructive operation.
2.  **Connect to Microsoft Graph:** Authenticates and establishes a session with the required permissions.
3.  **Fetch Users:** Retrieves all users (or a test subset) – initially fetching only `Id` and `UserPrincipalName`.
4.  **Iterate and Process:**
    * For each user, it individually fetches their current `Birthday` property.
    * If the `Birthday` is set (not null), it attempts to clear it by calling `Update-MgUser -UserId $user.Id -BodyParameter @{ "birthday" = $null }`.
    * If the `Birthday` is already clear, it skips the update for that user.
5.  **Logging & Summary:**
    * Provides progress updates in the console.
    * Outputs successes and failures for individual users.
    * A summary of operations (total users, cleared, skipped, failed reads, failed updates) is displayed at the end.
    * A detailed failure log (`BirthdayClear_Failures_YYYYMMDDHHMMSS.txt`) is created if any errors occur.
6.  **Throttling:** Includes a small `Start-Sleep` delay between API calls to help mitigate API throttling. This can be adjusted.

## Important Notes

* **Irreversible Action:** Clearing birthdays is permanent unless you have a backup.
* **SharePoint Dependency:** This script accounts for environments where the `Birthday` attribute is tied to SharePoint User Profiles by including `Sites.ReadWrite.All` permissions and using `-BodyParameter` for the update, which was found to be more reliable.
* **Error Handling:** The script includes `try-catch` blocks for major operations. Review console output and log files for any issues.
* **API Throttling:** For very large tenants, you might need to increase the `$sleepBetweenApiCalls` value or implement more sophisticated retry logic if throttling errors occur.

## Disclaimer

This script is provided "as is" without warranty of any kind. Always test thoroughly in a non-production environment before executing it in a live environment. The user assumes all responsibility for the use of this script.

---
---

## Script 2: Set User Birthdays in Bulk from CSV

**Suggested Script Name:** `Set-EntraUserBirthdaysFromCsv.ps1`

---

**`Readme.md` for `Set-EntraUserBirthdaysFromCsv.ps1`:**

```markdown
# Set-EntraUserBirthdaysFromCsv.ps1

**Date:** May 7, 2025
**Version:** 1.0

## Purpose

This PowerShell script updates the 'Birthday' attribute for Microsoft Entra ID users in bulk, based on data provided in a CSV file. It reads user identifiers and their corresponding birthdays from the CSV, then attempts to set these birthdays in Entra ID.

## Prerequisites

1.  **PowerShell:** Version 5.1 or later.
2.  **Microsoft Graph PowerShell SDK:** The `Microsoft.Graph.Users` module must be installed.
    ```powershell
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
    ```
3.  **Permissions:** The account used to run this script requires the following Microsoft Graph API permissions. You will be prompted to authenticate when the script connects.
    * `User.ReadWrite.All`: To find users and update their `Birthday` attribute.
    * `Sites.ReadWrite.All`: Required because the 'Birthday' attribute in the target environment was found to be linked with SharePoint User Profiles, and updating it via Graph API may require permissions to write to underlying site data.
    Admin consent may be required for these permissions in your tenant.
4.  **CSV File:** A CSV file containing the user data.

## CSV File Format

The CSV file **must** contain the following columns with these exact header names:

1.  `UserPrincipalName`: The User Principal Name (UPN) of the user whose birthday is to be set (e.g., `johndoe@contoso.com`).
2.  `Birthday`: The birth date of the user.
    * **Recommended Format:** `YYYY-MM-DD` (e.g., `1990-05-15`). The script is configured to parse this format by default. If you use a different format, you'll need to adjust the `$csvDateFormat` variable within the script.

**Example `user_birthdays.csv`:**
```csv
UserPrincipalName,Birthday
adelev@contoso.com,1985-03-15
alexw@contoso.com,1992-07-22
lynner@contoso.com,1978-11-05