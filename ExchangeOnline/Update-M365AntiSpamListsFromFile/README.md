# Update Microsoft 365 Anti-Spam Lists from Files

## Overview

This PowerShell script (`Update-M365AntiSpamListsFromFile.ps1`) helps administrators bulk-update the sender and domain allow/block lists within a specified Microsoft 365 inbound anti-spam policy (Hosted Content Filter Policy). It reads entries from simple text files, merges them with the existing entries in the policy, and applies the updates.

This is particularly useful for migrating lists from other systems or for managing large numbers of entries more easily than through the Microsoft Defender portal UI. Entries added to these policy lists do not expire by default.

## Prerequisites

1.  **PowerShell:** Version 5.1 or later.
2.  **ExchangeOnlineManagement Module:** The script will attempt to install this module from the PowerShell Gallery if it's not found (requires internet connectivity and potentially administrator rights depending on scope). You can also install it manually:
    ```powershell
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    ```
3.  **Permissions:** An account with sufficient permissions in Microsoft 365 to connect to Exchange Online and modify anti-spam policies (e.g., Security Administrator, Exchange Administrator).
4.  **Execution Policy:** You may need to adjust your PowerShell execution policy to run scripts. For example:
    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```

## Input Files

The script reads entries from four separate text files. By default, it looks for these files in the same directory where the script is run, but you can specify different paths using parameters.

**File Format:** Each file must contain **one entry per line**. Blank lines are typically ignored, but avoid leading/trailing spaces.

1.  **`allowed_addresses.txt`** (Parameter: `-AllowedAddressesFile`)
    * Contains full email addresses to **allow**.
    * Example:
        ```
        sender1@contoso.com
        newsletter@fabrikam.com
        user@partner.org
        ```
2.  **`allowed_domains.txt`** (Parameter: `-AllowedDomainsFile`)
    * Contains domains from which to **allow** mail.
    * Example:
        ```
        contoso.com
        fabrikam.com
        trustedomain.net
        ```
3.  **`blocked_addresses.txt`** (Parameter: `-BlockedAddressesFile`)
    * Contains full email addresses to **block**.
    * Example:
        ```
        spammer@evil.com
        phisher@bad.net
        ```
4.  **`blocked_domains.txt`** (Parameter: `-BlockedDomainsFile`)
    * Contains domains from which to **block** mail.
    * Example:
        ```
        evil.com
        bad.net
        spamdomain.org
        ```

**Important:** If a file is missing or empty, the script will attempt to **clear** the corresponding list in the policy if it currently contains any entries. Ensure your files contain the complete desired list (including entries already in the policy that you wish to keep) or are intentionally left empty/deleted if you want to clear a list.

## How to Run

1.  Save the script as `Update-M365AntiSpamListsFromFile.ps1`.
2.  Prepare your input text files (`allowed_addresses.txt`, etc.) in the same directory (or note their paths).
3.  Open PowerShell (run as administrator if needed for module installation or execution policy changes).
4.  Navigate to the directory where you saved the script:
    ```powershell
    cd C:\path\to\your\script
    ```
5.  Run the script. You will be prompted to log in to your Microsoft 365 tenant.

**Examples:**

* **Run with defaults (files in same directory, policy name 'Default'):**
    ```powershell
    .\Update-M365AntiSpamListsFromFile.ps1
    ```

* **Specify file paths and a custom policy name:**
    ```powershell
    .\Update-M365AntiSpamListsFromFile.ps1 -PolicyName "My Custom Policy" -AllowedAddressesFile "C:\Data\AllowEmails.txt" -AllowedDomainsFile "C:\Data\AllowDomains.txt" -BlockedAddressesFile "C:\Data\BlockEmails.txt" -BlockedDomainsFile "C:\Data\BlockDomains.txt"
    ```

* **Update only blocked domains and disconnect afterwards:**
    ```powershell
    .\Update-M365AntiSpamListsFromFile.ps1 -BlockedDomainsFile "C:\Data\BlockDomains.txt" -DisconnectWhenDone
    ```
    *(Note: This will clear other lists if their corresponding files are missing/empty and the lists currently have entries in the policy).*

## Notes

* **Merging Logic:** The script reads the current list from the policy, adds the entries from your file, removes duplicates, and then uses `Set-HostedContentFilterPolicy` to apply the complete, merged list.
* **Policy Replacement:** The `Set-HostedContentFilterPolicy` cmdlet *replaces* the entire list for each specified parameter (`AllowedSenders`, `AllowedSenderDomains`, etc.).
* **Verification:** After running the script, verify the changes in the Microsoft Defender portal (Policies & rules > Threat policies > Anti-spam > Your Policy Name).
* **Error Handling:** The script includes basic error handling for module installation, connection, file reading, and policy updates. Review console output for any warnings or errors.
