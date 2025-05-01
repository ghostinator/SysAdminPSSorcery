# Update Microsoft 365 Policy Lists from Files

## Overview

This PowerShell script (`Update-M365PolicyListsFromFile.ps1`) helps administrators bulk-update sender and domain allow/block lists across multiple Microsoft 365 security policies using four primary input files.

It targets the following policies:
* **Anti-Spam Inbound Policy (HostedContentFilterPolicy):** Updates Allowed Senders (Addresses), Allowed Sender Domains, Blocked Senders (Addresses), and Blocked Sender Domains.
* **Tenant Allow/Block List (TABL):** Adds Allowed Senders (combining addresses and domains from input files) and Blocked Senders (combining addresses and domains from input files).
* **Anti-Phishing Policy:** Updates Trusted Senders (Addresses) and Trusted Domains.

The script reads entries from four specified text files (`allowed_addresses.txt`, `allowed_domains.txt`, `blocked_addresses.txt`, `blocked_domains.txt`).

* For policies managed with `Set-*` cmdlets (Anti-Spam, Anti-Phishing), it **merges** the entries from the files with the existing entries in the policy, removing duplicates.
* For the TABL (managed with `New-TenantAllowBlockListItems`), it **attempts to add** the combined entries from the files. Running this multiple times with the same files may result in warnings or errors for entries that already exist, which is generally expected.

**Important:** This script **does not** manage IP Allow/Block lists in the Connection Filter policy. Use other methods (like the Microsoft Defender portal UI or separate PowerShell scripts using `Set-HostedConnectionFilterPolicy`) for managing allowed/blocked IPv4 addresses and CIDR ranges.

## Prerequisites

1.  **PowerShell:** Version 5.1 or later.
2.  **ExchangeOnlineManagement Module:** The script will attempt to install this module from the PowerShell Gallery if it's not found (requires internet connectivity and potentially administrator rights depending on scope). You can also install it manually:
    ```powershell
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
    ```
3.  **Permissions:** An account with sufficient permissions in Microsoft 365 to connect to Exchange Online and modify the relevant security policies (e.g., Security Administrator, Exchange Administrator).
4.  **Execution Policy:** By default, PowerShell might prevent running scripts downloaded from the internet or unsigned local scripts. You may need to adjust the execution policy for your user account.
    * **Check current policy:**
        ```powershell
        Get-ExecutionPolicy -Scope CurrentUser
        ```
    * **Allow local/signed remote scripts (Recommended):**
        ```powershell
        Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
        ```
    * You might be prompted to confirm the change. Answer 'Y' or 'A'.

## Input Files

The script reads entries from four separate text files. By default, it looks for these files in `C:\Temp\`, but you can specify different paths using parameters.

**File Format:** Each file must contain **one entry per line**. Blank lines are ignored. Leading/trailing whitespace is trimmed. Entries are typically treated case-insensitively.

1.  **`allowed_addresses.txt`** (Parameter: `-AllowedAddressesFile`, Default: `C:\Temp\allowed_addresses.txt`)
    * Contains full email addresses to allow/trust.
    * *Applies to:* Anti-Spam Allowed Senders, TABL Allowed Senders, Anti-Phishing Trusted Senders.
    * *Example:*
        ```
        sender1@contoso.com
        newsletter@fabrikam.com
        user@partner.org
        ```
2.  **`allowed_domains.txt`** (Parameter: `-AllowedDomainsFile`, Default: `C:\Temp\allowed_domains.txt`)
    * Contains domains from which to allow/trust mail.
    * *Applies to:* Anti-Spam Allowed Domains, TABL Allowed Senders, Anti-Phishing Trusted Domains.
    * *Example:*
        ```
        contoso.com
        fabrikam.com
        trustedomain.net
        ```
3.  **`blocked_addresses.txt`** (Parameter: `-BlockedAddressesFile`, Default: `C:\Temp\blocked_addresses.txt`)
    * Contains full email addresses to block.
    * *Applies to:* Anti-Spam Blocked Senders, TABL Blocked Senders.
    * *Example:*
        ```
        spammer@evil.com
        phisher@bad.net
        ```
4.  **`blocked_domains.txt`** (Parameter: `-BlockedDomainsFile`, Default: `C:\Temp\blocked_domains.txt`)
    * Contains domains from which to block mail.
    * *Applies to:* Anti-Spam Blocked Domains, TABL Blocked Senders.
    * *Example:*
        ```
        evil.com
        bad.net
        spamdomain.org
        ```

**Important:** For the Anti-Spam and Anti-Phishing policies (which use merging logic): If a file is missing or empty, the script will attempt to **clear** the corresponding list in the policy if that list currently contains any entries. Ensure your files contain the complete desired list (including entries already in the policy that you wish to keep) or are intentionally left empty/deleted if you want to clear a specific list in those policies.

## How to Run

1.  Save the script (e.g., `Update-M365PolicyListsFromFile.ps1`).
2.  Prepare your input text files (`allowed_addresses.txt`, etc.) in `C:\Temp\` (or your chosen location).
3.  Open PowerShell (run as administrator only if needed for module installation or execution policy changes affecting all users).
4.  Check and potentially set your execution policy (see Prerequisites section).
5.  Navigate to the directory where you saved the script (if not running using the full path):
    ```powershell
    cd C:\path\to\your\script
    ```
6.  Run the script. You will be prompted to log in to your Microsoft 365 tenant if not already connected.

**Examples:**

* **Run with defaults (files in C:\Temp, default policy names):**
    ```powershell
    .\Update-M365PolicyListsFromFile.ps1
    ```
    *(Note: This will skip Anti-Phishing updates unless `-AntiPhishPolicyName` is provided).*

* **Specify file paths, policy names, and TABL expiration:**
    ```powershell
    .\Update-M365PolicyListsFromFile.ps1 -AntiSpamPolicyName "Default" -AntiPhishPolicyName "Office 365 AntiPhish Default" -AllowedAddressesFile "C:\Data\AllowEmails.txt" -AllowedDomainsFile "C:\Data\AllowDomains.txt" -BlockedAddressesFile "C:\Data\BlockEmails.txt" -BlockedDomainsFile "C:\Data\BlockDomains.txt" -TablNoExpiration
    ```

* **Update only blocked lists using default file paths and disconnect afterwards:**
    ```powershell
    .\Update-M365PolicyListsFromFile.ps1 -BlockedAddressesFile "C:\Temp\blocked_addresses.txt" -BlockedDomainsFile "C:\Temp\blocked_domains.txt" -DisconnectWhenDone
    ```
    *(Note: This will clear allowed lists in Anti-Spam/Anti-Phishing if their corresponding files are missing/empty and the lists currently have entries in the policy).*

## Notes

* **Merging Logic (Anti-Spam, Anti-Phishing):** The script reads the current list from the policy, adds the unique entries from your file, removes duplicates from the combined list, and then uses `Set-*` cmdlets to apply the complete, merged list.
* **Adding Logic (TABL):** The script uses `New-TenantAllowBlockListItems` to *add* entries. Existing duplicates will generate warnings but won't stop the script.
* **Policy Replacement (Anti-Spam, Anti-Phishing):** The `Set-*` cmdlets *replace* the entire list for each specified parameter (`AllowedSenders`, `TrustedDomains`, etc.).
* **TABL Expiration:** Allowed entries added to TABL default to 30 days expiration unless `-TablExpirationDays` or `-TablNoExpiration` is used. Blocked entries added to TABL default to no expiration unless `-TablExpirationDays` is used.
* **Verification:** After running the script, verify the changes in the Microsoft Defender portal (Policies & rules > Threat policies > Anti-spam / Tenant Allow/Block Lists / Anti-phishing).
* **Error Handling:** The script includes basic error handling. Review console output for any warnings or errors.
