<#
.SYNOPSIS
Updates the Allow/Block lists (Senders/Domains) in a Microsoft 365 Anti-Spam policy using entries from text files.

.DESCRIPTION
This script connects to Exchange Online and modifies a specified inbound Anti-Spam policy (typically the 'Default' one).
It reads lists of allowed email addresses, allowed domains, blocked email addresses, and blocked domains from separate text files.
The script merges the entries from the files with the existing entries in the policy, ensuring duplicates are removed.
It uses the Set-HostedContentFilterPolicy cmdlet, which replaces the existing lists with the new merged lists.
Entries added via this method do not expire by default.

Requires the ExchangeOnlineManagement PowerShell module and appropriate permissions
(e.g., Security Administrator, Exchange Administrator).

.PARAMETER PolicyName
The name of the Hosted Content Filter Policy (Anti-Spam Inbound Policy) to modify.
Defaults to 'Default'. Use Get-HostedContentFilterPolicy to find other policy names.

.PARAMETER AllowedAddressesFile
The full path to the text file containing allowed email addresses (one per line).
Defaults to '.\allowed_addresses.txt' in the script's directory.

.PARAMETER AllowedDomainsFile
The full path to the text file containing allowed domains (one per line).
Defaults to '.\allowed_domains.txt' in the script's directory.

.PARAMETER BlockedAddressesFile
The full path to the text file containing blocked email addresses (one per line).
Defaults to '.\blocked_addresses.txt' in the script's directory.

.PARAMETER BlockedDomainsFile
The full path to the text file containing blocked domains (one per line).
Defaults to '.\blocked_domains.txt' in the script's directory.

.PARAMETER DisconnectWhenDone
Switch parameter. If specified, the script will disconnect the Exchange Online session upon completion.

.EXAMPLE
.\Update-M365AntiSpamListsFromFile.ps1

Description: Runs the script using the default policy name ('Default') and default file names
('allowed_addresses.txt', 'allowed_domains.txt', etc.) expected in the same directory as the script.

.EXAMPLE
.\Update-M365AntiSpamListsFromFile.ps1 -PolicyName "Contoso Strict Policy" -AllowedAddressesFile "C:\Temp\Contoso_Allow_Emails.txt" -AllowedDomainsFile "C:\Temp\Contoso_Allow_Domains.txt" -BlockedAddressesFile "C:\Temp\Contoso_Block_Emails.txt" -BlockedDomainsFile "C:\Temp\Contoso_Block_Domains.txt"

Description: Runs the script targeting a custom policy named "Contoso Strict Policy" and uses specific file paths for the input lists.

.EXAMPLE
.\Update-M365AntiSpamListsFromFile.ps1 -BlockedDomainsFile "C:\Path\OnlyUpdateBlocks.txt" -DisconnectWhenDone

Description: Runs the script updating only the blocked domains list (assuming other files don't exist or are empty)
for the 'Default' policy and disconnects the session afterwards.

.NOTES
Author: Gemini
Version: 1.0
Date: 2025-05-01
- Ensure the input text files contain one entry per line.
- The script replaces the entire list in the policy with the merged list. If a file is empty or missing,
  and the corresponding policy list currently has entries, those entries will be REMOVED (the list will be cleared).
- Requires PowerShell 5.1 or later.
- Run this script in a standard PowerShell console, not PowerShell ISE, for best results with interactive login.
#>
param(
    [Parameter(Mandatory=$false)]
    [string]$PolicyName = "Default",

    [Parameter(Mandatory=$false)]
    [string]$AllowedAddressesFile = ".\allowed_addresses.txt",

    [Parameter(Mandatory=$false)]
    [string]$AllowedDomainsFile = ".\allowed_domains.txt",

    [Parameter(Mandatory=$false)]
    [string]$BlockedAddressesFile = ".\blocked_addresses.txt",

    [Parameter(Mandatory=$false)]
    [string]$BlockedDomainsFile = ".\blocked_domains.txt",

    [Parameter(Mandatory=$false)]
    [switch]$DisconnectWhenDone
)

# --- Prerequisites ---

# Define the required module name
$moduleName = "ExchangeOnlineManagement"

# Check if the module is installed, if not attempt to install it
if (-not (Get-Module -Name $moduleName -ListAvailable)) {
    Write-Host "Module '$moduleName' not found. Attempting to install..." -ForegroundColor Yellow
    try {
        # Use -Scope CurrentUser if admin rights are not available for AllUsers installation
        Install-Module -Name $moduleName -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "Module '$moduleName' installed successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install module '$moduleName'. Please install it manually using 'Install-Module -Name $moduleName -Scope CurrentUser' and try again."
        exit 1 # Exit with an error code
    }
}

# Import the module into the current session
Write-Host "Importing module '$moduleName'..."
try {
    Import-Module -Name $moduleName -Force -ErrorAction Stop
} catch {
    Write-Error "Failed to import module '$moduleName'. Error: $($_.Exception.Message)"
    exit 1
}

# --- Connection ---

# Check if already connected, if not, connect interactively
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online. Please sign in with administrator credentials."
    try {
        # Connect using the interactive method (recommended outside of ISE)
        # Add -UserPrincipalName youradmin@domain.com if needed, but interactive usually prompts
        Connect-ExchangeOnline -ShowBanner:$false -WarningAction SilentlyContinue -ErrorAction Stop
        Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
        Write-Error "Please ensure you have the correct permissions and try again."
        exit 1 # Exit the script if connection fails
    }
} else {
    Write-Host "Already connected to Exchange Online." -ForegroundColor Cyan
}


# --- Initialize Parameter Hashtable for Set Cmdlet ---
$SetParams = @{ Identity = $PolicyName }
$PolicyUpdated = $false # Flag to track if any changes are made

# --- Function to Read File and Merge List ---
function Get-MergedList {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [Parameter(Mandatory=$false)]
        [array]$CurrentList,
        [Parameter(Mandatory=$true)]
        [string]$ListDescription
    )

    # Ensure CurrentList is a usable array, default to empty if $null
    if ($CurrentList -eq $null) {
        $CurrentList = @()
        Write-Verbose "Current $ListDescription list from policy was null/empty."
    }

    try {
        Write-Verbose "Attempting to read $ListDescription from: $FilePath"
        # Check if file exists before trying to read
        if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
             Write-Warning "File not found at '$FilePath'. Skipping $ListDescription."
             # If file not found, check if list needs clearing
             if ($CurrentList.Count -gt 0) {
                  Write-Host "$ListDescription list needs to be cleared (file not found)." -ForegroundColor Yellow
                  return @() # Return empty array signal to clear
             } else {
                  return $null # Indicate no update needed
             }
        }

        $NewEntries = Get-Content -Path $FilePath -ErrorAction Stop
        if (-not $NewEntries) { # Check if file is empty or Get-Content returned null/empty
            Write-Host "$ListDescription file ('$FilePath') is empty. No new entries to add for this list."
            # If file is empty, we might still need to clear the list if CurrentList has items
            if ($CurrentList.Count -gt 0) {
                 Write-Host "$ListDescription list needs to be cleared (file is empty)." -ForegroundColor Yellow
                 return @() # Return empty array signal to clear
            } else {
                 return $null # Indicate no update needed
            }
        } else {
             # Ensure $NewEntries is always an array, even if file has one line
            if ($NewEntries -isnot [array]) { $NewEntries = @($NewEntries) }

            Write-Host "Read $($NewEntries.Count) new entries for $ListDescription from '$FilePath'."
            Write-Host "Current $ListDescription list in policy '$PolicyName' has $($CurrentList.Count) entries."
            # Combine, sort, unique (case-insensitive by default)
            $CombinedList = ($CurrentList + $NewEntries) | Sort-Object -Unique
            Write-Host "Combined unique $ListDescription list has $($CombinedList.Count) entries."
            # Check if the list actually changed
            if (($CombinedList.Count -ne $CurrentList.Count) -or ($CombinedList | Compare-Object $CurrentList -CaseSensitive:$false)) {
                 Write-Host "$ListDescription list has changed." -ForegroundColor Green
                 return $CombinedList
            } else {
                 Write-Host "$ListDescription list has NOT changed."
                 return $null # Indicate no update needed
            }
        }
    } catch {
        Write-Error "Error processing file '$FilePath' for $ListDescription."
        Write-Error "Error Details: $($_.Exception.Message)"
        # Consider whether to stop the whole script or just skip this list
        throw "Failed to process $ListDescription file."
    }
}

# --- Main Logic ---
try {
    # Get the current policy settings
    Write-Host "Fetching current policy '$PolicyName'..." -ForegroundColor Cyan
    $Policy = Get-HostedContentFilterPolicy -Identity $PolicyName -ErrorAction Stop

    # Process Allowed Senders (Addresses)
    $UpdatedAllowedSenders = Get-MergedList -FilePath $AllowedAddressesFile -CurrentList $Policy.AllowedSenders -ListDescription "Allowed Senders (Addresses)"
    if ($UpdatedAllowedSenders -ne $null) {
        $SetParams.AllowedSenders = $UpdatedAllowedSenders
        $PolicyUpdated = $true
    }

    # Process Allowed Sender Domains
    $UpdatedAllowedDomains = Get-MergedList -FilePath $AllowedDomainsFile -CurrentList $Policy.AllowedSenderDomains -ListDescription "Allowed Sender Domains"
    if ($UpdatedAllowedDomains -ne $null) {
        $SetParams.AllowedSenderDomains = $UpdatedAllowedDomains
        $PolicyUpdated = $true
    }

    # Process Blocked Senders (Addresses)
    $UpdatedBlockedSenders = Get-MergedList -FilePath $BlockedAddressesFile -CurrentList $Policy.BlockedSenders -ListDescription "Blocked Senders (Addresses)"
    if ($UpdatedBlockedSenders -ne $null) {
        $SetParams.BlockedSenders = $UpdatedBlockedSenders
        $PolicyUpdated = $true
    }

    # Process Blocked Sender Domains
    $UpdatedBlockedDomains = Get-MergedList -FilePath $BlockedDomainsFile -CurrentList $Policy.BlockedSenderDomains -ListDescription "Blocked Sender Domains"
    if ($UpdatedBlockedDomains -ne $null) {
        $SetParams.BlockedSenderDomains = $UpdatedBlockedDomains
        $PolicyUpdated = $true
    }

    # --- Update the policy only if changes were needed ---
    if ($PolicyUpdated) {
        Write-Host "Updating policy '$PolicyName' with new parameters..." -ForegroundColor Cyan
        Write-Host "Parameters being set:"
        $SetParams | Format-List | Out-String | Write-Host # Show what's being set clearly

        Set-HostedContentFilterPolicy @SetParams -ErrorAction Stop
        Write-Host "Successfully updated policy '$PolicyName'." -ForegroundColor Green
    } else {
        Write-Host "No changes detected in any list compared to the files. Policy '$PolicyName' not updated." -ForegroundColor Yellow
    }

} catch {
    Write-Error "Failed to get or update policy '$PolicyName'."
    Write-Error "Error Details: $($_.Exception.Message)"
    # Consider adding more specific error handling if needed
} finally {
    # --- Optional Disconnect ---
    if ($DisconnectWhenDone) {
        Write-Host "Disconnecting from Exchange Online as requested..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected."
    }
}

Write-Host "Script finished."
