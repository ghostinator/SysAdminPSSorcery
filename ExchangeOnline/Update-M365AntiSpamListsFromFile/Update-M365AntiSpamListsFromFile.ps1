<#
.SYNOPSIS
Updates Allow/Block lists across multiple Microsoft 365 security policies using four primary sender/domain text files.

.DESCRIPTION
This script connects to Exchange Online and modifies specified security policies using four input files:
- Anti-Spam Inbound Policy (HostedContentFilterPolicy): Updates Allowed/Blocked Senders (Addresses/Domains).
- Tenant Allow/Block List (TABL): Adds Allowed/Blocked Senders (combining addresses and domains from files).
- Anti-Phishing Policy: Updates Trusted Senders (Addresses) and Trusted Domains.

It reads lists of entries from the four specified text files (allowed/blocked addresses and allowed/blocked domains).
For policies managed with 'Set-*' cmdlets (Anti-Spam, Anti-Phishing), it merges the entries from the files
with existing entries, removing duplicates.
For the TABL (managed with 'New-TenantAllowBlockListItems'), it attempts to add the combined entries from the files.
Note that running the TABL update repeatedly with the same files may cause errors for duplicate entries.

This script DOES NOT manage IP Allow/Block lists in the Connection Filter policy. Use other methods for that.

Requires the ExchangeOnlineManagement PowerShell module and appropriate permissions
(e.g., Security Administrator, Exchange Administrator).

.PARAMETER AntiSpamPolicyName
The name of the Hosted Content Filter Policy (Anti-Spam Inbound Policy) to modify. Defaults to 'Default'.

.PARAMETER AntiPhishPolicyName
The name of the Anti-Phishing Policy to modify. *This parameter is required* if you provide files for allowed addresses or domains.
Use Get-AntiPhishPolicy to find policy names.

.PARAMETER AllowedAddressesFile
Path to the text file containing allowed email addresses (one per line). Applied to Anti-Spam, TABL, Anti-Phishing.
Defaults to 'C:\Temp\allowed_addresses.txt'.

.PARAMETER AllowedDomainsFile
Path to the text file containing allowed domains (one per line). Applied to Anti-Spam, TABL, Anti-Phishing.
Defaults to 'C:\Temp\allowed_domains.txt'.

.PARAMETER BlockedAddressesFile
Path to the text file containing blocked email addresses (one per line). Applied to Anti-Spam, TABL.
Defaults to 'C:\Temp\blocked_addresses.txt'.

.PARAMETER BlockedDomainsFile
Path to the text file containing blocked domains (one per line). Applied to Anti-Spam, TABL.
Defaults to 'C:\Temp\blocked_domains.txt'.

.PARAMETER TablExpirationDays
The number of days after which TABL ALLOW entries should expire. If not specified, and -TablNoExpiration is not used, defaults to 30 days internally. Cannot be used with -TablNoExpiration.
Blocks added to TABL via this script default to No Expiration unless this parameter is used.

.PARAMETER TablNoExpiration
Switch parameter. If specified, TABL ALLOW entries will be set to never expire. Cannot be used with -TablExpirationDays.

.PARAMETER TablNotes
Optional notes to add to the TABL entries being created. Defaults to 'Bulk update via PowerShell script'.

.PARAMETER DisconnectWhenDone
Switch parameter. If specified, the script will disconnect the Exchange Online session upon completion.

.EXAMPLE
.\Update-M365PolicyListsFromFile_Consolidated.ps1 -AllowedAddressesFile "C:\Data\AllowEmails.txt" -AllowedDomainsFile "C:\Data\AllowDomains.txt" -BlockedAddressesFile "C:\Data\BlockEmails.txt" -BlockedDomainsFile "C:\Data\BlockDomains.txt" -AntiPhishPolicyName "Office 365 AntiPhish Default" -TablNoExpiration

Description: Updates Anti-Spam, TABL, and Anti-Phishing policies using the four specified files from C:\Data. Sets TABL allows to not expire.

.EXAMPLE
.\Update-M365PolicyListsFromFile_Consolidated.ps1

Description: Attempts to run using default file names from C:\Temp for the default Anti-Spam policy.
TABL Allows will default to 30-day expiration. Will likely skip Anti-Phishing updates unless files exist
and -AntiPhishPolicyName is provided.

.NOTES
Author: Gemini
Version: 3.7 (Handle Policy Object Types in Merge Function)
Date: 2025-05-01
- Ensure the input text files contain one entry per line.
- For Anti-Spam and Anti-Phishing policies, the script merges file entries with existing policy entries.
  If a file is empty/missing, and the corresponding policy list has entries, the list will be CLEARED.
- For TABL entries, the script uses New-TenantAllowBlockListItems. Running this multiple times with the same file
  may result in errors for entries that already exist.
- This script does NOT manage Connection Filter (IP Allow/Block) lists.
- Requires PowerShell 5.1 or later.
- Run this script in a standard PowerShell console.
#>
param(
    # Policy Names
    [Parameter(Mandatory=$false)]
    [string]$AntiSpamPolicyName = "Default",

    [Parameter(Mandatory=$false)] # Mandatory only if updating AntiPhish lists
    [string]$AntiPhishPolicyName,

    # Consolidated Input Files (Defaults updated)
    [Parameter(Mandatory=$false)]
    [string]$AllowedAddressesFile = "C:\Temp\allowed_addresses.txt",

    [Parameter(Mandatory=$false)]
    [string]$AllowedDomainsFile = "C:\Temp\allowed_domains.txt",

    [Parameter(Mandatory=$false)]
    [string]$BlockedAddressesFile = "C:\Temp\blocked_addresses.txt",

    [Parameter(Mandatory=$false)]
    [string]$BlockedDomainsFile = "C:\Temp\blocked_domains.txt",

    # TABL Options (No Parameter Sets defined here)
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 90)] # Max expiration for most TABL allows is 90 days
    [int]$TablExpirationDays,

    [Parameter(Mandatory=$false)]
    [switch]$TablNoExpiration,

    [Parameter(Mandatory=$false)]
    [string]$TablNotes = "Bulk update via PowerShell script",

    # General Options
    [Parameter(Mandatory=$false)]
    [switch]$DisconnectWhenDone
)

# --- Validation ---
# Manually check if both expiration parameters were used
if ($PSBoundParameters.ContainsKey('TablExpirationDays') -and $PSBoundParameters.ContainsKey('TablNoExpiration')) {
    Write-Error "Parameters -TablExpirationDays and -TablNoExpiration cannot be used together. Please choose one."
    exit 1
}


# --- Function to Read Entries from a File ---
function Get-EntriesFromFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [Parameter(Mandatory=$true)]
        [string]$ListDescription
    )
    if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
        Write-Warning "File not found at '$FilePath'. Skipping $ListDescription."
        return $null # Return null if file doesn't exist
    }
    try {
        $Entries = Get-Content -Path $FilePath -ErrorAction Stop
        if (-not $Entries) {
            Write-Host "$ListDescription file ('$FilePath') is empty."
            return @() # Return empty array if file is empty
        } else {
            if ($Entries -isnot [array]) { $Entries = @($Entries) }
            # Trim whitespace and convert to lowercase for consistency (especially domains/emails)
            $CleanedEntries = $Entries | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_.Length -gt 0 }
            # Get unique entries from the file itself first
            $UniqueEntries = $CleanedEntries | Sort-Object -Unique
            Write-Host "Read $($UniqueEntries.Count) unique entries for $ListDescription from '$FilePath'."
            return $UniqueEntries
        }
    } catch {
        Write-Error "Error reading file '$FilePath' for $ListDescription."
        Write-Error "Error Details: $($_.Exception.Message)"
        throw "Failed to read $ListDescription file." # Stop script if file read fails
    }
}

# --- Prerequisites ---
Write-Host "Starting Script: Update M365 Policy Lists (Consolidated Inputs)" -ForegroundColor Cyan
#region Prerequisites and Connection

# Define the required module name
$moduleName = "ExchangeOnlineManagement"

# Check if the module is installed, if not attempt to install it
if (-not (Get-Module -Name $moduleName -ListAvailable)) {
    Write-Host "Module '$moduleName' not found. Attempting to install..." -ForegroundColor Yellow
    try {
        Install-Module -Name $moduleName -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "Module '$moduleName' installed successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install module '$moduleName'. Please install it manually using 'Install-Module -Name $moduleName -Scope CurrentUser' and try again."
        exit 1
    }
}

# Import the module into the current session
Write-Host "Importing module '$moduleName'..."
try {
    # Removed -ErrorAction Stop as requested by user
    Import-Module -Name $moduleName -Force
} catch {
    Write-Error "Failed to import module '$moduleName'. Error: $($_.Exception.Message)"
    # Exit if import fails critically, even without ErrorAction Stop
    exit 1
}

# Check if already connected, if not, connect interactively
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online. Please sign in with administrator credentials."
    try {
        # Removed -ErrorAction Stop as requested by user
        Connect-ExchangeOnline -ShowBanner:$false -WarningAction SilentlyContinue
        Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
        Write-Error "Please ensure you have the correct permissions and try again."
        exit 1
    }
} else {
    Write-Host "Already connected to Exchange Online." -ForegroundColor Cyan
}
#endregion Prerequisites and Connection

# --- Read All Input Files ---
# Note: Get-EntriesFromFile now returns unique, trimmed, lowercase entries
Write-Host "`n--- Reading Input Files ---" -ForegroundColor Cyan
$AllowedAddresses = Get-EntriesFromFile -FilePath $AllowedAddressesFile -ListDescription "Allowed Addresses"
$AllowedDomains = Get-EntriesFromFile -FilePath $AllowedDomainsFile -ListDescription "Allowed Domains"
$BlockedAddresses = Get-EntriesFromFile -FilePath $BlockedAddressesFile -ListDescription "Blocked Addresses"
$BlockedDomains = Get-EntriesFromFile -FilePath $BlockedDomainsFile -ListDescription "Blocked Domains"

# --- Helper Function for Merging (Used for Set-* Cmdlets) ---
#region Helper Function: Get-MergedList
function Get-MergedList {
    param(
        [Parameter(Mandatory=$true)]
        [array]$NewEntries, # Unique entries read from file
        [Parameter(Mandatory=$false)]
        [array]$CurrentListRaw, # Current list from policy (might contain objects)
        [Parameter(Mandatory=$true)]
        [string]$ListDescription
        # Case sensitivity handled by ToLowerInvariant in Get-EntriesFromFile
    )

    # Ensure CurrentList is a usable array of STRINGS and also lowercase/trimmed
    $CurrentListNormalized = @() # Initialize empty array for normalized strings
    if ($CurrentListRaw -ne $null) {
         # *** FIX START: Convert policy objects to strings before processing ***
         $CurrentListNormalized = $CurrentListRaw | ForEach-Object { ($_.ToString()).Trim().ToLowerInvariant() } | Where-Object { $_.Length -gt 0 } | Sort-Object -Unique
         # *** FIX END ***
    }

    Write-Host "Current $ListDescription list in policy has $($CurrentListNormalized.Count) unique, normalized entries."

    # Combine the already unique new entries (from file) with the unique, normalized current list (from policy)
    $CombinedList = ($CurrentListNormalized + $NewEntries) | Sort-Object -Unique
    # The result of Sort-Object -Unique is already an array, ensure no empty strings remain just in case
    $CombinedList = $CombinedList | Where-Object { $_.Length -gt 0 }


    Write-Host "Combined unique $ListDescription list has $($CombinedList.Count) entries."

    # Create temporary lists for comparison to accurately detect changes
    $TempCurrent = $CurrentListNormalized | Sort-Object
    $TempCombined = $CombinedList | Sort-Object

    # Check if the list actually changed
    # Compare normalized lists
    if (($CombinedList.Count -ne $CurrentListNormalized.Count) -or (Compare-Object -ReferenceObject $TempCurrent -DifferenceObject $TempCombined -CaseSensitive:$false -SyncWindow 0)) {
         Write-Host "$ListDescription list has changed." -ForegroundColor Green
         # Return the final, unique, sorted list of strings
         return $CombinedList
    } else {
         Write-Host "$ListDescription list has NOT changed."
         return $null # Indicate no update needed
    }
}
#endregion Helper Function: Get-MergedList


# --- Policy Update Logic ---
# *** Wrap main logic in a try block for the finally ***
try {

    #region Anti-Spam Policy (HostedContentFilterPolicy)
    # Only proceed if relevant files were found/read
    if ($AllowedAddresses -ne $null -or $AllowedDomains -ne $null -or $BlockedAddresses -ne $null -or $BlockedDomains -ne $null) {
        Write-Host "`n--- Processing Anti-Spam Policy ($AntiSpamPolicyName) ---" -ForegroundColor Cyan
        $AntiSpamPolicyUpdated = $false
        $AntiSpamSetParams = @{ Identity = $AntiSpamPolicyName }
        try {
            $AntiSpamPolicy = Get-HostedContentFilterPolicy -Identity $AntiSpamPolicyName -ErrorAction Stop

            # Process Allowed Senders (Addresses)
            if ($AllowedAddresses -ne $null) {
                # Pass the raw policy list to the merge function
                $UpdatedASAllowedSenders = Get-MergedList -NewEntries $AllowedAddresses -CurrentListRaw $AntiSpamPolicy.AllowedSenders -ListDescription "Anti-Spam Allowed Senders (Addresses)"
                if ($UpdatedASAllowedSenders -ne $null) { $AntiSpamSetParams.AllowedSenders = $UpdatedASAllowedSenders; $AntiSpamPolicyUpdated = $true }
            } elseif ($AntiSpamPolicy.AllowedSenders -ne $null -and $AntiSpamPolicy.AllowedSenders.Count -gt 0) { # File missing/empty, clear if needed (check raw count)
                 Write-Host "Anti-Spam Allowed Senders (Addresses) list needs to be cleared (file missing/empty)." -ForegroundColor Yellow
                 $AntiSpamSetParams.AllowedSenders = @(); $AntiSpamPolicyUpdated = $true
            }

            # Process Allowed Sender Domains
            if ($AllowedDomains -ne $null) {
                $UpdatedASAllowedDomains = Get-MergedList -NewEntries $AllowedDomains -CurrentListRaw $AntiSpamPolicy.AllowedSenderDomains -ListDescription "Anti-Spam Allowed Sender Domains"
                if ($UpdatedASAllowedDomains -ne $null) { $AntiSpamSetParams.AllowedSenderDomains = $UpdatedASAllowedDomains; $AntiSpamPolicyUpdated = $true }
            } elseif ($AntiSpamPolicy.AllowedSenderDomains -ne $null -and $AntiSpamPolicy.AllowedSenderDomains.Count -gt 0) {
                 Write-Host "Anti-Spam Allowed Sender Domains list needs to be cleared (file missing/empty)." -ForegroundColor Yellow
                 $AntiSpamSetParams.AllowedSenderDomains = @(); $AntiSpamPolicyUpdated = $true
            }

            # Process Blocked Senders (Addresses)
            if ($BlockedAddresses -ne $null) {
                $UpdatedASBlockedSenders = Get-MergedList -NewEntries $BlockedAddresses -CurrentListRaw $AntiSpamPolicy.BlockedSenders -ListDescription "Anti-Spam Blocked Senders (Addresses)"
                if ($UpdatedASBlockedSenders -ne $null) { $AntiSpamSetParams.BlockedSenders = $UpdatedASBlockedSenders; $AntiSpamPolicyUpdated = $true }
            } elseif ($AntiSpamPolicy.BlockedSenders -ne $null -and $AntiSpamPolicy.BlockedSenders.Count -gt 0) {
                 Write-Host "Anti-Spam Blocked Senders (Addresses) list needs to be cleared (file missing/empty)." -ForegroundColor Yellow
                 $AntiSpamSetParams.BlockedSenders = @(); $AntiSpamPolicyUpdated = $true
            }

            # Process Blocked Sender Domains
            if ($BlockedDomains -ne $null) {
                $UpdatedASBlockedDomains = Get-MergedList -NewEntries $BlockedDomains -CurrentListRaw $AntiSpamPolicy.BlockedSenderDomains -ListDescription "Anti-Spam Blocked Sender Domains"
                if ($UpdatedASBlockedDomains -ne $null) { $AntiSpamSetParams.BlockedSenderDomains = $UpdatedASBlockedDomains; $AntiSpamPolicyUpdated = $true }
            } elseif ($AntiSpamPolicy.BlockedSenderDomains -ne $null -and $AntiSpamPolicy.BlockedSenderDomains.Count -gt 0) {
                 Write-Host "Anti-Spam Blocked Sender Domains list needs to be cleared (file missing/empty)." -ForegroundColor Yellow
                 $AntiSpamSetParams.BlockedSenderDomains = @(); $AntiSpamPolicyUpdated = $true
            }

            # Update the policy only if changes were needed
            if ($AntiSpamPolicyUpdated) {
                Write-Host "Updating Anti-Spam policy '$AntiSpamPolicyName'..." -ForegroundColor Yellow
                Set-HostedContentFilterPolicy @AntiSpamSetParams -ErrorAction Stop
                Write-Host "Successfully updated Anti-Spam policy '$AntiSpamPolicyName'." -ForegroundColor Green
            } else {
                Write-Host "No changes detected for Anti-Spam policy '$AntiSpamPolicyName'."
            }
        } catch {
            Write-Error "Failed to get or update Anti-Spam policy '$AntiSpamPolicyName'."
            Write-Error "Error Details: $($_.Exception.Message)"
        }
    } else {
         Write-Host "`n--- Skipping Anti-Spam Policy ($AntiSpamPolicyName) --- (No relevant input files found/read)" -ForegroundColor Gray
    }
    #endregion Anti-Spam Policy (HostedContentFilterPolicy)


    #region Tenant Allow/Block List (TABL)
    # Combine addresses and domains for TABL
    $TablAllowEntries = @()
    if ($AllowedAddresses -ne $null) { $TablAllowEntries += $AllowedAddresses }
    if ($AllowedDomains -ne $null) { $TablAllowEntries += $AllowedDomains }
    $TablAllowEntries = $TablAllowEntries | Sort-Object -Unique # Ensure combined list is unique

    $TablBlockEntries = @()
    if ($BlockedAddresses -ne $null) { $TablBlockEntries += $BlockedAddresses }
    if ($BlockedDomains -ne $null) { $TablBlockEntries += $BlockedDomains }
    $TablBlockEntries = $TablBlockEntries | Sort-Object -Unique # Ensure combined list is unique

    if ($TablAllowEntries.Count -gt 0 -or $TablBlockEntries.Count -gt 0) {
        Write-Host "`n--- Processing Tenant Allow/Block List (TABL) ---" -ForegroundColor Cyan

        # Determine Expiration Date or NoExpiration switch for Allows
        $TablAllowExpirationParams = @{}
        if ($TablNoExpiration.IsPresent) {
            $TablAllowExpirationParams.NoExpiration = $true
            Write-Host "TABL Allow entries will attempt creation with No Expiration."
        } elseif ($PSBoundParameters.ContainsKey('TablExpirationDays')) {
            $CalculatedExpirationDate = (Get-Date).AddDays($TablExpirationDays)
            $TablAllowExpirationParams.ExpirationDate = $CalculatedExpirationDate
            Write-Host "TABL Allow entries will attempt creation expiring on $($CalculatedExpirationDate) (UTC) (using provided days)."
        } else {
            $DefaultExpirationDays = 30
            $CalculatedExpirationDate = (Get-Date).AddDays($DefaultExpirationDays)
            $TablAllowExpirationParams.ExpirationDate = $CalculatedExpirationDate
            Write-Host "TABL Allow entries will attempt creation expiring on $($CalculatedExpirationDate) (UTC) (using default $DefaultExpirationDays days)."
        }

        # Process TABL Allowed Senders
        if ($TablAllowEntries.Count -gt 0) {
            Write-Host "Processing TABL Allowed Senders (Combined Addresses/Domains)..."
            try {
                $TablAllowParams = @{
                    ListType = "Sender"
                    Entries  = $TablAllowEntries # Pass the already unique list
                    Notes    = $TablNotes
                    Allow    = $true
                    ErrorAction = "SilentlyContinue"
                    WarningAction = "Continue"
                }
                $TablAllowParams += $TablAllowExpirationParams

                Write-Host "Attempting to add $($TablAllowEntries.Count) entries to TABL Allow List..." -ForegroundColor Yellow
                $Result = New-TenantAllowBlockListItems @TablAllowParams
                if ($Error.Count -gt 0) {
                     Write-Warning "Some TABL allow entries may not have been added (e.g., duplicates exist). Check previous errors."
                     $Error.Clear()
                } else {
                    Write-Host "Successfully submitted request to add allowed senders to TABL." -ForegroundColor Green
                }
            } catch {
                Write-Error "Failed to process TABL Allowed Senders."
                Write-Error "Error Details: $($_.Exception.Message)"
            }
        } else {
            Write-Host "No combined Allowed Addresses/Domains found to add to TABL."
        }

        # Process TABL Blocked Senders
        if ($TablBlockEntries.Count -gt 0) {
            Write-Host "Processing TABL Blocked Senders (Combined Addresses/Domains)..."
            try {
                 # Determine Expiration for Blocks
                 $TablBlockExpirationParams = @{}
                 if ($TablNoExpiration.IsPresent) {
                     $TablBlockExpirationParams.NoExpiration = $true
                 } elseif ($PSBoundParameters.ContainsKey('TablExpirationDays')) {
                     $CalculatedExpirationDate = (Get-Date).AddDays($TablExpirationDays)
                     $TablBlockExpirationParams.ExpirationDate = $CalculatedExpirationDate
                     Write-Host "Applying specified expiration date to TABL Block entries as well." -ForegroundColor Yellow
                 } else {
                     $TablBlockExpirationParams.NoExpiration = $true # Default Block = No Expiration
                 }

                $TablBlockParams = @{
                    ListType = "Sender"
                    Entries  = $TablBlockEntries # Pass the already unique list
                    Notes    = $TablNotes
                    Block    = $true
                    ErrorAction = "SilentlyContinue"
                    WarningAction = "Continue"
                }
                $TablBlockParams += $TablBlockExpirationParams

                Write-Host "Attempting to add $($TablBlockEntries.Count) entries to TABL Block List..." -ForegroundColor Yellow
                $Result = New-TenantAllowBlockListItems @TablBlockParams
                if ($Error.Count -gt 0) {
                     Write-Warning "Some TABL block entries may not have been added (e.g., duplicates exist). Check previous errors."
                     $Error.Clear()
                } else {
                    Write-Host "Successfully submitted request to add blocked senders to TABL." -ForegroundColor Green
                }
            } catch {
                Write-Error "Failed to process TABL Blocked Senders."
                Write-Error "Error Details: $($_.Exception.Message)"
            }
        } else {
            Write-Host "No combined Blocked Addresses/Domains found to add to TABL."
        }
    } else {
         Write-Host "`n--- Skipping Tenant Allow/Block List (TABL) --- (No relevant input files found/read)" -ForegroundColor Gray
    }
    #endregion Tenant Allow/Block List (TABL)


    #region Anti-Phishing Policy
    # Only proceed if relevant files were found AND policy name was provided
    if (($AllowedAddresses -ne $null -or $AllowedDomains -ne $null) -and (-not [string]::IsNullOrEmpty($AntiPhishPolicyName))) {
        Write-Host "`n--- Processing Anti-Phishing Policy ($AntiPhishPolicyName) ---" -ForegroundColor Cyan
        $AntiPhishPolicyUpdated = $false
        $AntiPhishSetParams = @{ Identity = $AntiPhishPolicyName }
        try {
            # Get-AntiPhishPolicy is needed to get the lists
            $AntiPhishPolicy = Get-AntiPhishPolicy -Identity $AntiPhishPolicyName -ErrorAction Stop

            # Process Trusted Senders (Addresses)
            if ($AllowedAddresses -ne $null) {
                 # Pass the raw policy list to the merge function
                 $UpdatedAPTrustedSenders = Get-MergedList -NewEntries $AllowedAddresses -CurrentListRaw $AntiPhishPolicy.TrustedSenders -ListDescription "Anti-Phishing Trusted Senders (Addresses)"
                 if ($UpdatedAPTrustedSenders -ne $null) { $AntiPhishSetParams.TrustedSenders = $UpdatedAPTrustedSenders; $AntiPhishPolicyUpdated = $true }
            } elseif ($AntiPhishPolicy.TrustedSenders -ne $null -and $AntiPhishPolicy.TrustedSenders.Count -gt 0) {
                 Write-Host "Anti-Phishing Trusted Senders list needs to be cleared (file missing/empty)." -ForegroundColor Yellow
                 $AntiPhishSetParams.TrustedSenders = @(); $AntiPhishPolicyUpdated = $true
            }

            # Process Trusted Domains
            if ($AllowedDomains -ne $null) {
                 # Pass the raw policy list to the merge function
                 $UpdatedAPTrustedDomains = Get-MergedList -NewEntries $AllowedDomains -CurrentListRaw $AntiPhishPolicy.TrustedDomains -ListDescription "Anti-Phishing Trusted Domains"
                 if ($UpdatedAPTrustedDomains -ne $null) { $AntiPhishSetParams.TrustedDomains = $UpdatedAPTrustedDomains; $AntiPhishPolicyUpdated = $true }
            } elseif ($AntiPhishPolicy.TrustedDomains -ne $null -and $AntiPhishPolicy.TrustedDomains.Count -gt 0) {
                 Write-Host "Anti-Phishing Trusted Domains list needs to be cleared (file missing/empty)." -ForegroundColor Yellow
                 $AntiPhishSetParams.TrustedDomains = @(); $AntiPhishPolicyUpdated = $true
            }

            # Update the policy only if changes were needed
            if ($AntiPhishPolicyUpdated) {
                Write-Host "Updating Anti-Phishing policy '$AntiPhishPolicyName'..." -ForegroundColor Yellow
                # Use Set-AntiPhishPolicy for these lists
                Set-AntiPhishPolicy @AntiPhishSetParams -ErrorAction Stop
                Write-Host "Successfully updated Anti-Phishing policy '$AntiPhishPolicyName'." -ForegroundColor Green
            } else {
                Write-Host "No changes detected for Anti-Phishing policy '$AntiPhishPolicyName'."
            }

        } catch {
            Write-Error "Failed to get or update Anti-Phishing policy '$AntiPhishPolicyName'."
            Write-Error "Error Details: $($_.Exception.Message)"
        }
    } elseif (($AllowedAddresses -ne $null -or $AllowedDomains -ne $null) -and ([string]::IsNullOrEmpty($AntiPhishPolicyName))) {
         Write-Warning "`n--- Skipping Anti-Phishing Policy --- (Relevant input files found, but -AntiPhishPolicyName parameter was not provided)"
    } else {
         Write-Host "`n--- Skipping Anti-Phishing Policy --- (No relevant input files found/read)" -ForegroundColor Gray
    }
    #endregion Anti-Phishing Policy

# *** End of main try block ***
}
finally {
    # --- Finalization ---
    #region Disconnect
    # Disconnect if requested
    if ($DisconnectWhenDone.IsPresent) { # Check switch value correctly
        Write-Host "`nDisconnecting from Exchange Online as requested..." -ForegroundColor Cyan
        # Check connection state before disconnecting
        if (Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host "Disconnected."
        } else {
            Write-Host "Already disconnected or connection not found."
        }
    }
    #endregion Disconnect
} # *** End of finally block ***

Write-Host "`nScript finished." -ForegroundColor Cyan
