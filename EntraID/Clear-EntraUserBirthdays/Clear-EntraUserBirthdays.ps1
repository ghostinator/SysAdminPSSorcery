#Requires -Modules Microsoft.Graph.Users

# Script to Clear the Birthday Field for All Entra ID Users
# Created Date: Wednesday, May 7, 2025
# !! HIGHLY CRITICAL SCRIPT - USE WITH EXTREME CAUTION !!
# !! ENSURE YOU HAVE BACKUPS AND HAVE TESTED IN A NON-PRODUCTION ENVIRONMENT !!

# --- Configuration ---
$sleepBetweenApiCalls = 0.2 # Seconds to sleep between API calls (both read and update)
                          # Increase if you face throttling.

# --- Confirmation ---
Write-Host "WARNING: This script will attempt to clear the 'Birthday' field for ALL users in your Entra ID tenant." -ForegroundColor Yellow
Write-Host "This action is irreversible without a prior backup of the birthday data." -ForegroundColor Yellow
Write-Host "It is STRONGLY recommended to test this on a few non-critical accounts first." -ForegroundColor Yellow
Write-Host "Ensure you have the required permissions: User.ReadWrite.All and Sites.ReadWrite.All." -ForegroundColor Yellow

$confirmation = Read-Host "Are you absolutely sure you want to proceed? (Type 'YES' to continue)"
if ($confirmation -ne 'YES') {
    Write-Host "Operation cancelled by the user." -ForegroundColor Green
    exit
}

# --- Script Body ---
Write-Host "Proceeding with operation to clear birthday fields..."

# Connect to Microsoft Graph with necessary write permissions
Write-Host "Connecting to Microsoft Graph..."
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Sites.ReadWrite.All" -NoWelcome -ErrorAction Stop
}
catch {
    Write-Error "Failed to initiate connection to Microsoft Graph. Error: $($_.Exception.Message)"
    exit
}

if (-not (Get-MgContext)) {
    Write-Error "Connection to Microsoft Graph failed or context is not available. Ensure you authenticated successfully."
    exit
}
Write-Host "Successfully connected to Microsoft Graph as `"$((Get-MgContext).Account)`" to tenant `"$((Get-MgContext).TenantId)`"."

Write-Host "Retrieving all users (ID and UPN only initially)..."
$processedUsersCount = 0
$birthdaysClearedCount = 0
$alreadyClearOrSkippedCount = 0
$failedUpdateCount = 0
$failedReadCount = 0
$failedUsersList = [System.Collections.Generic.List[string]]::new()

try {
    # Step 1: Get all users - only Id and UserPrincipalName. This should be reliable.
    $allUsers = Get-MgUser -All -Property "Id,UserPrincipalName" -ErrorAction Stop
    
    $totalUsersToProcess = $allUsers.Count
    if ($totalUsersToProcess -eq 0) {
        Write-Warning "No users found to process."
        exit
    }
    Write-Host "Found $totalUsersToProcess users. Starting to process birthday fields..."

    foreach ($baseUser in $allUsers) {
        $processedUsersCount++
        Write-Progress -Activity "Clearing Birthday Fields" -Status "Processing User: $($baseUser.UserPrincipalName) ($processedUsersCount of $totalUsersToProcess)" -PercentComplete (($processedUsersCount / $totalUsersToProcess) * 100)

        $currentBirthday = $null
        $birthdayReadSuccess = $false

        # Step 2: Fetch current birthday for this specific user
        try {
            # Write-Host "Checking birthday for '$($baseUser.UserPrincipalName)'..." # Can be noisy, enable for debugging
            $userWithBirthday = Get-MgUser -UserId $baseUser.Id -Property "Birthday" -ErrorAction Stop
            if ($null -ne $userWithBirthday) {
                $currentBirthday = $userWithBirthday.Birthday
                $birthdayReadSuccess = $true
            } else {
                 Write-Warning "Could not retrieve user details (for birthday check) for '$($baseUser.UserPrincipalName)' (ID: $($baseUser.Id)). Get-MgUser returned null."
                 $failedReadCount++
            }
            Start-Sleep -Seconds $sleepBetweenApiCalls
        } catch {
            Write-Warning "Failed to read birthday for '$($baseUser.UserPrincipalName)' (ID: $($baseUser.Id)). Error: $($_.Exception.Message)"
            $failedReadCount++
            $failedUsersList.Add("ReadFail: " + $baseUser.UserPrincipalName + " - " + $_.Exception.Message)
            Start-Sleep -Seconds $sleepBetweenApiCalls
            continue # Skip to next user if we can't read their birthday
        }

        # Step 3: If birthday is set, attempt to clear it
        if ($birthdayReadSuccess) {
            if ($null -eq $currentBirthday) {
                # Write-Host "Birthday for user '$($baseUser.UserPrincipalName)' is already clear. Skipping update." # Can be noisy
                $alreadyClearOrSkippedCount++
            } else {
                try {
                    Write-Host "Attempting to clear birthday (current: $currentBirthday) for user '$($baseUser.UserPrincipalName)'..."
                    # Use -BodyParameter to explicitly set the 'birthday' property to null
                    Update-MgUser -UserId $baseUser.Id -BodyParameter @{ "birthday" = $null } -ErrorAction Stop
                    
                    Write-Host "Successfully cleared birthday for user '$($baseUser.UserPrincipalName)'." -ForegroundColor Green
                    $birthdaysClearedCount++
                }
                catch {
                    Write-Error "Failed to clear birthday for user '$($baseUser.UserPrincipalName)' (ID: $($baseUser.Id)). Error: $($_.Exception.Message)"
                    $failedUsersList.Add("UpdateFail: " + $baseUser.UserPrincipalName + " - " + $_.Exception.Message)
                    $failedUpdateCount++
                }
            }
        }
        
        Start-Sleep -Seconds $sleepBetweenApiCalls # Pause after update attempt or skip
    }
}
catch {
    Write-Error "An unexpected error occurred during the main user processing loop: $($_.Exception.Message)"
    Write-Error "Processed $processedUsersCount users before this critical error."
}

# --- Summary ---
Write-Host "`n--- Operation Summary ---"
Write-Host "Total users found: $totalUsersToProcess"
Write-Host "Users for whom processing was attempted: $processedUsersCount"
Write-Host "Birthdays successfully cleared: $birthdaysClearedCount" -ForegroundColor Green
Write-Host "Users whose birthdays were already clear or skipped update: $alreadyClearOrSkippedCount"
Write-Host "Failed birthday reads (user skipped for update): $failedReadCount" -ForegroundColor Yellow
Write-Host "Failed birthday clear operations (after successful read): $failedUpdateCount" -ForegroundColor Red

if ($failedUsersList.Count -gt 0) {
    Write-Warning "`nDetails of failures (reads or updates):"
    $logTimestamp = Get-Date -Format "yyyyMMddHHmmss"
    $failureLogPath = ".\BirthdayClear_Failures_$logTimestamp.txt"
    foreach ($failedUserEntry in $failedUsersList) {
        Write-Warning $failedUserEntry
        Add-Content -Path $failureLogPath -Value $failedUserEntry
    }
    Write-Warning "Failure details also logged to: $failureLogPath"
}

# Optional: Disconnect from Microsoft Graph
# Write-Host "Disconnecting from Microsoft Graph..."
# Disconnect-MgGraph

Write-Host "`nScript finished."