# Timezone Auto-Configuration Setup Script (v1.3 - ET Fallback, Task Overwrite & Immediate Run)
# This script creates the timezone detection script and scheduled task.
# It ensures Windows automatic timezone features are disabled to prevent conflicts.
# It defaults to Eastern Time if geolocation fails or IANA mapping is unsuccessful.
# It deletes any existing scheduled task with the same name before creation.
# It runs the newly created scheduled task immediately after setup.

# Configuration
$scriptFolder = "C:\Scripts"
$scriptPath = Join-Path -Path $scriptFolder -ChildPath "UpdateTimezone.ps1"
$taskName = "AutoTimezoneUpdate"
$logFileForInnerScript = Join-Path -Path $scriptFolder -ChildPath "TimezoneUpdate.log"
$registryKeyForInnerScript = "HKLM:\SOFTWARE\AutoTimezone"

Write-Host "=== Timezone Auto-Configuration Setup (v1.3) ===" -ForegroundColor Cyan
Write-Host "Setting up automatic timezone detection and configuration..." -ForegroundColor Yellow
Write-Host "  - Script will be placed in: $scriptFolder"
Write-Host "  - Scheduled Task Name: $taskName"
Write-Host "  - Default Timezone on failure: Eastern Standard Time"
Write-Host "  - Log for detection script: $logFileForInnerScript"
Write-Host "  - Registry tracking: $registryKeyForInnerScript"

# Step 1: Create the Scripts folder if it doesn't exist
Write-Host "`nStep 1: Ensuring script directory exists..." -ForegroundColor Yellow
if (!(Test-Path $scriptFolder)) {
    try {
        New-Item -Path $scriptFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "  Created directory: $scriptFolder" -ForegroundColor Green
    }
    catch {
        Write-Error "  Failed to create directory $scriptFolder: $($_.Exception.Message)"
        exit 1
    }
}
else {
    Write-Host "  Directory already exists: $scriptFolder" -ForegroundColor Green
}

# Step 2: Create the timezone detection script (UpdateTimezone.ps1)
Write-Host "`nStep 2: Creating timezone detection script ($($scriptPath))..." -ForegroundColor Yellow

$timezoneScriptContent = @'
# Automatic Timezone Detection and Configuration Script (v1.3 - ET Fallback Enhanced)
# This script detects location based on public IP and sets appropriate timezone.
# Disables Windows automatic timezone features to prevent conflicts.
# Defaults to Eastern Time if geolocation fails or IANA mapping is unsuccessful.

# Configuration for this script
$ScriptVersion = "1.3-ET-Fallback"
$LogFile = "{0}" # Placeholder for $logFileForInnerScript
$RegistryPath = "{1}" # Placeholder for $registryKeyForInnerScript

# Function to robustly disable Windows automatic timezone features
function Disable-WindowsAutomaticTimezone {{
    Write-Host "Attempting to disable/control Windows automatic timezone features..."
    $ErrorActionPreference = 'SilentlyContinue' # Suppress errors for individual operations but log overall

    # Disable Time Zone Auto Update service
    try {{
        Set-Service -Name tzautoupdate -StartupType Disabled -ErrorAction Stop
        Stop-Service -Name tzautoupdate -Force -ErrorAction SilentlyContinue # Stop if running
        Write-Host "  ✓ Windows Time Zone Auto Update service (tzautoupdate) set to Disabled." -ForegroundColor Green
    }}
    catch {{
        Write-Warning "  ⚠ Could not set tzautoupdate service startup type or stop it: $($_.Exception.Message)"
    }}

    # Disable "Set time zone automatically" via known registry setting (Windows 10/11 UI)
    # This is a user-specific setting, but we try to set a system-wide default policy if possible
    # For SYSTEM context, this specific key might not be the primary control but worth attempting.
    $timeSettingsPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers"
    if (Test-Path $timeSettingsPath) {{
        try {{
            Set-ItemProperty -Path $timeSettingsPath -Name "Enabled" -Value 0 -Type DWord -Force -ErrorAction Stop
            Write-Host "  ✓ Set HKLM DateTime\Servers 'Enabled' to 0 (attempt to influence default)." -ForegroundColor Green
        }}
        catch {{
            Write-Warning "  ⚠ Could not set HKLM DateTime\Servers 'Enabled': $($_.Exception.Message)"
        }}
    }}

    # Deny location access for general location capability (less direct for system timezone but part of hardening)
    $locationCapabilityPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
    if (Test-Path $locationCapabilityPath) {{
        try {{
            Set-ItemProperty -Path $locationCapabilityPath -Name "Value" -Value "Deny" -ErrorAction Stop
            Write-Host "  ✓ Set CapabilityAccessManager\ConsentStore\location to Deny." -ForegroundColor Green
        }}
        catch {{
            Write-Warning "  ⚠ Could not set CapabilityAccessManager\ConsentStore\location Value: $($_.Exception.Message)"
        }}
    }}

    # Ensure DynamicDaylightTimeDisabled is NOT set to 1 (which would disable DST)
    # We want DST to work according to the chosen timezone's rules.
    $tzInfoPath = "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation"
    try {{
        $currentDynamicDst = Get-ItemProperty -Path $tzInfoPath -Name "DynamicDaylightTimeDisabled" -ErrorAction SilentlyContinue
        if ($currentDynamicDst -and $currentDynamicDst.DynamicDaylightTimeDisabled -ne 0) {{
            Set-ItemProperty -Path $tzInfoPath -Name "DynamicDaylightTimeDisabled" -Value 0 -Type DWord -Force -ErrorAction Stop
            Write-Host "  ✓ Ensured DynamicDaylightTimeDisabled is 0 (DST enabled per zone rules)." -ForegroundColor Green
        }} elseif (-not $currentDynamicDst) {{
            Set-ItemProperty -Path $tzInfoPath -Name "DynamicDaylightTimeDisabled" -Value 0 -Type DWord -Force -ErrorAction Stop # Create if not exists
            Write-Host "  ✓ Set DynamicDaylightTimeDisabled to 0 (DST enabled per zone rules)." -ForegroundColor Green
        }} else {{
            Write-Host "  ✓ DynamicDaylightTimeDisabled is already 0 (DST enabled per zone rules)." -ForegroundColor Green
        }}
    }}
    catch {{
        Write-Warning "  ⚠ Could not configure DynamicDaylightTimeDisabled: $($_.Exception.Message)"
    }}

    $ErrorActionPreference = 'Continue' # Restore default
    Write-Host "  Finished attempt to disable/control Windows automatic timezone features."
}}

# Function to get public IP address
function Get-PublicIPAddress {{
    $uris = @(
        "https://api.ipify.org/",
        "https://ipinfo.io/ip",
        "https://icanhazip.com/",
        "https://checkip.amazonaws.com/"
    )
    foreach ($uri in $uris) {{
        try {{
            Write-Host "Attempting to get public IP from $uri..."
            $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -TimeoutSec 7 -ErrorAction Stop
            $publicIP = $response.Content.Trim()
            if ($publicIP -match '^\d{{1,3}}\.\d{{1,3}}\.\d{{1,3}}\.\d{{1,3}}$') {{
                Write-Host "  Public IP Address: $publicIP (from $uri)" -ForegroundColor Green
                return $publicIP
            }}
            else {{
                Write-Warning "  Invalid IP format from $uri: '$publicIP'"
            }}
        }}
        catch {{
            Write-Warning "  Failed to retrieve public IP address from $uri: $($_.Exception.Message)"
        }}
    }}
    Write-Error "Failed to retrieve public IP address from all configured sources."
    return $null
}}

# Function to get geolocation data including timezone
function Get-GeoLocationData {{
    param (
        [Parameter(Mandatory=$true)]
        [string]$IPAddress
    )
    try {{
        $apiUrl = "http://ip-api.com/json/$IPAddress"
        Write-Host "Querying geolocation API: $apiUrl"
        $geoData = Invoke-RestMethod -Uri $apiUrl -Method Get -TimeoutSec 10 -ErrorAction Stop
        
        if ($geoData.status -eq "success") {{
            Write-Host "  Location Details (from ip-api.com):" -ForegroundColor Yellow
            Write-Host "    City: $($geoData.city), Region: $($geoData.regionName), Country: $($geoData.country)"
            Write-Host "    IANA Timezone: $($geoData.timezone)" -ForegroundColor Green
            return $geoData
        }}
        else {{
            Write-Error "  Geolocation lookup failed (ip-api.com status: $($geoData.status), message: $($geoData.message))"
            return $null
        }}
    }}
    catch {{
        Write-Error "  Exception during geolocation data retrieval: $($_.Exception.Message)"
        return $null
    }}
}}

# Function to convert IANA timezone to Windows timezone
function Convert-IANAToWindowsTimeZone {{
    param (
        [Parameter(Mandatory=$true)]
        [string]$IANATimeZone
    )
    
    $timezoneMapping = @{{
        # North America
        "America/New_York" = "Eastern Standard Time"; "America/Detroit" = "Eastern Standard Time";
        "America/Kentucky/Louisville" = "Eastern Standard Time"; "America/Kentucky/Monticello" = "Eastern Standard Time";
        "America/Indiana/Indianapolis" = "US Eastern Standard Time"; "America/Indiana/Vincennes" = "US Eastern Standard Time";
        "America/Indiana/Winamac" = "US Eastern Standard Time"; "America/Indiana/Marengo" = "US Eastern Standard Time";
        "America/Indiana/Petersburg" = "US Eastern Standard Time"; "America/Indiana/Vevay" = "US Eastern Standard Time";
        "America/Toronto" = "Eastern Standard Time";

        "America/Chicago" = "Central Standard Time"; "America/Winnipeg" = "Central Standard Time";
        "America/Indiana/Tell_City" = "Central Standard Time"; "America/Indiana/Knox" = "Central Standard Time"; # Knox, IN is CT
        "America/Menominee" = "Central Standard Time";

        "America/Denver" = "Mountain Standard Time"; "America/Edmonton" = "Mountain Standard Time";
        "America/Boise" = "Mountain Standard Time";

        "America/Phoenix" = "US Mountain Standard Time"; # Arizona does not observe DST

        "America/Los_Angeles" = "Pacific Standard Time"; "America/Vancouver" = "Pacific Standard Time";
        "America/Tijuana" = "Pacific Standard Time";

        "America/Anchorage" = "Alaskan Standard Time"; "America/Juneau" = "Alaskan Standard Time";
        "America/Nome" = "Alaskan Standard Time"; "America/Yakutat" = "Alaskan Standard Time";

        "America/Halifax" = "Atlantic Standard Time"; "America/Glace_Bay" = "Atlantic Standard Time";
        "America/Moncton" = "Atlantic Standard Time"; "America/Goose_Bay" = "Atlantic Standard Time";
        "Atlantic/Bermuda" = "Atlantic Standard Time";

        "America/St_Johns" = "Newfoundland Standard Time";
        "Pacific/Honolulu" = "Hawaiian Standard Time";

        # Europe
        "Europe/London" = "GMT Standard Time"; "Europe/Dublin" = "GMT Standard Time";
        "Europe/Lisbon" = "GMT Standard Time"; "Atlantic/Canary" = "GMT Standard Time";

        "Europe/Paris" = "Romance Standard Time"; "Europe/Brussels" = "Romance Standard Time";
        "Europe/Copenhagen" = "Romance Standard Time"; "Europe/Madrid" = "Romance Standard Time";

        "Europe/Berlin" = "W. Europe Standard Time"; "Europe/Amsterdam" = "W. Europe Standard Time";
        "Europe/Rome" = "W. Europe Standard Time"; "Europe/Stockholm" = "W. Europe Standard Time";
        "Europe/Vienna" = "W. Europe Standard Time"; "Europe/Oslo" = "W. Europe Standard Time";
        "Europe/Zurich" = "W. Europe Standard Time";

        "Europe/Prague" = "Central Europe Standard Time"; "Europe/Budapest" = "Central Europe Standard Time";
        "Europe/Belgrade" = "Central Europe Standard Time"; "Europe/Bratislava" = "Central Europe Standard Time";
        "Europe/Ljubljana" = "Central Europe Standard Time"; 
        
        "Europe/Warsaw" = "Central European Standard Time"; # Note slight difference in Windows name

        "Europe/Helsinki" = "FLE Standard Time"; "Europe/Kiev" = "FLE Standard Time";
        "Europe/Riga" = "FLE Standard Time"; "Europe/Sofia" = "FLE Standard Time";
        "Europe/Tallinn" = "FLE Standard Time"; "Europe/Vilnius" = "FLE Standard Time";
        "Europe/Athens" = "GTB Standard Time"; "Europe/Bucharest" = "GTB Standard Time";

        "Europe/Moscow" = "Russian Standard Time"; # This may vary, Russia has many zones
        "Europe/Istanbul" = "Turkey Standard Time";

        # Asia
        "Asia/Tokyo" = "Tokyo Standard Time";
        "Asia/Shanghai" = "China Standard Time"; "Asia/Hong_Kong" = "China Standard Time";
        "Asia/Singapore" = "Singapore Standard Time"; "Asia/Kuala_Lumpur" = "Singapore Standard Time";
        "Asia/Seoul" = "Korea Standard Time";
        "Asia/Kolkata" = "India Standard Time"; "Asia/Calcutta" = "India Standard Time";
        "Asia/Bangkok" = "SE Asia Standard Time"; "Asia/Jakarta" = "SE Asia Standard Time";
        "Asia/Dubai" = "Arabian Standard Time"; "Asia/Muscat" = "Arabian Standard Time";
        "Asia/Riyadh" = "Arab Standard Time"; "Asia/Kuwait" = "Arab Standard Time";
        "Asia/Tehran" = "Iran Standard Time";
        "Asia/Jerusalem" = "Israel Standard Time"; "Asia/Tel_Aviv" = "Israel Standard Time";

        # Australia / Pacific
        "Australia/Sydney" = "AUS Eastern Standard Time"; "Australia/Melbourne" = "AUS Eastern Standard Time";
        "Australia/Brisbane" = "E. Australia Standard Time";
        "Australia/Perth" = "W. Australia Standard Time";
        "Australia/Adelaide" = "Cen. Australia Standard Time";
        "Australia/Darwin" = "AUS Central Standard Time";
        "Pacific/Auckland" = "New Zealand Standard Time";
        "Pacific/Fiji" = "Fiji Standard Time";

        # Africa
        "Africa/Cairo" = "Egypt Standard Time";
        "Africa/Nairobi" = "E. Africa Standard Time";
        "Africa/Johannesburg" = "South Africa Standard Time";
        "Africa/Lagos" = "W. Central Africa Standard Time";
        "Africa/Casablanca" = "Morocco Standard Time";

        # UTC
        "UTC" = "UTC"; "Etc/GMT" = "UTC"; "GMT" = "UTC"
    }}
    
    if ($IANATimeZone -and $timezoneMapping.ContainsKey($IANATimeZone)) {{
        Write-Host "  Found explicit mapping for IANA '$IANATimeZone': '$($timezoneMapping[$IANATimeZone])'" -ForegroundColor Green
        return $timezoneMapping[$IANATimeZone]
    }}
    elseif ($IANATimeZone) {{
        Write-Warning "  No explicit mapping found for IANA timezone: '$IANATimeZone'. Attempting approximate match..."
        $ianaCityOrRegion = $IANATimeZone.Split('/')[-1].Replace("_", " ")
        $availableTimeZones = Get-TimeZone -ListAvailable
        
        $matchingZone = $availableTimeZones | Where-Object {{
            $_.StandardName -replace '\s\(.*\)', '' -eq $ianaCityOrRegion -or # Exact match on StandardName (ignoring parenthesized part)
            $_.Id -replace '\s\(.*\)', '' -eq $ianaCityOrRegion -or           # Exact match on ID (ignoring parenthesized part)
            $_.StandardName -like "*$ianaCityOrRegion*" -or
            $_.Id -like "*$ianaCityOrRegion*"
        }} | Select-Object -First 1
        
        if ($matchingZone) {{
            Write-Host "  Found approximate Windows match for '$IANATimeZone': '$($matchingZone.Id)'" -ForegroundColor Yellow
            return $matchingZone.Id
        }}
    }}
    
    # Fallback if IANA is null, empty, or no match is found
    Write-Warning "  Could not map IANA '$IANATimeZone' (or IANA was null/empty). Defaulting to 'Eastern Standard Time' as per script requirement."
    return "Eastern Standard Time"
}}

# Function to set the system timezone
function Set-SystemTimeZone {{
    param (
        [Parameter(Mandatory=$true)]
        [string]$WindowsTimeZoneId
    )
    try {{
        $currentTimeZone = Get-TimeZone
        if ($currentTimeZone.Id -eq $WindowsTimeZoneId) {{
            Write-Host "  System is already set to the target timezone: $WindowsTimeZoneId" -ForegroundColor Green
            return $true
        }}
        
        Write-Host "  Attempting to set timezone from '$($currentTimeZone.Id)' to '$WindowsTimeZoneId'..."
        Set-TimeZone -Id $WindowsTimeZoneId -PassThru -ErrorAction Stop
        Start-Sleep -Seconds 1 # Brief pause for system to process
        
        $newTimeZone = Get-TimeZone
        if ($newTimeZone.Id -eq $WindowsTimeZoneId) {{
            Write-Host "  Successfully set timezone to '$($newTimeZone.Id)'" -ForegroundColor Green
            Write-Host "  Current local time: $(Get-Date)" -ForegroundColor Cyan
            return $true
        }}
        else {{
            Write-Error "  Failed to set timezone. Current timezone is: '$($newTimeZone.Id)', attempted: '$WindowsTimeZoneId'."
            # Try with tzutil as a fallback for setting
            Write-Host "  Attempting to set timezone with tzutil.exe /s `"$WindowsTimeZoneId`"..."
            tzutil.exe /s "$WindowsTimeZoneId"
            Start-Sleep -Seconds 1
            $newTimeZoneViaTzUtil = Get-TimeZone
            if ($newTimeZoneViaTzUtil.Id -eq $WindowsTimeZoneId) {{
                Write-Host "  Successfully set timezone to '$($newTimeZoneViaTzUtil.Id)' using tzutil." -ForegroundColor Green
                Write-Host "  Current local time: $(Get-Date)" -ForegroundColor Cyan
                return $true
            }} else {{
                Write-Error "  Failed to set timezone with tzutil as well. Current is $($newTimeZoneViaTzUtil.Id)"
                return $false
            }}
        }}
    }}
    catch {{
        Write-Error "  Exception while setting timezone to '$WindowsTimeZoneId': $($_.Exception.Message)"
        # Last resort with tzutil on exception
        Write-Host "  Attempting to set timezone with tzutil.exe /s `"$WindowsTimeZoneId`" due to exception..."
        tzutil.exe /s "$WindowsTimeZoneId"
        Start-Sleep -Seconds 1
        $newTimeZoneOnException = Get-TimeZone
        if ($newTimeZoneOnException.Id -eq $WindowsTimeZoneId) {{
            Write-Host "  Successfully set timezone to '$($newTimeZoneOnException.Id)' using tzutil after exception." -ForegroundColor Green
            return $true
        }} else {{
             Write-Error "  Also failed to set timezone with tzutil after exception. Current is $($newTimeZoneOnException.Id)"
            return $false
        }}
    }}
}}

# Function to update registry tracking
function Update-RegistryTracking {{
    param (
        [string]$DetectedIANATz,
        [string]$SetWindowsTz,
        [string]$GeoInfo,
        [string]$UpdateStatus
    )
    try {{
        if (!(Test-Path $RegistryPath)) {{
            New-Item -Path $RegistryPath -Force -ErrorAction SilentlyContinue | Out-Null
        }}
        Set-ItemProperty -Path $RegistryPath -Name "LastIANATimezoneDetected" -Value $DetectedIANATz -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastWindowsTimezoneSet" -Value $SetWindowsTz -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastGeolocationInfo" -Value $GeoInfo -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastUpdateStatus" -Value $UpdateStatus -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastUpdateTime" -Value (Get-Date -Format 'u') -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "ScriptVersionRun" -Value $ScriptVersion -Force -ErrorAction SilentlyContinue
        Write-Host "  Updated registry tracking information."
    }}
    catch {{
        Write-Warning "  Failed to update registry tracking: $($_.Exception.Message)"
    }}
}}

# Main script execution
function Main {{
    # Start logging to file
    try {{
        Start-Transcript -Path $LogFile -Append -Force -ErrorAction Stop
    }}
    catch {{
        Write-Warning "Could not start transcript logging to $LogFile. $($_.Exception.Message)"
    }}

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logPrefix = "[$timestamp] ($ScriptVersion)"
    
    Write-Output "$logPrefix Script execution started."
    Write-Host "`n=== Automatic Timezone Configuration Script ($ScriptVersion) ===" -ForegroundColor Cyan
    
    Disable-WindowsAutomaticTimezone
    
    $publicIP = $null
    $geoData = $null
    $ianaTimezoneDetected = "N/A"
    $windowsTimezoneToSet = "Eastern Standard Time" # Overall script default
    $locationInfo = "Initialization - Defaulting to Eastern Time"
    $finalStatusMessage = ""

    $publicIP = Get-PublicIPAddress
    
    if ($publicIP) {{
        $locationInfo = "Public IP: $publicIP"
        $geoData = Get-GeoLocationData -IPAddress $publicIP
        
        if ($geoData -and $geoData.timezone) {{
            $ianaTimezoneDetected = $geoData.timezone
            $locationInfo = "IP: $publicIP, City: $($geoData.city), Region: $($geoData.regionName), Country: $($geoData.country), IANA TZ: $ianaTimezoneDetected"
            Write-Host "  Successfully retrieved geolocation: $locationInfo" -ForegroundColor Green
            $windowsTimezoneToSet = Convert-IANAToWindowsTimeZone -IANATimeZone $ianaTimezoneDetected
        }}
        else {{
            $finalStatusMessage = "Failed to get valid geolocation data or IANA timezone for IP $publicIP. Using default '$windowsTimezoneToSet'."
            Write-Warning "  $finalStatusMessage"
            $ianaTimezoneDetected = "Unknown (GeoData/IANA missing)"
            $locationInfo = "IP: $publicIP (GeoData/IANA missing) - Defaulting to ET"
        }}
    }}
    else {{
        $finalStatusMessage = "Failed to retrieve public IP address. Using default '$windowsTimezoneToSet'."
        Write-Warning "  $finalStatusMessage"
        $ianaTimezoneDetected = "Unknown (No Public IP)"
        $locationInfo = "No Public IP - Defaulting to ET"
    }}

    Write-Output "$logPrefix Determined target Windows timezone: $windowsTimezoneToSet (IANA: $ianaTimezoneDetected, Location: $locationInfo)"
    Write-Host "`nAttempting to set system timezone to '$windowsTimeZoneToSet'..." -ForegroundColor Yellow
    
    if (Set-SystemTimeZone -WindowsTimeZone $windowsTimeZoneToSet) {{
        $finalStatusMessage = "Successfully set timezone to '$windowsTimezoneToSet'. ($locationInfo)"
        Write-Host $finalStatusMessage -ForegroundColor Green
    }}
    else {{
        $finalStatusMessage = "Failed to set timezone to '$windowsTimezoneToSet'. ($locationInfo)"
        Write-Error $finalStatusMessage
    }}
    
    Update-RegistryTracking -DetectedIANATz $ianaTimezoneDetected -SetWindowsTz $windowsTimezoneToSet -GeoInfo $locationInfo -UpdateStatus $finalStatusMessage
    
    Write-Output "$logPrefix Script execution finished. Status: $finalStatusMessage"
    Write-Host "`n=== Timezone Configuration Attempt Complete ($ScriptVersion) ===" -ForegroundColor Cyan
    
    Stop-Transcript
}}

# Run the main function
Main
'@ # End of $timezoneScriptContent heredoc

# Inject dynamic paths into the script content
$timezoneScriptContent = $timezoneScriptContent -replace '\{0\}', [regex]::Escape($logFileForInnerScript)
$timezoneScriptContent = $timezoneScriptContent -replace '\{1\}', [regex]::Escape($registryKeyForInnerScript)

try {
    Set-Content -Path $scriptPath -Value $timezoneScriptContent -Encoding UTF8 -Force -ErrorAction Stop
    Write-Host "  Created/Updated timezone detection script: $scriptPath" -ForegroundColor Green
}
catch {
    Write-Error "  Failed to create script file '$scriptPath': $($_.Exception.Message)"
    exit 1
}

# Step 3: Create or update the scheduled task (delete if exists, then create)
Write-Host "`nStep 3: Creating/Updating scheduled task '$taskName'..." -ForegroundColor Yellow

try {
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "  Found existing task '$taskName'. Unregistering..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
        Write-Host "  Existing task '$taskName' unregistered." -ForegroundColor Green
    }

    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$scriptPath`"" -ErrorAction Stop
    
    $triggers = @()
    $triggers += New-ScheduledTaskTrigger -AtLogOn -ErrorAction Stop
    $triggers += New-ScheduledTaskTrigger -AtStartup -ErrorAction Stop
    
    # Trigger on network connection event (NetworkProfile Event ID 10000 - a network is connected and identified)
    # This is more reliable than generic "network available" for detecting actual network changes that might imply location change.
    try {
         $triggers += New-ScheduledTaskTrigger -EventIdentifier 10000 -Source "NetworkProfile" -LogName "Microsoft-Windows-NetworkProfile/Operational" -ErrorAction Stop
         Write-Host "  Added trigger for NetworkProfile Event ID 10000." -ForegroundColor DarkGray
    } catch {
        Write-Warning "  Could not create NetworkProfile event trigger (Source/Log may not exist or permissions issue): $($_.Exception.Message). Task will use Logon/Startup only."
    }
    
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -MultipleInstances IgnoreNew `
        -Priority 4 `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 5) `
        -RunOnlyIfNetworkAvailable:$false # Script handles network checks internally; task should always try to run
        # -WakeToRun # Consider if waking the computer is desired for this task

    $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest -ErrorAction Stop
    
    Register-ScheduledTask -TaskName $taskName -TaskPath "\" -Description "Automatically updates timezone based on IP geolocation (v1.3 - ET Fallback). Runs as SYSTEM." -Action $action -Settings $settings -Trigger $triggers -Principal $principal -Force -ErrorAction Stop | Out-Null
    
    Write-Host "  Successfully created/updated scheduled task: '$taskName'" -ForegroundColor Green
    Write-Host "  Task will run as SYSTEM at Startup, User Logon, and on specific Network Events (if trigger created)." -ForegroundColor Green
}
catch {
    Write-Error "  Failed to create/update scheduled task '$taskName': $($_.Exception.Message)"
    # Do not exit here, attempt to verify what was done and then run if task exists.
}

# Step 4: Verify the setup
Write-Host "`nStep 4: Verifying setup..." -ForegroundColor Yellow
if (Test-Path $scriptPath) {
    Write-Host "  ✓ Script file found: $scriptPath" -ForegroundColor Green
} else {
    Write-Host "  ✗ Script file NOT found: $scriptPath" -ForegroundColor Red
}

$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($task) {
    Write-Host "  ✓ Scheduled task '$($task.TaskName)' confirmed." -ForegroundColor Green
    Write-Host "    State: $($task.State)" -ForegroundColor Gray
    Write-Host "    Triggers: $($task.Triggers.TriggerType -join ', ')" -ForegroundColor Gray
} else {
    Write-Host "  ✗ Scheduled task '$taskName' NOT found or failed to create." -ForegroundColor Red
}

# Step 5: Run the scheduled task immediately
Write-Host "`nStep 5: Attempting to run the scheduled task '$taskName' immediately..." -ForegroundColor Yellow
if ($task) {
    try {
        Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
        Write-Host "  Successfully initiated an immediate run of task '$taskName'." -ForegroundColor Green
        Write-Host "  Check the log file for execution details: $logFileForInnerScript"
    }
    catch {
        Write-Error "  Failed to start scheduled task '$taskName': $($_.Exception.Message)"
        Write-Warning "  You might need to run it manually from Task Scheduler or wait for the next trigger."
    }
} else {
    Write-Warning "  Cannot run task '$taskName' as it was not found or failed to create."
}

# Check Windows automatic timezone service status (informational)
Write-Host "`nStep 6: Informational - Checking Windows automatic timezone service (tzautoupdate) status..." -ForegroundColor Yellow
try {
    $tzService = Get-Service -Name "tzautoupdate" -ErrorAction SilentlyContinue
    if ($tzService) {
        Write-Host "  tzautoupdate service status: $($tzService.Status), StartType: $($tzService.StartType)" -ForegroundColor Gray
    } else {
        Write-Host "  tzautoupdate service not found (this is normal if it was removed or on older OS)." -ForegroundColor Gray
    }
} catch {
    Write-Warning "  Could not verify tzautoupdate service status: $($_.Exception.Message)"
}


Write-Host "`n=== Setup Script Finished (v1.3) ===" -ForegroundColor Cyan