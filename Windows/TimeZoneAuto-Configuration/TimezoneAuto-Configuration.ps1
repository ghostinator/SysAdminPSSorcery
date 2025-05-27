# Timezone Auto-Configuration Setup Script (v1.6.0 - XML Event Trigger)
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

Write-Host "=== Timezone Auto-Configuration Setup (v1.6.0) ===" -ForegroundColor Cyan
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
        Write-Error ("  Failed to create directory {0} : {1}" -f $scriptFolder, $_.Exception.Message)
        exit 1
    }
}
else {
    Write-Host "  Directory already exists: $scriptFolder" -ForegroundColor Green
}

# Step 2: Create the timezone detection script (UpdateTimezone.ps1)
Write-Host "`nStep 2: Creating timezone detection script ($($scriptPath))..." -ForegroundColor Yellow

$timezoneScriptContent = @'
# Automatic Timezone Detection and Configuration Script (v1.6.0 - XML Event Trigger Setup)
# This script detects location based on public IP and sets appropriate timezone.
# Disables Windows automatic timezone features to prevent conflicts.
# Defaults to Eastern Time if geolocation fails or IANA mapping is unsuccessful.

# Configuration for this script
$ScriptVersion = "1.6.0-ET-Fallback"
$LogFile = "{0}" # Placeholder for $logFileForInnerScript
$RegistryPath = "{1}" # Placeholder for $registryKeyForInnerScript

# Function to robustly disable Windows automatic timezone features
function Disable-WindowsAutomaticTimezone {
    Write-Host "Attempting to disable/control Windows automatic timezone features..."
    $ErrorActionPreference = 'SilentlyContinue' 

    try {
        Set-Service -Name tzautoupdate -StartupType Disabled -ErrorAction Stop
        Stop-Service -Name tzautoupdate -Force -ErrorAction SilentlyContinue 
        Write-Host "  ✓ Windows Time Zone Auto Update service (tzautoupdate) set to Disabled." -ForegroundColor Green
    }
    catch {
        Write-Warning ("  ⚠ Could not set tzautoupdate service startup type or stop it: {0}" -f $_.Exception.Message)
    }

    $timeSettingsPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers"
    if (Test-Path $timeSettingsPath) {
        try {
            Set-ItemProperty -Path $timeSettingsPath -Name "Enabled" -Value 0 -Type DWord -Force -ErrorAction Stop
            Write-Host "  ✓ Set HKLM DateTime\Servers 'Enabled' to 0 (attempt to influence default)." -ForegroundColor Green
        }
        catch {
            Write-Warning ("  ⚠ Could not set HKLM DateTime\Servers 'Enabled': {0}" -f $_.Exception.Message)
        }
    }

    $locationCapabilityPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
    if (Test-Path $locationCapabilityPath) {
        try {
            Set-ItemProperty -Path $locationCapabilityPath -Name "Value" -Value "Deny" -ErrorAction Stop
            Write-Host "  ✓ Set CapabilityAccessManager\ConsentStore\location to Deny." -ForegroundColor Green
        }
        catch {
            Write-Warning ("  ⚠ Could not set CapabilityAccessManager\ConsentStore\location Value: {0}" -f $_.Exception.Message)
        }
    }

    $tzInfoPath = "HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation"
    try {
        $currentDynamicDst = Get-ItemProperty -Path $tzInfoPath -Name "DynamicDaylightTimeDisabled" -ErrorAction SilentlyContinue
        if ($currentDynamicDst -and $currentDynamicDst.DynamicDaylightTimeDisabled -ne 0) {
            Set-ItemProperty -Path $tzInfoPath -Name "DynamicDaylightTimeDisabled" -Value 0 -Type DWord -Force -ErrorAction Stop
            Write-Host "  ✓ Ensured DynamicDaylightTimeDisabled is 0 (DST enabled per zone rules)." -ForegroundColor Green
        } elseif (-not $currentDynamicDst) {
            Set-ItemProperty -Path $tzInfoPath -Name "DynamicDaylightTimeDisabled" -Value 0 -Type DWord -Force -ErrorAction Stop
            Write-Host "  ✓ Set DynamicDaylightTimeDisabled to 0 (DST enabled per zone rules)." -ForegroundColor Green
        } else {
            Write-Host "  ✓ DynamicDaylightTimeDisabled is already 0 (DST enabled per zone rules)." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning ("  ⚠ Could not configure DynamicDaylightTimeDisabled: {0}" -f $_.Exception.Message)
    }

    $ErrorActionPreference = 'Continue' 
    Write-Host "  Finished attempt to disable/control Windows automatic timezone features."
}

# Function to get public IP address
function Get-PublicIPAddress {
    $uris = @(
        "https://api.ipify.org/",
        "https://ipinfo.io/ip",
        "https://icanhazip.com/",
        "https://checkip.amazonaws.com/"
    )
    foreach ($uri in $uris) {
        try {
            Write-Host "Attempting to get public IP from $uri..."
            $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -TimeoutSec 7 -ErrorAction Stop
            $publicIP = $response.Content.Trim()
            if ($publicIP -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                Write-Host ("  Public IP Address: {0} (from {1})" -f $publicIP, $uri) -ForegroundColor Green
                return $publicIP
            }
            else {
                Write-Warning ("  Invalid IP format from {0}: '{1}'" -f $uri, $publicIP)
            }
        }
        catch {
            Write-Warning ("  Failed to retrieve public IP address from {0}: {1}" -f $uri, $_.Exception.Message)
        }
    }
    Write-Error "Failed to retrieve public IP address from all configured sources."
    return $null
}

# Function to get geolocation data including timezone
function Get-GeoLocationData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$IPAddress
    )
    try {
        $apiUrl = "http://ip-api.com/json/$IPAddress"
        Write-Host "Querying geolocation API: $apiUrl"
        $geoData = Invoke-RestMethod -Uri $apiUrl -Method Get -TimeoutSec 10 -ErrorAction Stop
        
        if ($geoData.status -eq "success") {
            Write-Host "  Location Details (from ip-api.com):" -ForegroundColor Yellow
            Write-Host ("    City: {0}, Region: {1}, Country: {2}" -f $geoData.city, $geoData.regionName, $geoData.country)
            Write-Host ("    IANA Timezone: {0}" -f $geoData.timezone) -ForegroundColor Green
            return $geoData
        }
        else {
            Write-Error ("  Geolocation lookup failed (ip-api.com status: {0}, message: {1})" -f $geoData.status, $geoData.message)
            return $null
        }
    }
    catch {
        Write-Error ("  Exception during geolocation data retrieval: {0}" -f $_.Exception.Message)
        return $null
    }
}

# Function to convert IANA timezone to Windows timezone
function Convert-IANAToWindowsTimeZone {
    param (
        [Parameter(Mandatory=$true)]
        [string]$IANATimeZone
    )
    
    $timezoneMapping = @{
        # North America
        "America/New_York" = "Eastern Standard Time"; "America/Detroit" = "Eastern Standard Time";
        "America/Kentucky/Louisville" = "Eastern Standard Time"; "America/Kentucky/Monticello" = "Eastern Standard Time";
        "America/Indiana/Indianapolis" = "US Eastern Standard Time"; "America/Indiana/Vincennes" = "US Eastern Standard Time";
        "America/Indiana/Winamac" = "US Eastern Standard Time"; "America/Indiana/Marengo" = "US Eastern Standard Time";
        "America/Indiana/Petersburg" = "US Eastern Standard Time"; "America/Indiana/Vevay" = "US Eastern Standard Time";
        "America/Toronto" = "Eastern Standard Time";

        "America/Chicago" = "Central Standard Time"; "America/Winnipeg" = "Central Standard Time";
        "America/Indiana/Tell_City" = "Central Standard Time"; "America/Indiana/Knox" = "Central Standard Time";
        "America/Menominee" = "Central Standard Time";

        "America/Denver" = "Mountain Standard Time"; "America/Edmonton" = "Mountain Standard Time";
        "America/Boise" = "Mountain Standard Time";

        "America/Phoenix" = "US Mountain Standard Time";

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
        
        "Europe/Warsaw" = "Central European Standard Time";

        "Europe/Helsinki" = "FLE Standard Time"; "Europe/Kiev" = "FLE Standard Time";
        "Europe/Riga" = "FLE Standard Time"; "Europe/Sofia" = "FLE Standard Time";
        "Europe/Tallinn" = "FLE Standard Time"; "Europe/Vilnius" = "FLE Standard Time";
        "Europe/Athens" = "GTB Standard Time"; "Europe/Bucharest" = "GTB Standard Time";

        "Europe/Moscow" = "Russian Standard Time"; 
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
    }
    
    if ($IANATimeZone -and $timezoneMapping.ContainsKey($IANATimeZone)) {
        Write-Host ("  Found explicit mapping for IANA '{0}': '{1}'" -f $IANATimeZone, $timezoneMapping[$IANATimeZone]) -ForegroundColor Green
        return $timezoneMapping[$IANATimeZone]
    }
    elseif ($IANATimeZone) {
        Write-Warning ("  No explicit mapping found for IANA timezone: '{0}'. Attempting approximate match..." -f $IANATimeZone)
        $ianaCityOrRegion = $IANATimeZone.Split('/')[-1].Replace("_", " ")
        $availableTimeZones = Get-TimeZone -ListAvailable
        
        $matchingZone = $availableTimeZones | Where-Object {
            $_.StandardName -replace '\s\(.*\)', '' -eq $ianaCityOrRegion -or 
            $_.Id -replace '\s\(.*\)', '' -eq $ianaCityOrRegion -or          
            $_.StandardName -like "*$ianaCityOrRegion*" -or
            $_.Id -like "*$ianaCityOrRegion*"
        } | Select-Object -First 1
        
        if ($matchingZone) {
            Write-Host ("  Found approximate Windows match for '{0}': '{1}'" -f $IANATimeZone, $matchingZone.Id) -ForegroundColor Yellow
            return $matchingZone.Id
        }
    }
    
    Write-Warning ("  Could not map IANA '{0}' (or IANA was null/empty). Defaulting to 'Eastern Standard Time' as per script requirement." -f $IANATimeZone)
    return "Eastern Standard Time"
}

# Function to set the system timezone
function Set-SystemTimeZone {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WindowsTimeZoneId
    )
    try {
        $currentTimeZone = Get-TimeZone
        if ($currentTimeZone.Id -eq $WindowsTimeZoneId) {
            Write-Host "  System is already set to the target timezone: $WindowsTimeZoneId" -ForegroundColor Green
            return $true
        }
        
        Write-Host ("  Attempting to set timezone from '{0}' to '{1}'..." -f $currentTimeZone.Id, $WindowsTimeZoneId)
        Set-TimeZone -Id $WindowsTimeZoneId -PassThru -ErrorAction Stop
        Start-Sleep -Seconds 1 
        
        $newTimeZone = Get-TimeZone
        if ($newTimeZone.Id -eq $WindowsTimeZoneId) {
            Write-Host ("  Successfully set timezone to '{0}'" -f $newTimeZone.Id) -ForegroundColor Green
            Write-Host ("  Current local time: {0}" -f (Get-Date)) -ForegroundColor Cyan
            return $true
        }
        else {
            Write-Error ("  Failed to set timezone using Set-TimeZone. Current timezone is: '{0}', attempted: '{1}'." -f $newTimeZone.Id, $WindowsTimeZoneId)
            Write-Host ("  Attempting to set timezone with tzutil.exe /s `"{0}`"..." -f $WindowsTimeZoneId)
            tzutil.exe /s "$WindowsTimeZoneId"
            Start-Sleep -Seconds 1
            $newTimeZoneViaTzUtil = Get-TimeZone
            if ($newTimeZoneViaTzUtil.Id -eq $WindowsTimeZoneId) {
                Write-Host ("  Successfully set timezone to '{0}' using tzutil." -f $newTimeZoneViaTzUtil.Id) -ForegroundColor Green
                Write-Host ("  Current local time: {0}" -f (Get-Date)) -ForegroundColor Cyan
                return $true
            } else {
                Write-Error ("  Failed to set timezone with tzutil as well. Current is {0}" -f $newTimeZoneViaTzUtil.Id)
                return $false
            }
        }
    }
    catch {
        Write-Error ("  Exception while setting timezone to '{0}': {1}" -f $WindowsTimeZoneId, $_.Exception.Message)
        Write-Host ("  Attempting to set timezone with tzutil.exe /s `"{0}`" due to exception..." -f $WindowsTimeZoneId)
        tzutil.exe /s "$WindowsTimeZoneId"
        Start-Sleep -Seconds 1
        $newTimeZoneOnException = Get-TimeZone
        if ($newTimeZoneOnException.Id -eq $WindowsTimeZoneId) {
            Write-Host ("  Successfully set timezone to '{0}' using tzutil after exception." -f $newTimeZoneOnException.Id) -ForegroundColor Green
            return $true
        } else {
             Write-Error ("  Also failed to set timezone with tzutil after exception. Current is {0}" -f $newTimeZoneOnException.Id)
            return $false
        }
    }
}

# Function to update registry tracking
function Update-RegistryTracking {
    param (
        [string]$DetectedIANATz,
        [string]$SetWindowsTz,
        [string]$GeoInfo,
        [string]$UpdateStatus
    )
    try {
        if (!(Test-Path $RegistryPath)) {
            New-Item -Path $RegistryPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-ItemProperty -Path $RegistryPath -Name "LastIANATimezoneDetected" -Value $DetectedIANATz -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastWindowsTimezoneSet" -Value $SetWindowsTz -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastGeolocationInfo" -Value $GeoInfo -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastUpdateStatus" -Value $UpdateStatus -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "LastUpdateTime" -Value (Get-Date -Format 'u') -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $RegistryPath -Name "ScriptVersionRun" -Value $ScriptVersion -Force -ErrorAction SilentlyContinue
        Write-Host "  Updated registry tracking information."
    }
    catch {
        Write-Warning ("  Failed to update registry tracking: {0}" -f $_.Exception.Message)
    }
}

# Main script execution
function Main {
    try {
        Start-Transcript -Path $LogFile -Append -Force -ErrorAction Stop
    }
    catch {
        Write-Warning ("Could not start transcript logging to {0}. {1}" -f $LogFile, $_.Exception.Message)
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logPrefix = "[$timestamp] ($ScriptVersion)"
    
    Write-Output "$logPrefix Script execution started."
    Write-Host "`n=== Automatic Timezone Configuration Script ($ScriptVersion) ===" -ForegroundColor Cyan
    
    Disable-WindowsAutomaticTimezone
    
    $publicIP = $null
    $geoData = $null
    $ianaTimezoneDetected = "N/A"
    $windowsTimeZoneToSet = "Eastern Standard Time" 
    $locationInfo = "Initialization - Defaulting to Eastern Time"
    $finalStatusMessage = ""

    $publicIP = Get-PublicIPAddress
    
    if ($publicIP) {
        $locationInfo = "Public IP: $publicIP"
        $geoData = Get-GeoLocationData -IPAddress $publicIP
        
        if ($geoData -and $geoData.timezone) {
            $ianaTimezoneDetected = $geoData.timezone
            $locationInfo = "IP: $publicIP, City: $($geoData.city), Region: $($geoData.regionName), Country: $($geoData.country), IANA TZ: $ianaTimezoneDetected"
            Write-Host "  Successfully retrieved geolocation: $locationInfo" -ForegroundColor Green
            $windowsTimeZoneToSet = Convert-IANAToWindowsTimeZone -IANATimeZone $ianaTimezoneDetected
        }
        else {
            $finalStatusMessage = "Failed to get valid geolocation data or IANA timezone for IP $publicIP. Using default '$windowsTimeZoneToSet'."
            Write-Warning "  $finalStatusMessage"
            $ianaTimezoneDetected = "Unknown (GeoData failed or TZ missing)"
            $locationInfo = "IP: $publicIP (GeoData/IANA missing) - Defaulting to ET"
        }
    }
    else {
        $finalStatusMessage = "Failed to retrieve public IP address. Using default '$windowsTimeZoneToSet'."
        Write-Warning "  $finalStatusMessage"
        $ianaTimezoneDetected = "Unknown (No Public IP)"
        $locationInfo = "No Public IP - Defaulting to ET"
    }

    Write-Output "$logPrefix Determined target Windows timezone: $windowsTimeZoneToSet (IANA: $ianaTimezoneDetected, Location: $locationInfo)"
    Write-Host "`nAttempting to set system timezone to '$windowsTimeZoneToSet'..." -ForegroundColor Yellow
    
    if (Set-SystemTimeZone -WindowsTimeZone $windowsTimeZoneToSet) {
        $finalStatusMessage = "Successfully set timezone to '$windowsTimeZoneToSet'. ($locationInfo)"
        Write-Host $finalStatusMessage -ForegroundColor Green
    }
    else {
        $finalStatusMessage = "Failed to set timezone to '$windowsTimeZoneToSet'. ($locationInfo)"
        Write-Error $finalStatusMessage
    }
    
    Update-RegistryTracking -DetectedIANATz $ianaTimezoneDetected -SetWindowsTz $windowsTimeZoneToSet -GeoInfo $locationInfo -UpdateStatus $finalStatusMessage
    
    Write-Output "$logPrefix Script execution finished. Status: $finalStatusMessage"
    Write-Host "`n=== Timezone Configuration Attempt Complete ($ScriptVersion) ===" -ForegroundColor Cyan
    
    Stop-Transcript
}

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
    Write-Error ("  Failed to create script file '{0}': {1}" -f $scriptPath, $_.Exception.Message)
    exit 1
}

# Step 3: Create or update the scheduled task (delete if exists, then create)
Write-Host "`nStep 3: Creating/Updating scheduled task '$taskName' using XML definition..." -ForegroundColor Yellow

try {
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "  Found existing task '$taskName'. Unregistering..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
        Write-Host "  Existing task '$taskName' unregistered." -ForegroundColor Green
    }

    $powershellExecutable = "PowerShell.exe"
    $taskArguments = "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$($scriptPath)`""

    $TaskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Automatically updates timezone based on IP geolocation (v1.6.3 - XML Principal Fix). Runs as SYSTEM.</Description>
    <Author>PowerShell Script</Author>
    <URI>\$($taskName)</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
    <BootTrigger>
      <Enabled>true</Enabled>
    </BootTrigger>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription><![CDATA[
        <QueryList>
          <Query Id="0" Path="Microsoft-Windows-NetworkProfile/Operational">
            <Select Path="Microsoft-Windows-NetworkProfile/Operational">*[System[Provider[@Name='NetworkProfile'] and EventID=10000]]</Select>
          </Query>
        </QueryList>
      ]]></Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId> <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>4</Priority> 
    <RestartOnFailure>
        <Interval>PT5M</Interval>
        <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$($powershellExecutable)</Command>
      <Arguments>$($taskArguments)</Arguments>
    </Exec>
  </Actions>
</Task>
"@

    Register-ScheduledTask -TaskName $taskName -TaskPath "\" -Xml $TaskXML -User "NT AUTHORITY\SYSTEM" -Force -ErrorAction Stop | Out-Null
    
    Write-Host "  Successfully created/updated scheduled task: '$taskName' using XML definition." -ForegroundColor Green
    Write-Host "  Task will run as SYSTEM at Startup, User Logon, and on NetworkProfile Event 10000." -ForegroundColor Green
}
catch {
    Write-Error ("  Failed to create/update scheduled task '{0}' using XML: {1}" -f $taskName, $_.Exception.Message)
    Write-Warning "  Full XML used for task definition (review for issues):"
    Write-Warning $TaskXML 
}

# Step 4: Verify the setup
Write-Host "`nStep 4: Verifying setup..." -ForegroundColor Yellow
if (Test-Path $scriptPath) {
    Write-Host "  ✓ Script file found: $scriptPath" -ForegroundColor Green
} 
else {
    Write-Host "  ✗ Script file NOT found: $scriptPath" -ForegroundColor Red
}

$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($task) {
    Write-Host "  ✓ Scheduled task '$($task.TaskName)' confirmed." -ForegroundColor Green
    Write-Host "    State: $($task.State)" -ForegroundColor Gray
    if ($task.Triggers) {
        $triggerTypes = $task.Triggers | ForEach-Object { $_.TriggerType }
        Write-Host "    Triggers: $($triggerTypes -join ', ')" -ForegroundColor Gray
    } else {
        Write-Host "    Triggers: None found or unable to read." -ForegroundColor Yellow
    }
} 
else {
    Write-Host "  ✗ Scheduled task '$taskName' NOT found or may have failed to create properly." -ForegroundColor Red
}

# Step 5: Run the scheduled task immediately
Write-Host "`nStep 5: Attempting to run the scheduled task '$taskName' immediately..." -ForegroundColor Yellow
$taskToRun = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue # Re-fetch to be sure
if ($taskToRun) {
    try {
        Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
        Write-Host "  Successfully initiated an immediate run of task '$taskName'." -ForegroundColor Green
        Write-Host "  Check the log file for execution details: $logFileForInnerScript"
    }
    catch {
        Write-Error ("  Failed to start scheduled task '{0}': {1}" -f $taskName, $_.Exception.Message)
        Write-Warning "  You might need to run it manually from Task Scheduler or wait for the next trigger."
    }
} else {
    Write-Warning "  Cannot run task '$taskName' as it was not found (or creation failed)."
}

# Step 6: Informational - Checking Windows automatic timezone service (tzautoupdate) status
Write-Host "`nStep 6: Informational - Checking Windows automatic timezone service (tzautoupdate) status..." -ForegroundColor Yellow
try {
    $tzService = Get-Service -Name "tzautoupdate" -ErrorAction SilentlyContinue
    if ($tzService) {
        Write-Host ("  tzautoupdate service status: {0}, StartType: {1}" -f $tzService.Status, $tzService.StartType) -ForegroundColor Gray
    } else {
        Write-Host "  tzautoupdate service not found (this can be normal if it was removed or on some OS versions)." -ForegroundColor Gray
    }
} catch {
    Write-Warning ("  Could not verify tzautoupdate service status: {0}" -f $_.Exception.Message)
}

Write-Host "`n=== Setup Script Finished (v1.6.0) ===" -ForegroundColor Cyan
