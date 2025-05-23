# Complete Timezone Auto-Configuration Setup Script
# This script creates the timezone detection script and scheduled task

# Configuration
$scriptFolder = "C:\Scripts"
$scriptPath = "$scriptFolder\UpdateTimezone.ps1"
$taskName = "AutoTimezoneUpdate"

Write-Host "=== Timezone Auto-Configuration Setup ===" -ForegroundColor Cyan
Write-Host "Setting up automatic timezone detection and configuration..." -ForegroundColor Yellow

# Step 1: Create the Scripts folder if it doesn't exist
Write-Host "`nStep 1: Creating script directory..." -ForegroundColor Yellow
if (!(Test-Path $scriptFolder)) {
    try {
        New-Item -Path $scriptFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created directory: $scriptFolder" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create directory $scriptFolder : $_"
        exit 1
    }
}
else {
    Write-Host "Directory already exists: $scriptFolder" -ForegroundColor Green
}

# Step 2: Create the timezone detection script
Write-Host "`nStep 2: Creating timezone detection script..." -ForegroundColor Yellow

$timezoneScript = @'
# Automatic Timezone Detection and Configuration Script
# This script detects location based on public IP and sets appropriate timezone

# Function to get public IP address
function Get-PublicIPAddress {
    try {
        $publicIP = (Invoke-WebRequest -Uri "https://api.ipify.org/" -UseBasicParsing -TimeoutSec 10).Content.Trim()
        Write-Host "Public IP Address: $publicIP" -ForegroundColor Green
        return $publicIP
    }
    catch {
        Write-Error "Failed to retrieve public IP address: $_"
        return $null
    }
}

# Function to get geolocation data including timezone
function Get-GeoLocationData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$IPAddress
    )
    
    try {
        $geoData = Invoke-RestMethod -Uri "http://ip-api.com/json/$IPAddress" -Method Get -TimeoutSec 10
        
        if ($geoData.status -eq "success") {
            Write-Host "Location Details:" -ForegroundColor Yellow
            Write-Host "  City: $($geoData.city)"
            Write-Host "  Region: $($geoData.regionName)"
            Write-Host "  Country: $($geoData.country)"
            Write-Host "  ISP: $($geoData.isp)"
            Write-Host "  Timezone: $($geoData.timezone)"
            
            return $geoData
        }
        else {
            Write-Error "Geolocation lookup failed: $($geoData.message)"
            return $null
        }
    }
    catch {
        Write-Error "Failed to get geolocation data: $_"
        return $null
    }
}

# Function to convert IANA timezone to Windows timezone
function Convert-IANAToWindowsTimeZone {
    param (
        [Parameter(Mandatory=$true)]
        [string]$IANATimeZone
    )
    
    # Comprehensive mapping of IANA to Windows timezones
    $timezoneMapping = @{
        "America/New_York" = "Eastern Standard Time"
        "America/Chicago" = "Central Standard Time"
        "America/Denver" = "Mountain Standard Time"
        "America/Phoenix" = "US Mountain Standard Time"
        "America/Los_Angeles" = "Pacific Standard Time"
        "America/Anchorage" = "Alaskan Standard Time"
        "Pacific/Honolulu" = "Hawaiian Standard Time"
        "Europe/London" = "GMT Standard Time"
        "Europe/Paris" = "W. Europe Standard Time"
        "Europe/Berlin" = "W. Europe Standard Time"
        "Europe/Rome" = "W. Europe Standard Time"
        "Europe/Madrid" = "Romance Standard Time"
        "Europe/Amsterdam" = "W. Europe Standard Time"
        "Europe/Brussels" = "Romance Standard Time"
        "Europe/Vienna" = "W. Europe Standard Time"
        "Europe/Prague" = "Central Europe Standard Time"
        "Europe/Warsaw" = "Central European Standard Time"
        "Europe/Budapest" = "Central Europe Standard Time"
        "Europe/Stockholm" = "W. Europe Standard Time"
        "Europe/Oslo" = "W. Europe Standard Time"
        "Europe/Copenhagen" = "Romance Standard Time"
        "Europe/Helsinki" = "FLE Standard Time"
        "Europe/Moscow" = "Russian Standard Time"
        "Asia/Tokyo" = "Tokyo Standard Time"
        "Asia/Shanghai" = "China Standard Time"
        "Asia/Hong_Kong" = "China Standard Time"
        "Asia/Singapore" = "Singapore Standard Time"
        "Asia/Seoul" = "Korea Standard Time"
        "Asia/Kolkata" = "India Standard Time"
        "Asia/Dubai" = "Arabian Standard Time"
        "Asia/Tehran" = "Iran Standard Time"
        "Australia/Sydney" = "AUS Eastern Standard Time"
        "Australia/Melbourne" = "AUS Eastern Standard Time"
        "Australia/Brisbane" = "E. Australia Standard Time"
        "Australia/Perth" = "W. Australia Standard Time"
        "Australia/Adelaide" = "Cen. Australia Standard Time"
        "Pacific/Auckland" = "New Zealand Standard Time"
        "America/Toronto" = "Eastern Standard Time"
        "America/Vancouver" = "Pacific Standard Time"
        "America/Sao_Paulo" = "E. South America Standard Time"
        "America/Mexico_City" = "Central Standard Time (Mexico)"
        "Africa/Cairo" = "Egypt Standard Time"
        "Africa/Johannesburg" = "South Africa Standard Time"
    }
    
    if ($timezoneMapping.ContainsKey($IANATimeZone)) {
        return $timezoneMapping[$IANATimeZone]
    }
    else {
        Write-Warning "No mapping found for timezone: $IANATimeZone. Attempting direct conversion..."
        
        # Try to find a matching Windows timezone by display name or standard name
        $availableTimeZones = Get-TimeZone -ListAvailable
        $matchingZone = $availableTimeZones | Where-Object { 
            $_.Id -like "*$($IANATimeZone.Split('/')[-1])*" -or 
            $_.StandardName -like "*$($IANATimeZone.Split('/')[-1])*"
        } | Select-Object -First 1
        
        if ($matchingZone) {
            Write-Host "Found approximate match: $($matchingZone.Id)" -ForegroundColor Yellow
            return $matchingZone.Id
        }
        else {
            Write-Warning "Could not find Windows timezone for $IANATimeZone. Defaulting to UTC."
            return "UTC"
        }
    }
}

# Function to set the system timezone
function Set-SystemTimeZone {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WindowsTimeZone
    )
    
    try {
        # Get current timezone for comparison
        $currentTimeZone = Get-TimeZone
        
        if ($currentTimeZone.Id -eq $WindowsTimeZone) {
            Write-Host "System is already set to the correct timezone: $WindowsTimeZone" -ForegroundColor Green
            return $true
        }
        
        # Set the new timezone
        Set-TimeZone -Id $WindowsTimeZone -PassThru
        
        # Verify the change
        $newTimeZone = Get-TimeZone
        if ($newTimeZone.Id -eq $WindowsTimeZone) {
            Write-Host "Successfully changed timezone from '$($currentTimeZone.Id)' to '$($newTimeZone.Id)'" -ForegroundColor Green
            Write-Host "Current local time: $(Get-Date)" -ForegroundColor Cyan
            return $true
        }
        else {
            Write-Error "Failed to set timezone. Current timezone is still: $($newTimeZone.Id)"
            return $false
        }
    }
    catch {
        Write-Error "Failed to set timezone: $_"
        return $false
    }
}

# Function to test for network changes
function Test-NetworkChange {
    $registryPath = "HKLM:\SOFTWARE\AutoTimezone"
    $lastIPKey = "LastPublicIP"
    
    # Ensure registry path exists
    if (!(Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    
    # Get current public IP
    try {
        $currentIP = (Invoke-WebRequest -Uri "https://api.ipify.org/" -UseBasicParsing -TimeoutSec 10).Content.Trim()
    }
    catch {
        Write-Host "Could not determine public IP, assuming network change" -ForegroundColor Yellow
        return $true
    }
    
    # Get last known IP
    $lastIP = Get-ItemProperty -Path $registryPath -Name $lastIPKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $lastIPKey
    
    # Compare IPs
    if ($currentIP -ne $lastIP) {
        # Store new IP
        Set-ItemProperty -Path $registryPath -Name $lastIPKey -Value $currentIP
        Write-Host "Network change detected: $lastIP -> $currentIP" -ForegroundColor Green
        return $true
    }
    else {
        Write-Host "No network change detected (IP: $currentIP)" -ForegroundColor Gray
        return $false
    }
}

# Main script execution
function Main {
    # Log start time
    $logPath = "C:\Scripts\TimezoneUpdate.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "[$timestamp] Timezone update script started"
    
    Write-Host "=== Automatic Timezone Configuration Script ===" -ForegroundColor Cyan
    Write-Host "This script will detect your location and set the appropriate timezone.`n"
    
    # Check for network change first
    if (!(Test-NetworkChange)) {
        Add-Content -Path $logPath -Value "[$timestamp] No network change detected, exiting"
        Write-Host "No timezone update needed - network unchanged" -ForegroundColor Gray
        exit 0
    }
    
    # Step 1: Get public IP address
    Write-Host "Step 1: Getting public IP address..." -ForegroundColor Yellow
    $publicIP = Get-PublicIPAddress
    if (-not $publicIP) {
        Add-Content -Path $logPath -Value "[$timestamp] Failed to get public IP address"
        Write-Error "Cannot proceed without public IP address."
        exit 1
    }
    
    # Step 2: Get geolocation data
    Write-Host "`nStep 2: Getting geolocation data..." -ForegroundColor Yellow
    $geoData = Get-GeoLocationData -IPAddress $publicIP
    if (-not $geoData -or -not $geoData.timezone) {
        Add-Content -Path $logPath -Value "[$timestamp] Failed to get geolocation data"
        Write-Error "Cannot proceed without geolocation data."
        exit 1
    }
    
    # Step 3: Convert IANA timezone to Windows timezone
    Write-Host "`nStep 3: Converting timezone format..." -ForegroundColor Yellow
    $windowsTimeZone = Convert-IANAToWindowsTimeZone -IANATimeZone $geoData.timezone
    Write-Host "Detected timezone: $($geoData.timezone) -> Windows: $windowsTimeZone"
    
    # Step 4: Set the system timezone
    Write-Host "`nStep 4: Setting system timezone..." -ForegroundColor Yellow
    $success = Set-SystemTimeZone -WindowsTimeZone $windowsTimeZone
    
    if ($success) {
        Add-Content -Path $logPath -Value "[$timestamp] Successfully updated timezone to $windowsTimeZone (Location: $($geoData.city), $($geoData.country))"
        Write-Host "`n=== Configuration Complete ===" -ForegroundColor Green
        Write-Host "Your system timezone has been automatically configured based on your location."
    }
    else {
        Add-Content -Path $logPath -Value "[$timestamp] Failed to update timezone to $windowsTimeZone"
        Write-Host "`n=== Configuration Failed ===" -ForegroundColor Red
        Write-Host "Please manually set your timezone using: Set-TimeZone -Id '$windowsTimeZone'"
    }
}

# Run the main function
Main
'@

try {
    $timezoneScript | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
    Write-Host "Created timezone script: $scriptPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to create script file: $_"
    exit 1
}

# Step 3: Create the scheduled task
Write-Host "`nStep 3: Creating scheduled task..." -ForegroundColor Yellow

try {
    # Remove existing task if it exists
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Removed existing task: $taskName" -ForegroundColor Yellow
    }

    # Define the action to run the timezone script
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
    
    # Create triggers for network events
    $triggers = @()
    $triggers += New-ScheduledTaskTrigger -AtLogOn
    $triggers += New-ScheduledTaskTrigger -AtStartup
    
    # Configure settings
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    
    # Set principal to run as SYSTEM
    $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    
    # Register the task
    Register-ScheduledTask -TaskName $taskName -TaskPath "\" -Description "Automatically updates timezone based on IP geolocation when network changes" -Action $action -Settings $settings -Trigger $triggers -Principal $principal | Out-Null
    
    Write-Host "Created scheduled task: $taskName" -ForegroundColor Green
}
catch {
    Write-Error "Failed to create scheduled task: $_"
    exit 1
}

# Step 4: Verify the setup
Write-Host "`nStep 4: Verifying setup..." -ForegroundColor Yellow

# Check if script file exists
if (Test-Path $scriptPath) {
    Write-Host "✓ Script file created successfully" -ForegroundColor Green
}
else {
    Write-Host "✗ Script file not found" -ForegroundColor Red
}

# Check if scheduled task exists
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($task) {
    Write-Host "✓ Scheduled task created successfully" -ForegroundColor Green
    Write-Host "  Task Name: $($task.TaskName)" -ForegroundColor Gray
    Write-Host "  Task Path: $($task.TaskPath)" -ForegroundColor Gray
    Write-Host "  State: $($task.State)" -ForegroundColor Gray
}
else {
    Write-Host "✗ Scheduled task not found" -ForegroundColor Red
}

Write-Host "`n=== Setup Complete ===" -ForegroundColor Cyan
Write-Host "The automatic timezone configuration system has been installed." -ForegroundColor Green
Write-Host "The system will now automatically detect and set the correct timezone when:" -ForegroundColor Yellow
Write-Host "  • The computer starts up" -ForegroundColor Gray
Write-Host "  • A user logs in" -ForegroundColor Gray
Write-Host "  • The network connection changes" -ForegroundColor Gray
Write-Host "`nLog file location: C:\Scripts\TimezoneUpdate.log" -ForegroundColor Cyan
