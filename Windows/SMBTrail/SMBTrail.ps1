#Requires -Version 5.1
<#
.SYNOPSIS
  SMBTrail inventories SMB sessions, open files, shares, and SMB1 audit events over time.
.DESCRIPTION
  Installs scheduled tasks that sample SMB activity, create CSV/JSON summaries, package results,
  and optionally deliver the report through SMTP, SendGrid, SMTP2GO, Postal, or a generic HTTP API.
.NOTES
  Project: SMBTrail
  License: MIT
#>
[CmdletBinding()]
param(
    [ValidateSet('Install','Collect','Finalize','Status','TestDelivery','Uninstall')]
    [string]$Mode = 'Status',
    [string]$ConfigPath = "$PSScriptRoot\SMBTrail.config.json"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-Config {
    if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
    $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    foreach ($required in 'CollectionDays','CollectionIntervalMinutes','OutputPath','Delivery') {
        if (-not $cfg.PSObject.Properties.Name.Contains($required)) { throw "Missing config property: $required" }
    }
    return $cfg
}

$Config = Get-Config
$InstalledScript = Join-Path $Config.OutputPath 'SMBTrail.ps1'
$InstalledConfig = Join-Path $Config.OutputPath 'SMBTrail.config.json'
$StatePath = Join-Path $Config.OutputPath 'state.json'
$SessionCsv = Join-Path $Config.OutputPath 'SMB-Sessions.csv'
$OpenFileCsv = Join-Path $Config.OutputPath 'SMB-OpenFiles.csv'
$Smb1Csv = Join-Path $Config.OutputPath 'SMB1-Audit-Events.csv'
$ShareCsv = Join-Path $Config.OutputPath 'SMB-Shares.csv'
$ServerConfigCsv = Join-Path $Config.OutputPath 'SMB-Server-Configuration.csv'
$ErrorLog = Join-Path $Config.OutputPath 'Errors.log'
$TaskPrefix = if ($Config.TaskPrefix) { [string]$Config.TaskPrefix } else { 'SMBTrail' }

function Initialize-Path { New-Item -Path $Config.OutputPath -ItemType Directory -Force | Out-Null }
function Write-Log([string]$Message) { Initialize-Path; "$(Get-Date -Format o) | $Message" | Out-File $ErrorLog -Append -Encoding utf8 }
function Save-State($State) { $State | ConvertTo-Json -Depth 5 | Set-Content $StatePath -Encoding utf8 }
function Get-State { if (Test-Path $StatePath) { Get-Content $StatePath -Raw | ConvertFrom-Json } }
function Resolve-Secret([string]$Value,[string]$EnvironmentVariable) {
    if ($EnvironmentVariable) {
        $resolved = [Environment]::GetEnvironmentVariable($EnvironmentVariable,'Machine')
        if (-not $resolved) { $resolved = [Environment]::GetEnvironmentVariable($EnvironmentVariable,'Process') }
        if (-not $resolved) { throw "Secret environment variable is empty: $EnvironmentVariable" }
        return $resolved
    }
    return $Value
}

function Export-Baseline {
    Get-SmbServerConfiguration | Select-Object EnableSMB1Protocol,EnableSMB2Protocol,AuditSmb1Access,EnableSecuritySignature,RequireSecuritySignature |
        Export-Csv $ServerConfigCsv -NoTypeInformation -Encoding utf8
    Get-SmbShare | Where-Object { -not $_.Special } | Select-Object Name,Path,Description,FolderEnumerationMode,EncryptData,ConcurrentUserLimit |
        Export-Csv $ShareCsv -NoTypeInformation -Encoding utf8
}

function Invoke-Collection {
    Initialize-Path
    $now = (Get-Date).ToString('o')
    try {
        @(Get-SmbSession -IncludeHidden) | Select-Object @{N='Captured';E={$now}},ClientComputerName,ClientUserName,Dialect,NumOpens,SecondsExists,SecondsIdle,SessionId |
            Export-Csv $SessionCsv -Append -NoTypeInformation -Encoding utf8
    } catch { Write-Log "Get-SmbSession: $($_.Exception.Message)" }
    try {
        @(Get-SmbOpenFile) | Select-Object @{N='Captured';E={$now}},ClientComputerName,ClientUserName,Path,ShareRelativePath,SessionId,FileId,Permissions,Locks |
            Export-Csv $OpenFileCsv -Append -NoTypeInformation -Encoding utf8
    } catch { Write-Log "Get-SmbOpenFile: $($_.Exception.Message)" }
}

function Export-Smb1Events([datetime]$Start,[datetime]$End) {
    try {
        $events = @(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-SMBServer/Audit';Id=3000;StartTime=$Start;EndTime=$End})
        $rows = foreach ($event in $events) {
            $xml = [xml]$event.ToXml(); $data = @{}
            foreach ($node in $xml.Event.EventData.Data) { $data[[string]$node.Name] = [string]$node.'#text' }
            [pscustomobject]@{
                TimeCreated=$event.TimeCreated.ToString('o'); EventId=$event.Id; MachineName=$event.MachineName
                ClientAddress=($data['ClientAddress'],$data['ClientIP'],$data['Address'] | Where-Object {$_} | Select-Object -First 1)
                ClientName=($data['ClientName'],$data['ComputerName'] | Where-Object {$_} | Select-Object -First 1)
                UserName=($data['UserName'],$data['User'] | Where-Object {$_} | Select-Object -First 1)
                Message=($event.Message -replace "`r?`n",' ')
            }
        }
        if ($rows) { $rows | Export-Csv $Smb1Csv -NoTypeInformation -Encoding utf8 }
        else { 'TimeCreated,EventId,MachineName,ClientAddress,ClientName,UserName,Message' | Set-Content $Smb1Csv -Encoding utf8 }
    } catch { Write-Log "SMB1 event export: $($_.Exception.Message)" }
}

function New-Report([datetime]$Start,[datetime]$End) {
    $sessions = if (Test-Path $SessionCsv) { @(Import-Csv $SessionCsv) } else { @() }
    $opens = if (Test-Path $OpenFileCsv) { @(Import-Csv $OpenFileCsv) } else { @() }
    $smb1 = if (Test-Path $Smb1Csv) { @(Import-Csv $Smb1Csv | Where-Object EventId) } else { @() }
    $clients = @($sessions.ClientComputerName + $opens.ClientComputerName | Where-Object {$_} | Sort-Object -Unique)
    $dialects = @($sessions | Where-Object Dialect | Group-Object Dialect | Sort-Object Count -Descending)
    $paths = @($opens | Where-Object Path | Group-Object Path | Sort-Object Count -Descending | Select-Object -First 50)
    $summary = [ordered]@{
        Computer=$env:COMPUTERNAME; Started=$Start.ToString('o'); Completed=$End.ToString('o')
        UniqueClients=$clients.Count; SessionObservations=$sessions.Count; OpenFileObservations=$opens.Count; Smb1Events=$smb1.Count
        Dialects=@($dialects | ForEach-Object {[ordered]@{Dialect=$_.Name;Observations=$_.Count}})
        Clients=$clients; TopPaths=@($paths | ForEach-Object {[ordered]@{Path=$_.Name;Observations=$_.Count}})
    }
    $summary | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $Config.OutputPath 'Summary.json') -Encoding utf8
    $text = @("SMBTrail report for $env:COMPUTERNAME","Observation window: $Start through $End","Unique clients: $($clients.Count)","Session observations: $($sessions.Count)","Open-file observations: $($opens.Count)","SMB1 events: $($smb1.Count)",'','Dialects:')
    foreach ($d in $dialects) { $text += "  $($d.Name): $($d.Count)" }
    $text += ''; $text += 'Clients:'; foreach ($c in $clients) { $text += "  $c" }
    $text -join "`r`n" | Set-Content (Join-Path $Config.OutputPath 'Summary.txt') -Encoding utf8
    return ($text -join "`r`n")
}

function New-Archive {
    $zip = Join-Path $Config.OutputPath ("SMBTrail-{0}-{1}.zip" -f $env:COMPUTERNAME,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    $files = Get-ChildItem $Config.OutputPath -File | Where-Object {$_.Extension -in '.csv','.json','.txt','.log' -and $_.Name -ne 'SMBTrail.config.json'}
    Compress-Archive -Path $files.FullName -DestinationPath $zip -Force
    return $zip
}

function Send-Smtp([string]$Subject,[string]$Body,[string]$Attachment) {
    $d=$Config.Delivery; $password=Resolve-Secret $d.Password $d.PasswordEnvironmentVariable
    $mail=New-Object System.Net.Mail.MailMessage; $mail.From=$d.From; [void]$mail.To.Add($d.To); $mail.Subject=$Subject; $mail.Body=$Body; [void]$mail.Attachments.Add($Attachment)
    $client=New-Object System.Net.Mail.SmtpClient($d.Host,[int]$d.Port); $client.EnableSsl=[bool]$d.UseTls
    if ($d.Username) { $client.Credentials=New-Object System.Net.NetworkCredential($d.Username,$password) }
    $client.Send($mail); $mail.Dispose(); $client.Dispose()
}
function Send-SendGrid([string]$Subject,[string]$Body,[string]$Attachment) {
    $d=$Config.Delivery; $key=Resolve-Secret $d.ApiKey $d.ApiKeyEnvironmentVariable; $b64=[Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))
    $payload=@{personalizations=@(@{to=@(@{email=$d.To});subject=$Subject});from=@{email=$d.From;name=$d.FromName};content=@(@{type='text/plain';value=$Body});attachments=@(@{content=$b64;type='application/zip';filename=[IO.Path]::GetFileName($Attachment);disposition='attachment'})}|ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method Post -Uri 'https://api.sendgrid.com/v3/mail/send' -Headers @{Authorization="Bearer $key"} -ContentType 'application/json' -Body ([Text.Encoding]::UTF8.GetBytes($payload)) | Out-Null
}
function Send-Smtp2Go([string]$Subject,[string]$Body,[string]$Attachment) {
    $d=$Config.Delivery; $key=Resolve-Secret $d.ApiKey $d.ApiKeyEnvironmentVariable; $b64=[Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))
    $payload=@{sender=$d.From;to=@($d.To);subject=$Subject;text_body=$Body;attachments=@(@{filename=[IO.Path]::GetFileName($Attachment);fileblob=$b64;mimetype='application/zip'})}|ConvertTo-Json -Depth 8
    Invoke-RestMethod -Method Post -Uri 'https://api.smtp2go.com/v3/email/send' -Headers @{'X-Smtp2go-Api-Key'=$key} -ContentType 'application/json' -Body $payload | Out-Null
}
function Send-Postal([string]$Subject,[string]$Body,[string]$Attachment) {
    $d=$Config.Delivery; $key=Resolve-Secret $d.ApiKey $d.ApiKeyEnvironmentVariable; $b64=[Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))
    $uri=($d.BaseUrl.TrimEnd('/') + '/api/v1/send/message')
    $payload=@{to=@($d.To);from=$d.From;sender=$d.From;subject=$Subject;plain_body=$Body;attachments=@(@{name=[IO.Path]::GetFileName($Attachment);data=$b64})}|ConvertTo-Json -Depth 8
    Invoke-RestMethod -Method Post -Uri $uri -Headers @{'X-Server-API-Key'=$key} -ContentType 'application/json' -Body $payload | Out-Null
}
function Send-GenericApi([string]$Subject,[string]$Body,[string]$Attachment) {
    $d=$Config.Delivery; $secret=Resolve-Secret $d.ApiKey $d.ApiKeyEnvironmentVariable; $b64=[Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))
    $headers=@{}; if($d.HeaderName){$headers[$d.HeaderName]=("$($d.HeaderPrefix)$secret")}
    $payload=@{from=$d.From;to=@($d.To);subject=$Subject;text=$Body;attachment=@{filename=[IO.Path]::GetFileName($Attachment);contentBase64=$b64}}|ConvertTo-Json -Depth 8
    Invoke-RestMethod -Method Post -Uri $d.Uri -Headers $headers -ContentType 'application/json' -Body $payload | Out-Null
}
function Send-Report([string]$Subject,[string]$Body,[string]$Attachment) {
    switch ([string]$Config.Delivery.Method) {
        'None' { return }
        'SMTP' { Send-Smtp $Subject $Body $Attachment }
        'SendGrid' { Send-SendGrid $Subject $Body $Attachment }
        'SMTP2GO' { Send-Smtp2Go $Subject $Body $Attachment }
        'Postal' { Send-Postal $Subject $Body $Attachment }
        'GenericApi' { Send-GenericApi $Subject $Body $Attachment }
        default { throw "Unsupported Delivery.Method: $($Config.Delivery.Method)" }
    }
}

function Install-SMBTrail {
    Initialize-Path

    # Only copy the script if it is not already running from the install path
    $sourceScript = [IO.Path]::GetFullPath($PSCommandPath)
    $destinationScript = [IO.Path]::GetFullPath($InstalledScript)

    if ($sourceScript -ne $destinationScript) {
        Copy-Item -Path $PSCommandPath -Destination $InstalledScript -Force
    }

    # Only copy the config if it is not already the installed config
    $sourceConfig = [IO.Path]::GetFullPath($ConfigPath)
    $destinationConfig = [IO.Path]::GetFullPath($InstalledConfig)

    if ($sourceConfig -ne $destinationConfig) {
        Copy-Item -Path $ConfigPath -Destination $InstalledConfig -Force
    }

    Export-Baseline

    $start = Get-Date
    $end = $start.AddDays([double]$Config.CollectionDays)

    Save-State @{
        StartTime = $start.ToString('o')
        EndTime   = $end.ToString('o')
        Server    = $env:COMPUTERNAME
    }

    $principal = New-ScheduledTaskPrincipal `
        -UserId 'SYSTEM' `
        -LogonType ServiceAccount `
        -RunLevel Highest

    $settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

    $collectAction = New-ScheduledTaskAction `
        -Execute 'PowerShell.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$InstalledScript`" -Mode Collect -ConfigPath `"$InstalledConfig`""

    $collectTrigger = New-ScheduledTaskTrigger `
        -Once `
        -At $start.AddMinutes(1) `
        -RepetitionInterval (New-TimeSpan -Minutes ([int]$Config.CollectionIntervalMinutes)) `
        -RepetitionDuration (New-TimeSpan -Days ([double]$Config.CollectionDays + 1))

    Register-ScheduledTask `
        -TaskName "$TaskPrefix-Collect" `
        -Action $collectAction `
        -Trigger $collectTrigger `
        -Principal $principal `
        -Settings $settings `
        -Force | Out-Null

    $finalizeAction = New-ScheduledTaskAction `
        -Execute 'PowerShell.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$InstalledScript`" -Mode Finalize -ConfigPath `"$InstalledConfig`""

    $finalizeTrigger = New-ScheduledTaskTrigger `
        -Once `
        -At $end

    $finalizeSettings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

    Register-ScheduledTask `
        -TaskName "$TaskPrefix-Finalize" `
        -Action $finalizeAction `
        -Trigger $finalizeTrigger `
        -Principal $principal `
        -Settings $finalizeSettings `
        -Force | Out-Null

    Invoke-Collection

    Write-Host "SMBTrail installed successfully."
    Write-Host "Collection started: $start"
    Write-Host "Finalization scheduled: $end"
    Write-Host "Collection interval: $($Config.CollectionIntervalMinutes) minutes"
    Write-Host "Delivery method: $($Config.Delivery.Method)"
}
function Finalize-SMBTrail {
    $state=Get-State; if(-not $state){throw 'State file not found'}; $start=[datetime]$state.StartTime; $end=Get-Date
    Invoke-Collection; Export-Smb1Events $start $end; $body=New-Report $start $end; $zip=New-Archive
    Send-Report "SMBTrail complete - $env:COMPUTERNAME" $body $zip
    Unregister-ScheduledTask "$TaskPrefix-Collect" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask "$TaskPrefix-Finalize" -Confirm:$false -ErrorAction SilentlyContinue
    "Complete. Report: $zip"
}
function Test-Delivery { Initialize-Path; $f=Join-Path $Config.OutputPath 'SMBTrail-delivery-test.txt'; 'SMBTrail delivery test'|Set-Content $f; $z=Join-Path $Config.OutputPath 'SMBTrail-delivery-test.zip'; Compress-Archive $f $z -Force; Send-Report 'SMBTrail delivery test' 'Delivery is configured correctly.' $z; 'Delivery test completed.' }
function Show-Status {
    $s = Get-State

    $collectTask = Get-ScheduledTask `
        -TaskName "$TaskPrefix-Collect" `
        -ErrorAction SilentlyContinue

    $finalizeTask = Get-ScheduledTask `
        -TaskName "$TaskPrefix-Finalize" `
        -ErrorAction SilentlyContinue

    [pscustomobject]@{
        Computer = $env:COMPUTERNAME
        Installed = [bool]$s
        Start = if ($s -and $s.PSObject.Properties['StartTime']) {
            $s.StartTime
        } else {
            'Not initialized'
        }
        End = if ($s -and $s.PSObject.Properties['EndTime']) {
            $s.EndTime
        } else {
            'Not initialized'
        }
        DeliveryMethod = $Config.Delivery.Method
        OutputPath = $Config.OutputPath
        CollectTask = if ($collectTask) {
            $collectTask.State
        } else {
            'Not installed'
        }
        FinalizeTask = if ($finalizeTask) {
            $finalizeTask.State
        } else {
            'Not installed'
        }
        StateFile = if (Test-Path $StatePath) {
            'Present'
        } else {
            'Missing'
        }
        SessionData = Test-Path $SessionCsv
        OpenFileData = Test-Path $OpenFileCsv
    } | Format-List
}
function Uninstall-SMBTrail { Unregister-ScheduledTask "$TaskPrefix-Collect" -Confirm:$false -ErrorAction SilentlyContinue; Unregister-ScheduledTask "$TaskPrefix-Finalize" -Confirm:$false -ErrorAction SilentlyContinue; 'Tasks removed. Output retained.' }

switch($Mode){'Install'{Install-SMBTrail};'Collect'{Invoke-Collection};'Finalize'{Finalize-SMBTrail};'Status'{Show-Status};'TestDelivery'{Test-Delivery};'Uninstall'{Uninstall-SMBTrail}}
