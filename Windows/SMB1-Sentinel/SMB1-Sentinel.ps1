<#
.SYNOPSIS
SMB1 Sentinel: temporary inbound SMB1 session sampling with provider-neutral notifications.
.DESCRIPTION
Edit the configuration below and run from elevated Windows PowerShell 5.1.
Default mode launches a separate process. No files, modules, services or tasks are installed.
.PARAMETER Foreground
Run in the current terminal for troubleshooting. The terminal must remain open.
.PARAMETER ValidateOnly
Parse launcher and worker with the native PowerShell parser. Does not query SMB,
send notifications, validate credentials, or start a worker.
#>
[CmdletBinding()]
param([switch]$Foreground, [switch]$ValidateOnly)

# ========================= CONFIGURATION =========================
$Provider = 'SendGrid' # SendGrid, Slack, Discord, Pushover, Ntfy, Webhook
$DurationHours = 48
$PollSeconds = 5
$NotificationTimeoutSeconds = 15
$NotificationMaxAttempts = 3
$IncludeUserNames = $false
$ProviderSettings = @{
    SendGrid = @{ ApiKey = 'SG_KEYGOESHERE'; From = 'sender@example.com'; To = @('recipient@example.com') }
    Slack = @{ Uri = 'YOUR_SLACK_HTTPS_WEBHOOK_URL' }
    Discord = @{ Uri = 'YOUR_DISCORD_HTTPS_WEBHOOK_URL' }
    Pushover = @{ Token = 'YOUR_PUSHOVER_APPLICATION_TOKEN'; User = 'YOUR_PUSHOVER_USER_KEY' }
    Ntfy = @{ Uri = 'YOUR_PROTECTED_HTTPS_TOPIC_URL'; Token = 'YOUR_NTFY_ACCESS_TOKEN' }
    Webhook = @{ Uri = 'YOUR_HTTPS_WEBHOOK_URL'; Token = '' }
}
# ======================= IMPLEMENTATION ==========================
$ErrorActionPreference = 'Stop'
$workerSource = @'
param([Parameter(Mandatory = $true)]$Config)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$script:DeliveryFailures = 0
$script:HostName = $env:COMPUTERNAME

function Split-NotificationText {
    param([string]$Text, [int]$Limit = 800)
    if ([string]::IsNullOrEmpty($Text)) { return '(empty)' }
    $offset = 0
    while ($offset -lt $Text.Length) {
        $length = [Math]::Min($Limit, $Text.Length - $offset)
        # Do not split a UTF-16 surrogate pair between requests.
        if (($offset + $length) -lt $Text.Length -and
            [char]::IsHighSurrogate($Text[$offset + $length - 1])) { $length-- }
        $Text.Substring($offset, $length)
        $offset += $length
    }
}

function New-NotificationRequest {
    param([string]$Title, [string]$Message, [string]$NotificationId)
    $s = $Config.Settings
    $headers = @{}
    $contentType = 'application/json; charset=utf-8'
    $text = "$Title`n`n$Message"
    switch ($Config.Provider) {
        'SendGrid' {
            $uri = 'https://api.sendgrid.com/v3/mail/send'
            $headers.Authorization = "Bearer $($s.ApiKey)"
            $recipients = @(foreach ($address in @($s.To)) { @{ email = [string]$address } })
            $payload = @{
                personalizations = @(@{ to = $recipients })
                from = @{ email = [string]$s.From }
                subject = $Title
                content = @(@{ type = 'text/plain'; value = $Message })
            }
        }
        'Slack' {
            $uri = [string]$s.Uri
            # Escape Slack control sequences in externally derived identifiers.
            $escaped = $text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
            $payload = @{ text = $escaped; mrkdwn = $false; link_names = 0; unfurl_links = $false; unfurl_media = $false }
        }
        'Discord' {
            $builder = New-Object System.UriBuilder([string]$s.Uri)
            $query = @($builder.Query.TrimStart('?').Split('&') | Where-Object { $_ -and $_ -notmatch '^wait=' })
            $builder.Query = (@($query) + @('wait=true')) -join '&'
            $uri = $builder.Uri.AbsoluteUri
            $payload = @{ content = $text; allowed_mentions = @{ parse = @() } }
        }
        'Pushover' {
            $uri = 'https://api.pushover.net/1/messages.json'
            $payload = @{ token = [string]$s.Token; user = [string]$s.User; title = $Title; message = $Message }
        }
        'Ntfy' {
            $uri = [string]$s.Uri
            $headers.Authorization = "Bearer $($s.Token)"
            $contentType = 'text/plain; charset=utf-8'
        }
        'Webhook' {
            $uri = [string]$s.Uri
            if (-not [string]::IsNullOrWhiteSpace([string]$s.Token)) { $headers.Authorization = "Bearer $($s.Token)" }
            $payload = @{ title = $Title; message = $Message; host = $script:HostName; timeUtc = [datetime]::UtcNow.ToString('o'); notificationId = $NotificationId }
        }
        default { throw 'Unsupported notification provider.' }
    }
    $target = [uri]$uri
    if (-not $target.IsAbsoluteUri -or $target.Scheme -ne 'https' -or $target.UserInfo -or $target.Fragment) { throw 'Invalid HTTPS endpoint.' }
    $body = if ($Config.Provider -eq 'Ntfy') { $text } else { $payload | ConvertTo-Json -Depth 8 -Compress }
    return @{
        Uri = $target; Method = 'Post'; Headers = $headers
        ContentType = $contentType; Body = [Text.Encoding]::UTF8.GetBytes($body)
        TimeoutSec = [int]$Config.TimeoutSeconds; MaximumRedirection = 0
        UseBasicParsing = $true; ErrorAction = 'Stop'
    }
}

function Test-ProviderResponse {
    param($Response)
    if ([int]$Response.StatusCode -lt 200 -or [int]$Response.StatusCode -ge 300) { return $false }
    try {
    switch ($Config.Provider) {
        'SendGrid' { return ([int]$Response.StatusCode -eq 202) }
        'Slack' { return ([string]$Response.Content).Trim() -eq 'ok' }
        'Pushover' { return (($Response.Content | ConvertFrom-Json).status -eq 1) }
        'Discord' { return -not [string]::IsNullOrWhiteSpace([string]($Response.Content | ConvertFrom-Json).id) }
        'Ntfy' { return -not [string]::IsNullOrWhiteSpace([string]($Response.Content | ConvertFrom-Json).id) }
        'Webhook' { return $true }
    }
    } catch { return $false }
    return $false
}

function Send-NotificationPart {
    param([string]$Title, [string]$Message, [string]$NotificationId)
    $request = New-NotificationRequest -Title $Title -Message $Message -NotificationId $NotificationId
    for ($attempt = 1; $attempt -le [int]$Config.MaxAttempts; $attempt++) {
        try {
            $response = Invoke-WebRequest @request
            if (Test-ProviderResponse -Response $response) { return $true }
            # A 2xx response with an invalid provider acknowledgement is not retried.
            break
        }
        catch {
            $status = 0
            $retryAfter = 0
            $httpResponse = $_.Exception.Response
            if ($null -ne $httpResponse) {
                $status = [int]$httpResponse.StatusCode
                $rawRetry = [string]$httpResponse.Headers['Retry-After']
                if (-not [int]::TryParse($rawRetry, [ref]$retryAfter)) {
                    $retryDate = [datetimeoffset]::MinValue
                    if ([datetimeoffset]::TryParse($rawRetry, [ref]$retryDate)) {
                        $retryAfter = [int][Math]::Ceiling(($retryDate - [datetimeoffset]::UtcNow).TotalSeconds)
                    }
                }
            }
            $retryable = $status -eq 0 -or $status -eq 408 -or $status -eq 429 -or ($status -ge 500 -and $status -le 599)
            if (-not $retryable -or $attempt -eq [int]$Config.MaxAttempts -or $retryAfter -gt 60) { break }
            $delay = [Math]::Max([Math]::Max(0, $retryAfter), [int][Math]::Pow(2, $attempt))
            Start-Sleep -Seconds $delay
        }
    }
    $script:DeliveryFailures++
    Write-Warning "Notification part was not confirmed by $($Config.Provider). Raw errors suppressed to protect secrets."
    return $false
}

function Send-Notification {
    param([Parameter(Mandatory = $true)][string]$Title, [Parameter(Mandatory = $true)][string]$Message)
    # Leave room for part numbering and provider title limits.
    $Title = ($Title -replace '[\r\n]', ' ')
    if ($Title.Length -gt 140) {
        $n = 140
        if ([char]::IsHighSurrogate($Title[$n - 1])) { $n-- }
        $Title = $Title.Substring(0, $n)
    }
    $parts = @(Split-NotificationText -Text $Message)
    $allAccepted = $true
    $id = [guid]::NewGuid().ToString('N')
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $partTitle = $Title
        if ($parts.Count -gt 1) { $partTitle += " [$($i + 1)/$($parts.Count)]" }
        if (-not (Send-NotificationPart -Title $partTitle -Message $parts[$i] -NotificationId "$id-$i")) { $allAccepted = $false }
    }
    return $allAccepted
}

function Get-ValidatedSessions {
    $sessions = @(Get-SmbSession -ErrorAction Stop)
    foreach ($session in $sessions) {
        $dialect = ([string]$session.Dialect).Trim()
        if ($dialect -notmatch '^\d+(\.\d+)*$') { throw 'Session dialect is missing or not recognized.' }
        if ([string]::IsNullOrWhiteSpace([string]$session.ClientComputerName)) { throw 'Session client identifier is missing.' }
    }
    return $sessions
}

# A fixed name intentionally prevents overlapping runs, even with different providers.
$mutex = New-Object System.Threading.Mutex($false, 'Global\SMB1SentinelMonitor')
$ownsMutex = $false
$exitCode = 1
try {
    try { $ownsMutex = $mutex.WaitOne(0) }
    catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
    if (-not $ownsMutex) {
        $null = Send-Notification -Title "[$script:HostName] SMB1 watch not started" -Message 'Another SMB1 Sentinel worker owns the monitor mutex.'
        $exitCode = 2
    }
    else {
        $null = Get-Command Get-SmbSession -ErrorAction Stop
        $null = Get-ValidatedSessions
        $seen = @{}
        $polls = 0
        $failed = $false
        $startedUtc = [datetime]::UtcNow
        $plannedEnd = $startedUtc.AddHours([double]$Config.DurationHours)
        $started = Send-Notification -Title "[$script:HostName] SMB1 watch started" -Message (
            "Worker PID: $PID`nStart UTC: $($startedUtc.ToString('o'))`nPlanned end UTC: $($plannedEnd.ToString('o'))`n" +
            "Poll delay: $($Config.PollSeconds) seconds.`nInbound sessions only. Sampling can miss short connections."
        )
        if (-not $started) { throw 'Startup notification was not confirmed.' }
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $durationSeconds = [double]$Config.DurationHours * 3600
        try {
            while ($timer.Elapsed.TotalSeconds -lt $durationSeconds) {
                $sessions = @(Get-ValidatedSessions)
                $polls++
                foreach ($session in $sessions) {
                    if (([string]$session.Dialect).Trim() -notmatch '^1(\.\d+)*$') { continue }
                    $client = [string]$session.ClientComputerName
                    $now = [datetime]::UtcNow.ToString('o')
                    if ($seen.ContainsKey($client)) { $seen[$client].LastSeenUtc = $now; continue }
                    $seen[$client] = @{ FirstSeenUtc = $now; LastSeenUtc = $now; InitialAlertAccepted = $false }
                    $lines = @("Client: $client", "Dialect: $($session.Dialect)", "First observed UTC: $now")
                    if ($Config.IncludeUsers) { $lines += "User: $($session.ClientUserName)" }
                    $seen[$client].InitialAlertAccepted = Send-Notification -Title "[$script:HostName] SMB1 client observed" -Message ($lines -join "`n")
                }
                $remaining = $durationSeconds - $timer.Elapsed.TotalSeconds
                if ($remaining -gt 0) {
                    $sleepMs = [int][Math]::Min([double]$Config.PollSeconds * 1000, $remaining * 1000)
                    if ($sleepMs -gt 0) { Start-Sleep -Milliseconds $sleepMs }
                }
            }
        }
        catch { $failed = $true; Write-Warning 'SMB session enumeration or validation failed. Observation window is incomplete.' }
        finally { $timer.Stop() }
        $endedUtc = [datetime]::UtcNow.ToString('o')
        # Send details before the final summary so their failures are counted.
        foreach ($client in @($seen.Keys | Sort-Object)) {
            $record = $seen[$client]
            $detail = @("Client: $client", "First observed UTC: $($record.FirstSeenUtc)", "Last observed UTC: $($record.LastSeenUtc)", "Initial alert accepted: $($record.InitialAlertAccepted)") -join "`n"
            $null = Send-Notification -Title "[$script:HostName] SMB1 result" -Message $detail
        }
        $status = if ($failed) { 'INCOMPLETE' } else { 'COMPLETED' }
        $summary = @(
            "Observation status: $status"
            "Start UTC: $($startedUtc.ToString('o'))"
            "End UTC: $endedUtc"
            "Successful polls: $polls"
            "Unique SMB1 client identifiers: $($seen.Count)"
            "Failed notification parts before this summary: $script:DeliveryFailures"
            'Zero observations does not prove SMB1 is unused.'
            'Notification delivery pauses sampling. This summary cannot report its own delivery failure.'
        ) -join "`n"
        $summaryAccepted = Send-Notification -Title "[$script:HostName] SMB1 watch $status" -Message $summary
        if (-not $failed -and $summaryAccepted -and $script:DeliveryFailures -eq 0) { $exitCode = 0 }
    }
}
catch {
    Write-Warning 'Worker initialization or execution failed. Raw error details suppressed to protect secrets.'
    try { $null = Send-Notification -Title "[$script:HostName] SMB1 watch FAILED" -Message 'Worker initialization or execution failed. Do not treat this run as a completed observation window.' } catch { }
}
finally {
    if ($ownsMutex) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
# Return a code rather than exiting the caller's interactive shell.
return [int]$exitCode
'@

# Parse the generated worker before attempting to launch it.
$workerTokens = $null
$workerErrors = $null
$null = [System.Management.Automation.Language.Parser]::ParseInput($workerSource, [ref]$workerTokens, [ref]$workerErrors)
if ($workerErrors.Count -gt 0) { throw 'Worker source failed native PowerShell parsing.' }
if ($ValidateOnly) {
    if ($PSCommandPath) {
        $tokens = $null
        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) { throw 'Launcher failed native PowerShell parsing.' }
    }
    [pscustomobject]@{ Validation = 'Native parser passed'; WorkerChecked = $true; RuntimeTested = $false }
    return
}
if ($env:OS -ne 'Windows_NT' -or $PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -lt 1) {
    throw 'Run operational modes in Windows PowerShell 5.1. ValidateOnly can be used separately.'
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run from an elevated Windows PowerShell terminal.' }
if ($Provider -notin @('SendGrid', 'Slack', 'Discord', 'Pushover', 'Ntfy', 'Webhook')) { throw 'Unsupported provider.' }
if ($DurationHours -le 0 -or $DurationHours -gt 720 -or [double]::IsNaN([double]$DurationHours)) { throw 'DurationHours must be greater than zero and no more than 720.' }
if ($PollSeconds -lt 1 -or $PollSeconds -gt 3600 -or [double]$PollSeconds -ne [int]$PollSeconds) { throw 'PollSeconds must be an integer from 1 to 3600.' }
if ($NotificationTimeoutSeconds -lt 1 -or $NotificationTimeoutSeconds -gt 120 -or [double]$NotificationTimeoutSeconds -ne [int]$NotificationTimeoutSeconds) { throw 'NotificationTimeoutSeconds must be an integer from 1 to 120.' }
if ($NotificationMaxAttempts -lt 1 -or $NotificationMaxAttempts -gt 5 -or [double]$NotificationMaxAttempts -ne [int]$NotificationMaxAttempts) { throw 'NotificationMaxAttempts must be an integer from 1 to 5.' }
$selected = $ProviderSettings[$Provider]
$required = switch ($Provider) {
    'SendGrid' { @('ApiKey', 'From', 'To') }
    'Slack' { @('Uri') }
    'Discord' { @('Uri') }
    'Pushover' { @('Token', 'User') }
    'Ntfy' { @('Uri', 'Token') }
    'Webhook' { @('Uri') }
}
foreach ($name in $required) {
    $values = @($selected[$name])
    if ($values.Count -eq 0) { throw "Missing setting: $Provider.$name" }
    foreach ($value in $values) {
        if ([string]::IsNullOrWhiteSpace([string]$value) -or [string]$value -match '^(YOUR_|SG_KEYGOESHERE$)' -or [string]$value -in @('sender@example.com', 'recipient@example.com')) { throw "Replace the placeholder for $Provider.$name." }
    }
}
foreach ($name in @('ApiKey', 'Token')) {
    if ($selected.ContainsKey($name) -and [string]$selected[$name] -match '[\r\n]') { throw 'Credential values cannot contain line breaks.' }
}
if ($selected.ContainsKey('Uri')) {
    $target = [uri]$selected.Uri
    if (-not $target.IsAbsoluteUri -or $target.Scheme -ne 'https' -or $target.UserInfo -or $target.Fragment) { throw 'Use an absolute HTTPS endpoint without user-info or a fragment.' }
}
if ($Provider -eq 'SendGrid') {
    foreach ($address in @($selected.From) + @($selected.To)) { $null = New-Object System.Net.Mail.MailAddress([string]$address) }
}
$null = Get-Command Get-SmbSession -ErrorAction Stop
$null = Get-SmbSession -ErrorAction Stop
$config = @{
    Provider = $Provider; Settings = $selected; DurationHours = [double]$DurationHours
    PollSeconds = [int]$PollSeconds; TimeoutSeconds = [int]$NotificationTimeoutSeconds
    MaxAttempts = [int]$NotificationMaxAttempts; IncludeUsers = [bool]$IncludeUserNames
}
if ($Foreground) {
    $code = & ([scriptblock]::Create($workerSource)) -Config $config
    [pscustomobject]@{ Script = 'SMB1 Sentinel'; Mode = 'Foreground'; WorkerExitCode = [int]$code }
    return
}

# Compress source before encryption to stay below the process command-line limit.
# The same Windows identity must be able to decrypt current-user DPAPI data.
Add-Type -AssemblyName System.Security
$package = @{ Config = $config; Worker = $workerSource } | ConvertTo-Json -Depth 10 -Compress
$plain = [Text.Encoding]::UTF8.GetBytes($package)
$memory = New-Object IO.MemoryStream
$gzip = New-Object IO.Compression.GZipStream($memory, [IO.Compression.CompressionMode]::Compress, $true)
try { $gzip.Write($plain, 0, $plain.Length) } finally { $gzip.Dispose(); [Array]::Clear($plain, 0, $plain.Length) }
$compressed = $memory.ToArray()
$memory.Dispose()
$protected = [Security.Cryptography.ProtectedData]::Protect($compressed, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
[Array]::Clear($compressed, 0, $compressed.Length)
$blob = [Convert]::ToBase64String($protected)
$bootstrap = @'
$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Security
    $bytes = [Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String('__BLOB__'), $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $inputStream = New-Object IO.MemoryStream(,$bytes)
    $gzip = New-Object IO.Compression.GZipStream($inputStream, [IO.Compression.CompressionMode]::Decompress)
    $reader = New-Object IO.StreamReader($gzip, [Text.Encoding]::UTF8)
    try { $package = $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose(); $inputStream.Dispose(); [Array]::Clear($bytes, 0, $bytes.Length) }
    $code = & ([scriptblock]::Create($package.Worker)) -Config $package.Config
    exit ([int]$code)
}
catch { exit 1 }
'@
$bootstrap = $bootstrap.Replace('__BLOB__', $blob)
$bootstrapTokens = $null
$bootstrapErrors = $null
$null = [System.Management.Automation.Language.Parser]::ParseInput($bootstrap, [ref]$bootstrapTokens, [ref]$bootstrapErrors)
if ($bootstrapErrors.Count -gt 0) { throw 'Bootstrap failed native parsing.' }
$encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($bootstrap))
$exe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) { $exe = "$env:WINDIR\Sysnative\WindowsPowerShell\v1.0\powershell.exe" }
$command = '"' + $exe + '" -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand ' + $encoded
if ($command.Length -gt 30000) { throw 'Encrypted launch command is too long. Reduce configuration size or use -Foreground.' }
$result = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $command } -ErrorAction Stop
if ($result.ReturnValue -ne 0) { throw "Process creation failed. Return value: $($result.ReturnValue)" }
[pscustomobject]@{
    Script = 'SMB1 Sentinel'; ProcessId = $result.ProcessId; Provider = $Provider
    DurationHours = $DurationHours; LaunchStatus = 'Process created. Confirm startup notification.'
    StopCommand = "Stop-Process -Id $($result.ProcessId)"
}
