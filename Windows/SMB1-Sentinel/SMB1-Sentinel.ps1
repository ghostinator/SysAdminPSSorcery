# SMB1 Sentinel
# Windows PowerShell 5.1
#
# Polls inbound SMB sessions for SMB1.
# Runs in a separate process after this terminal closes.
# Does not install software or change SMB configuration.
#
# Limitations:
# - Polling can miss short-lived connections.
# - A reboot or process termination stops monitoring.
# - Notification retries can produce duplicates.
# - Results are held in memory, not written to disk.
# - Provider acceptance does not guarantee recipient delivery.

# ========================= CONFIGURATION =========================

$Provider = 'SendGrid'
# Supported: SendGrid, Slack, Discord, Pushover, Ntfy, Webhook

$DurationHours = 48
$PollSeconds = 5
$NotificationTimeoutSeconds = 15
$NotificationMaxAttempts = 3

# False avoids sending account names to notification providers.
$IncludeUserNames = $false

# Only the selected provider's settings are used.
$ProviderSettings = @{
    SendGrid = @{
        ApiKey = 'SG_KEYGOESHERE'
        From   = 'sender@sender.com'
        To     = @('recipient@recipient.com')
    }

    Slack = @{
        Uri = 'YOUR_SLACK_HTTPS_WEBHOOK_URL'
    }

    Discord = @{
        Uri = 'YOUR_DISCORD_HTTPS_WEBHOOK_URL'
    }

    Pushover = @{
        Token = 'YOUR_PUSHOVER_APPLICATION_TOKEN'
        User  = 'YOUR_PUSHOVER_USER_KEY'
    }

    Ntfy = @{
        # Use an access-controlled topic, not an open public topic.
        Uri   = 'YOUR_PROTECTED_HTTPS_TOPIC_URL'
        Token = 'YOUR_NTFY_ACCESS_TOKEN'
    }

    Webhook = @{
        # Receiver must accept the JSON schema documented below.
        Uri   = 'YOUR_HTTPS_WEBHOOK_URL'
        Token = ''
    }
}

# ========================= LAUNCHER CHECKS ========================

$ErrorActionPreference = 'Stop'

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)) {
    throw 'Run this script from an elevated PowerShell terminal.'
}

if ($Provider -notin @(
    'SendGrid', 'Slack', 'Discord', 'Pushover', 'Ntfy', 'Webhook'
)) {
    throw "Unsupported notification provider: $Provider"
}

if ($DurationHours -le 0 -or $DurationHours -gt 720) {
    throw 'DurationHours must be greater than zero and no more than 720.'
}

if ($PollSeconds -lt 1 -or $PollSeconds -gt 3600) {
    throw 'PollSeconds must be between 1 and 3600.'
}

if ($NotificationTimeoutSeconds -lt 1 -or
    $NotificationTimeoutSeconds -gt 120) {
    throw 'NotificationTimeoutSeconds must be between 1 and 120.'
}

if ($NotificationMaxAttempts -lt 1 -or
    $NotificationMaxAttempts -gt 5) {
    throw 'NotificationMaxAttempts must be between 1 and 5.'
}

$selected = $ProviderSettings[$Provider]

$required = switch ($Provider) {
    'SendGrid' { @('ApiKey', 'From', 'To') }
    'Slack'    { @('Uri') }
    'Discord'  { @('Uri') }
    'Pushover' { @('Token', 'User') }
    'Ntfy'     { @('Uri', 'Token') }
    'Webhook'  { @('Uri') }
}

foreach ($name in $required) {
    $values = @($selected[$name])

    if ($values.Count -eq 0) {
        throw "Missing configuration: $Provider.$name"
    }

    foreach ($value in $values) {
        if ([string]::IsNullOrWhiteSpace([string]$value) -or
            [string]$value -match '^(YOUR_|SG_KEYGOESHERE$)' -or
            [string]$value -in @(
                'sender@sender.com', 'recipient@recipient.com'
            )) {
            throw "Replace the placeholder for $Provider.$name."
        }
    }
}

if ($selected.ContainsKey('Uri')) {
    $uri = [uri]$selected.Uri

    if (-not $uri.IsAbsoluteUri -or
        $uri.Scheme -ne 'https' -or
        $uri.UserInfo -or
        $uri.Fragment) {
        throw 'Use an absolute HTTPS endpoint without user-info or a fragment.'
    }
}

if ($Provider -eq 'SendGrid') {
    foreach ($address in @($selected.From) + @($selected.To)) {
        $null = [System.Net.Mail.MailAddress]::new([string]$address)
    }
}

# Confirm that SMB session enumeration works before launching.
$null = Get-Command Get-SmbSession -ErrorAction Stop
$null = Get-SmbSession -ErrorAction Stop

# Encrypt configuration for the current Windows identity.
# This avoids placing plaintext provider secrets in the child command line.
# It does not protect secrets from administrators, the same identity,
# terminal history, transcription, or inspection of process memory.

Add-Type -AssemblyName System.Security

$config = @{
    Provider        = $Provider
    Settings        = $selected
    DurationHours   = $DurationHours
    PollSeconds     = $PollSeconds
    TimeoutSeconds  = $NotificationTimeoutSeconds
    MaxAttempts     = $NotificationMaxAttempts
    IncludeUsers    = $IncludeUserNames
}

$configBytes = [Text.Encoding]::UTF8.GetBytes(
    ($config | ConvertTo-Json -Depth 8 -Compress)
)

$protectedBytes = [Security.Cryptography.ProtectedData]::Protect(
    $configBytes,
    $null,
    [Security.Cryptography.DataProtectionScope]::CurrentUser
)

$protectedConfig = [Convert]::ToBase64String($protectedBytes)
[Array]::Clear($configBytes, 0, $configBytes.Length)

# ========================= DETACHED WORKER ========================

$worker = @'
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Security

$bytes = [Security.Cryptography.ProtectedData]::Unprotect(
    [Convert]::FromBase64String('__PROTECTED_CONFIG__'),
    $null,
    [Security.Cryptography.DataProtectionScope]::CurrentUser
)

$c = [Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json
[Array]::Clear($bytes, 0, $bytes.Length)

[Net.ServicePointManager]::SecurityProtocol = (
    [Net.ServicePointManager]::SecurityProtocol -bor
    [Net.SecurityProtocolType]::Tls12
)

$script:DeliveryFailures = 0
$script:HostName = $env:COMPUTERNAME

function Send-Notification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [string]$Message
    )

    $s = $c.Settings
    $headers = @{}
    $contentType = 'application/json; charset=utf-8'
    $text = "$Title`n`n$Message"
    $payload = $null

    switch ($c.Provider) {
        'SendGrid' {
            $uri = 'https://api.sendgrid.com/v3/mail/send'
            $headers.Authorization = "Bearer $($s.ApiKey)"
            $payload = @{
                personalizations = @(
                    @{
                        to = @(
                            foreach ($address in @($s.To)) {
                                @{ email = [string]$address }
                            }
                        )
                    }
                )
                from = @{ email = [string]$s.From }
                subject = $Title
                content = @(
                    @{ type = 'text/plain'; value = $Message }
                )
            }
        }

        'Slack' {
            $uri = [string]$s.Uri
            $payload = @{
                text = $text
                mrkdwn = $false
                unfurl_links = $false
                unfurl_media = $false
            }
        }

        'Discord' {
            $uri = [string]$s.Uri

            if ($text.Length -gt 2000) {
                $text = $text.Substring(0, 1950) +
                    "`n[Notification truncated]"
            }

            $payload = @{
                content = $text
                allowed_mentions = @{ parse = @() }
            }
        }

        'Pushover' {
            $uri = 'https://api.pushover.net/1/messages.json'

            if ($Title.Length -gt 250) {
                $Title = $Title.Substring(0, 250)
            }

            if ($Message.Length -gt 1024) {
                $Message = $Message.Substring(0, 975) +
                    "`n[Notification truncated]"
            }

            $payload = @{
                token = [string]$s.Token
                user = [string]$s.User
                title = $Title
                message = $Message
            }
        }

        'Ntfy' {
            $uri = [string]$s.Uri
            $headers.Authorization = "Bearer $($s.Token)"
            $contentType = 'text/plain; charset=utf-8'
        }

        'Webhook' {
            $uri = [string]$s.Uri

            if (-not [string]::IsNullOrWhiteSpace(
                [string]$s.Token
            )) {
                $headers.Authorization = "Bearer $($s.Token)"
            }

            # Generic receiver contract:
            # title, message, host, timeUtc
            $payload = @{
                title = $Title
                message = $Message
                host = $script:HostName
                timeUtc = [datetime]::UtcNow.ToString('o')
            }
        }

        default {
            throw 'Unsupported notification provider.'
        }
    }

    $target = [uri]$uri

    if (-not $target.IsAbsoluteUri -or
        $target.Scheme -ne 'https' -or
        $target.UserInfo -or
        $target.Fragment) {
        throw 'Invalid notification endpoint.'
    }

    $body = if ($c.Provider -eq 'Ntfy') {
        $text
    } else {
        $payload | ConvertTo-Json -Depth 8 -Compress
    }

    $request = @{
        Uri = $target
        Method = 'Post'
        Headers = $headers
        ContentType = $contentType
        Body = [Text.Encoding]::UTF8.GetBytes($body)
        TimeoutSec = [int]$c.TimeoutSeconds
        MaximumRedirection = 0
        ErrorAction = 'Stop'
    }

    for ($attempt = 1; $attempt -le [int]$c.MaxAttempts; $attempt++) {
        try {
            $response = Invoke-RestMethod @request

            if ($c.Provider -eq 'Pushover' -and
                $response.status -ne 1) {
                $script:DeliveryFailures++
                return $false
            }

            return $true
        }
        catch {
            $status = 0
            $retryAfter = 0
            $httpResponse = $_.Exception.Response

            if ($null -ne $httpResponse) {
                try {
                    $status = [int]$httpResponse.StatusCode
                    $rawRetry = [string]$httpResponse.Headers[
                        'Retry-After'
                    ]

                    if (-not [int]::TryParse(
                        $rawRetry, [ref]$retryAfter
                    )) {
                        $retryDate = [datetimeoffset]::MinValue

                        if ([datetimeoffset]::TryParse(
                            $rawRetry, [ref]$retryDate
                        )) {
                            $retryAfter = [int][Math]::Ceiling(
                                ($retryDate -
                                    [datetimeoffset]::UtcNow
                                ).TotalSeconds
                            )
                        }
                    }
                }
                catch {
                    # Keep the original request failure classification.
                }
            }

            $retryable = (
                $status -eq 0 -or
                $status -eq 408 -or
                $status -eq 429 -or
                $status -ge 500
            )

            if (-not $retryable -or
                $attempt -eq [int]$c.MaxAttempts -or
                $retryAfter -gt 60) {
                $script:DeliveryFailures++
                return $false
            }

            $delay = [Math]::Max(
                [Math]::Max(0, $retryAfter),
                [int][Math]::Pow(2, $attempt)
            )

            Start-Sleep -Seconds $delay
        }
    }
}

# Prevent concurrent copies under the same Windows identity.
$created = $false
$mutex = [Threading.Mutex]::new(
    $false, 'Global\SMB1SentinelMonitor', [ref]$created
)
$ownsMutex = $false
$exitCode = 1

try {
    try {
        $ownsMutex = $mutex.WaitOne(0)
    }
    catch [Threading.AbandonedMutexException] {
        $ownsMutex = $true
    }

    if (-not $ownsMutex) {
        $null = Send-Notification `
            -Title "[$script:HostName] SMB1 watch not started" `
            -Message 'Another SMB1 Sentinel monitor is already running.'
        exit 2
    }

    $null = Get-Command Get-SmbSession -ErrorAction Stop
    $null = Get-SmbSession -ErrorAction Stop

    $seen = @{}
    $polls = 0
    $missingDialect = 0
    $failed = $false

    $startUtc = [datetime]::UtcNow
    $endUtc = $startUtc.AddHours([double]$c.DurationHours)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $durationSeconds = [double]$c.DurationHours * 3600

    $started = Send-Notification `
        -Title "[$script:HostName] SMB1 watch started" `
        -Message (
            "Worker PID: $PID`n" +
            "Start UTC: $($startUtc.ToString('o'))`n" +
            "Planned end UTC: $($endUtc.ToString('o'))`n" +
            "Polling interval: $($c.PollSeconds) seconds`n" +
            'Inbound sessions only. Polling can miss short connections.'
        )

    if (-not $started) {
        throw 'Startup notification failed. Monitoring aborted.'
    }

    try {
        while ($timer.Elapsed.TotalSeconds -lt $durationSeconds) {
            $sessions = @(Get-SmbSession -ErrorAction Stop)
            $polls++

            foreach ($session in $sessions) {
                $dialect = [string]$session.Dialect

                if ([string]::IsNullOrWhiteSpace($dialect)) {
                    $missingDialect++
                    continue
                }

                if ($dialect -notmatch '^1(\.|$)') {
                    continue
                }

                $client = [string]$session.ClientComputerName

                if ([string]::IsNullOrWhiteSpace($client)) {
                    $client = "Unknown-client-session-$($session.SessionId)"
                }

                $now = [datetime]::UtcNow.ToString('o')

                if ($seen.ContainsKey($client)) {
                    $seen[$client].LastSeenUtc = $now
                    continue
                }

                $seen[$client] = @{
                    FirstSeenUtc = $now
                    LastSeenUtc = $now
                    AlertAccepted = $false
                }

                $lines = @(
                    "Client: $client"
                    "Dialect: $dialect"
                    "First observed UTC: $now"
                )

                if ($c.IncludeUsers) {
                    $lines += "User: $($session.ClientUserName)"
                }

                $accepted = Send-Notification `
                    -Title "[$script:HostName] SMB1 client observed" `
                    -Message ($lines -join "`n")

                $seen[$client].AlertAccepted = $accepted
            }

            $remaining = $durationSeconds - $timer.Elapsed.TotalSeconds

            if ($remaining -gt 0) {
                $sleepMs = [int][Math]::Min(
                    [double]$c.PollSeconds * 1000,
                    $remaining * 1000
                )

                if ($sleepMs -gt 0) {
                    Start-Sleep -Milliseconds $sleepMs
                }
            }
        }
    }
    catch {
        $failed = $true
        # Avoid forwarding raw exception details or configuration.
    }

    $timer.Stop()

    $status = if ($failed) { 'INCOMPLETE' } else { 'COMPLETED' }

    $summary = @(
        "Status: $status"
        "Start UTC: $($startUtc.ToString('o'))"
        "End UTC: $([datetime]::UtcNow.ToString('o'))"
        "Successful polls: $polls"
        "Unique SMB1 clients observed: $($seen.Count)"
        "Session samples without a dialect: $missingDialect"
        "Prior notification failures: $script:DeliveryFailures"
        ''
        'Zero observations does not prove SMB1 is unused.'
        'Polling and notification delivery introduce observation gaps.'
    )

    if ($failed) {
        $summary += ''
        $summary += 'An SMB query failed. The observation window is incomplete.'
    }

    $summaryAccepted = Send-Notification `
        -Title "[$script:HostName] SMB1 watch $status" `
        -Message ($summary -join "`n")

    # Send client details separately to avoid oversized summary messages.
    foreach ($client in @($seen.Keys | Sort-Object)) {
        $record = $seen[$client]

        $detail =e $detail
    }

    if (-not $failed -and
        $missingDialect -eq 0 -and
        $script:DeliveryFailures -eq 0 -and
        $summaryAccepted) {
        $exitCode = 0
    }
}
catch {
    try {
        $null = Send-Notification `
            -Title "[$script:HostName] SMB1 watch FAILED" `
            -Message (
                'The worker could not initialize or complete monitoring. ' +
                'Do not treat this run as a completed observation window.'
            )
    }
    catch {
        # No working notification channel remains.
    }
}
finally {
    if ($ownsMutex) {
        $mutex.ReleaseMutex()
    }

    $mutex.Dispose()
}

exit $exitCode
'@

# Configuration is inserted as encrypted data, not executable settings.
$worker = $worker.Replace('__PROTECTED_CONFIG__', $protectedConfig)

$encoded = [Convert]::
    [Text.Encoding]::Unicode.GetBytes($worker)
)

# Select native Windows PowerShell even from a 32-bit RMM terminal.
$powerShellPath = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"

if ([Environment]:: -and
    -not [Environment]:: {
    $powerShellPath = "$env:WINDIR\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
}

$commandLine = (
    '"' + $powerShellPath + '"' +
    ' -NoLogo -NoProfile -NonInteractive' +
    ' -WindowStyle Hidden -EncodedCommand ' + $encoded
)

if ($commandLine.Length -gt 30000) {
    throw 'Worker command exceeds the conservative command-line size limit.'
}

$result = Invoke-CimMethod `
    -ClassName Win32_Process `
    -MethodName Create `
    -Arguments @{ CommandLine = $commandLine } `
    -ErrorAction Stop

if ($result.ReturnValue -ne 0) {
    throw "Worker creation failed. Return code: $($result.ReturnValue)"
}

[pscustomobject]@{
    Script        = 'SMB1 Sentinel'
    ProcessId     = $result.ProcessId
    Provider      = $Provider
    DurationHours = $DurationHours
    LaunchStatus  = 'Process created. Confirm the startup notification.'
    StopCommand   = "Stop-Process -Id $($result.ProcessId)"
}
