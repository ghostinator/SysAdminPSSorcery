<# Offline native PowerShell checks. No Pester, credentials, SMB server, or network required.
   Run: .\Test-SMB1-Sentinel.ps1
   HTTP, sleep, and SMB enumeration are replaced with in-process test doubles.
#>
[CmdletBinding()]
param([string]$ScriptPath = (Join-Path $PSScriptRoot 'SMB1-Sentinel.ps1'))
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath $ScriptPath -Raw
$match = [regex]::Match($source, '(?ms)^\$workerSource = @''\r?\n(.*?)^''@')
if (-not $match.Success) { throw 'Embedded worker was not found.' }
$worker = $match.Groups[1].Value
$bootstrapMatch = [regex]::Match($source, '(?ms)^\$bootstrap = @''\r?\n(.*?)^''@')
if (-not $bootstrapMatch.Success) { throw 'Embedded bootstrap was not found.' }
foreach ($entry in @(@{ Name = 'launcher'; Text = $source }, @{ Name = 'worker'; Text = $worker }, @{ Name = 'bootstrap'; Text = $bootstrapMatch.Groups[1].Value })) {
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($entry.Text, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) { throw "Native parser rejected $($entry.Name): $($errors.Message -join ', ')" }
}
$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($worker, [ref]$tokens, [ref]$errors)
# Load functions only. Do not execute the worker's monitoring or mutex code.
$functions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
foreach ($definition in $functions) { . ([scriptblock]::Create($definition.Extent.Text)) }

$script:Passed = 0
function Assert-Check {
    param([bool]$Condition, [string]$Name)
    if (-not $Condition) { throw "FAIL: $Name" }
    $script:Passed++
    Write-Host "PASS: $Name"
}
$script:HostName = 'TEST-SERVER'
$script:DeliveryFailures = 0
$Config = @{ Provider = 'Webhook'; Settings = @{ Uri = 'https://example.invalid/hook'; Token = 'test-token' }; TimeoutSeconds = 1; MaxAttempts = 3 }

$parts = @(Split-NotificationText -Text ('x' * 1601))
Assert-Check ($parts.Count -eq 3 -and ($parts -join '') -eq ('x' * 1601)) 'Long messages split without data loss'
$unicode = ('x' * 799) + [char]::ConvertFromUtf32(128512) + 'tail'
$parts = @(Split-NotificationText -Text $unicode)
Assert-Check (($parts -join '') -ceq $unicode) 'Unicode survives chunking'
Assert-Check (-not [char]::IsHighSurrogate($parts[0][$parts[0].Length - 1])) 'Surrogate pair is not split'

foreach ($provider in @('SendGrid', 'Slack', 'Discord', 'Pushover', 'Ntfy', 'Webhook')) {
    $Config.Provider = $provider
    $Config.Settings = @{ Uri = 'https://example.invalid/hook?wait=false&thread_id=123'; Token = 'test-token'; User = 'test-user'; ApiKey = 'test-key'; From = 'sender@example.invalid'; To = @('one@example.invalid', 'two@example.invalid') }
    $r = New-NotificationRequest -Title 'Test' -Message 'Body' -NotificationId 'stable-id'
    Assert-Check ($r.Method -eq 'Post' -and $r.Uri.Scheme -eq 'https' -and $r.MaximumRedirection -eq 0) "$provider HTTPS POST without redirects"
    $body = [Text.Encoding]::UTF8.GetString($r.Body)
    if ($provider -ne 'Ntfy') { $json = $body | ConvertFrom-Json }
    switch ($provider) {
        'SendGrid' { Assert-Check ($json.personalizations[0].to.Count -eq 2 -and $r.Headers.Authorization -eq 'Bearer test-key') 'SendGrid recipient array and authentication' }
        'Slack' { Assert-Check ($json.mrkdwn -eq $false -and $json.unfurl_links -eq $false) 'Slack formatting and previews disabled' }
        'Discord' { Assert-Check ($r.Uri.Query -match 'wait=true' -and $r.Uri.Query -notmatch 'wait=false' -and $json.allowed_mentions.parse.Count -eq 0) 'Discord confirmation enabled and mentions disabled' }
        'Pushover' { Assert-Check ($json.token -eq 'test-token' -and $json.user -eq 'test-user') 'Pushover credentials in payload' }
        'Ntfy' { Assert-Check ($body -eq "Test`n`nBody" -and $r.Headers.Authorization -eq 'Bearer test-token') 'ntfy plain text and authentication' }
        'Webhook' { Assert-Check ($json.notificationId -eq 'stable-id' -and $json.host -eq 'TEST-SERVER') 'Generic webhook contract' }
    }
    $ack = switch ($provider) {
        'SendGrid' { @{ StatusCode = 202; Content = '' } }
        'Slack' { @{ StatusCode = 200; Content = 'ok' } }
        'Pushover' { @{ StatusCode = 200; Content = '{"status":1}' } }
        'Discord' { @{ StatusCode = 200; Content = '{"id":"123"}' } }
        'Ntfy' { @{ StatusCode = 200; Content = '{"id":"abc"}' } }
        'Webhook' { @{ StatusCode = 204; Content = '' } }
    }
    Assert-Check (Test-ProviderResponse -Response $ack) "$provider acknowledgement accepted"
}
$Config.Provider = 'Pushover'
Assert-Check (-not (Test-ProviderResponse -Response @{ StatusCode = 200; Content = '{"status":0}' })) 'Provider rejection is not success'
Assert-Check (-not (Test-ProviderResponse -Response @{ StatusCode = 200; Content = 'not-json' })) 'Malformed acknowledgement is not success'
$Config.Provider = 'Webhook'
$Config.Settings = @{ Uri = 'http://example.invalid/hook'; Token = '' }
$rejected = $false
try { $null = New-NotificationRequest -Title 'T' -Message 'M' -NotificationId 'id' } catch { $rejected = $true }
Assert-Check $rejected 'HTTP endpoint rejected'

# Mock HTTP and sleep. These tests never contact an endpoint.
$Config.Settings.Uri = 'https://example.invalid/hook'
$script:Calls = 0
$script:SleepCalls = 0
$script:FailUntil = 0
function Invoke-WebRequest {
    [CmdletBinding()]
    param($Uri, $Method, $Headers, $ContentType, $Body, $TimeoutSec, $MaximumRedirection, [switch]$UseBasicParsing)
    $script:Calls++
    if ($script:Calls -le $script:FailUntil) { throw (New-Object Net.WebException('Simulated transport failure')) }
    return @{ StatusCode = 204; Content = '' }
}
function Start-Sleep { param([int]$Seconds, [int]$Milliseconds); $script:SleepCalls++ }
$script:FailUntil = 2
$accepted = Send-Notification -Title 'Test' -Message 'Body'
Assert-Check ($accepted -and $script:Calls -eq 3 -and $script:SleepCalls -eq 2) 'Transient failure retries within budget'
$script:Calls = 0
$script:FailUntil = 10
$script:DeliveryFailures = 0
$accepted = Send-Notification -Title 'Test' -Message 'Body'
Assert-Check (-not $accepted -and $script:Calls -eq 3 -and $script:DeliveryFailures -eq 1) 'Exhausted delivery counted once per part'
$script:Calls = 0
$script:FailUntil = 0
$accepted = Send-Notification -Title 'Test' -Message ('x' * 1601)
Assert-Check ($accepted -and $script:Calls -eq 3) 'All message parts sent'

$script:Sessions = @()
function Get-SmbSession { [CmdletBinding()]param(); return $script:Sessions }
Assert-Check (@(Get-ValidatedSessions).Count -eq 0) 'Empty successful SMB query is valid'
$script:Sessions = @([pscustomobject]@{ Dialect = '1.5'; ClientComputerName = 'TEST-CLIENT' })
Assert-Check (@(Get-ValidatedSessions).Count -eq 1) 'SMB1 dialect sample accepted'
$script:Sessions = @([pscustomobject]@{ Dialect = ''; ClientComputerName = 'TEST-CLIENT' })
$rejected = $false
try { $null = Get-ValidatedSessions } catch { $rejected = $true }
Assert-Check $rejected 'Missing dialect fails instead of appearing clean'
$script:Sessions = @([pscustomobject]@{ Dialect = '3.1.1'; ClientComputerName = '' })
$rejected = $false
try { $null = Get-ValidatedSessions } catch { $rejected = $true }
Assert-Check $rejected 'Missing client identifier fails'
# Exercise the same compression and DPAPI constructors used by detached launch.
if ($env:OS -eq 'Windows_NT') {
    Add-Type -AssemblyName System.Security
    $original = [Text.Encoding]::UTF8.GetBytes('offline launch package test')
    $memory = New-Object IO.MemoryStream
    $zip = New-Object IO.Compression.GZipStream($memory, [IO.Compression.CompressionMode]::Compress, $true)
    try { $zip.Write($original, 0, $original.Length) } finally { $zip.Dispose() }
    $protected = [Security.Cryptography.ProtectedData]::Protect($memory.ToArray(), $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $memory.Dispose()
    $bytes = [Security.Cryptography.ProtectedData]::Unprotect($protected, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $inputStream = New-Object IO.MemoryStream(,$bytes)
    $zip = New-Object IO.Compression.GZipStream($inputStream, [IO.Compression.CompressionMode]::Decompress)
    $reader = New-Object IO.StreamReader($zip, [Text.Encoding]::UTF8)
    try { $decoded = $reader.ReadToEnd() } finally { $reader.Dispose(); $inputStream.Dispose() }
    Assert-Check ($decoded -eq 'offline launch package test') 'Windows DPAPI and compression round trip'
}
$validation = & $ScriptPath -ValidateOnly
Assert-Check ($validation.WorkerChecked -eq $true -and $validation.RuntimeTested -eq $false) 'ValidateOnly parses without starting a worker'
Write-Host "All $script:Passed offline checks passed. Live Windows monitoring and provider delivery are separate integration tests."
