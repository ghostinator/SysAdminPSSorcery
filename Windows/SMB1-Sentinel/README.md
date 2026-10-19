# SMB1 Sentinel

A lightweight PowerShell watcher for discovering inbound SMB1 sessions on a Windows SMB server, with notifications through your preferred provider.

No software installation. No scheduled task. No automatic SMB configuration changes.

> **Release status: pre-release, untested.** The script supplied alongside this initial documentation contains malformed PowerShell expressions and must be corrected and validated before use or publication as a working release. This README describes the intended behavior of that implementation, not a verified release. No Windows compatibility or provider integration testing has been completed.

## What it does

SMB1 Sentinel periodically queries `Get-SmbSession`, identifies sessions whose reported dialect matches SMB1, and sends notifications for newly observed clients.

The launcher starts a separate Windows PowerShell process so monitoring is not tied to the interactive terminal. Confirm this behavior in your remote management environment before relying on it.

The intended workflow includes:

- A startup notification with the worker process ID and observation window.
- One initial alert per unique client identifier during a run.
- A completion or failure summary.
- Separate end-of-run client results with first and last observation timestamps.
- Bounded retries for notification delivery.
- A named mutex to prevent overlapping workers that can access the same mutex.
- Configuration encrypted with Windows DPAPI for transfer to the worker process.

## What it does not do

- Disable, enable, or remediate SMB1.
- Enable SMB auditing or change server configuration.
- Capture packets or continuously record network traffic.
- Detect outbound SMB1 connections initiated by the monitored server.
- Discover SMB1 activity on other servers automatically.
- Install a service, scheduled task, module, or third-party dependency.
- Resume after a reboot or process termination.
- Save findings, logs, or pending notifications to disk.
- Prove that SMB1 is unused.

**No observations means no SMB1 sessions were observed during successful samples. It does not mean no SMB1 traffic occurred.** Short-lived sessions can fall between polls. Notification requests and retries also pause polling.

## Requirements

The implementation targets:

- Windows with Windows PowerShell 5.1.
- An elevated administrator or appropriately privileged management session.
- The `Get-SmbSession` command and permission to enumerate local SMB sessions.
- Access to the configured HTTPS notification endpoint.
- Valid credentials and any provider-specific sender or endpoint configuration.
- Permission to start a worker through `Win32_Process.Create`.
- A worker identity able to decrypt the launcher's current-user DPAPI configuration.

The target server and notification provider must be tested before operational use. This is not a macOS, Linux, or PowerShell 7 compatibility claim.

## Quick start

**Complete the release validation checklist below before following these steps.**

1. Save the corrected script as `SMB1-Sentinel.ps1`.
2. Review the code and select a notification provider in the configuration section.
3. Replace only the selected provider's placeholder settings.
4. Open an elevated Windows PowerShell terminal on the SMB server.
5. Run the script under your organization's approved script execution policy:

   ```powershell
   .\SMB1-Sentinel.ps1
   ```

   Alternatively, paste the complete corrected script into an elevated terminal.

6. Record the returned worker process ID and confirm receipt of the startup notification.

A successful process creation response is **not** confirmation that the worker initialized, can query SMB sessions, or can send notifications. If the startup notification does not arrive, treat the run as unconfirmed.

Do not weaken execution policy, endpoint protection, or application control to run this tool. Use your approved administrative deployment process.

## Configuration

| Setting | Default | Purpose |
| --- | --- | --- |
| `Provider` | `SendGrid` | Selects the notification adapter. |
| `DurationHours` | `48` | Observation duration. Allowed range is greater than zero through 720 hours. |
| `PollSeconds` | `5` | Delay between polling cycles. Allowed range is 1 through 3600 seconds. |
| `NotificationTimeoutSeconds` | `15` | Configured timeout for each HTTP request. Allowed range is 1 through 120 seconds. |
| `NotificationMaxAttempts` | `3` | Maximum attempts per notification. Allowed range is 1 through 5. |
| `IncludeUserNames` | `$false` | Includes account names in initial client alerts when enabled. |

The actual interval between samples includes session enumeration, notification requests, retries, and the configured polling delay. The request timeout is not a guarantee of the entire worker's maximum runtime.

### Supported notification adapters

These adapters are implemented in the supplied design but have not been integration-tested.

| Provider | Required settings | Notes |
| --- | --- | --- |
| SendGrid | `ApiKey`, `From`, `To` | Configure an authorized sender and a narrowly scoped Mail Send key. `To` accepts an array. |
| Slack | `Uri` | Use an incoming webhook for the intended destination. |
| Discord | `Uri` | Use a webhook for the intended channel. The adapter disables allowed mentions. |
| Pushover | `Token`, `User` | Uses an application token and recipient user key. |
| ntfy | `Uri`, `Token` | Use a protected topic with appropriate read and write permissions. |
| Generic webhook | `Uri`, optional `Token` | Receiver must accept the JSON contract below. A token is sent as bearer authentication when configured. |

All configurable endpoints must use HTTPS. Redirects are disabled. Do not embed credentials in the URL user-info field.

The generic webhook is **not** a universal Microsoft Teams adapter. A receiving workflow must be configured to accept this schema and perform any downstream action.

### Generic webhook contract

```json
{
  "title": "Notification title",
  "message": "Notification body",
  "host": "SERVER-NAME",
  "timeUtc": "ISO-8601 UTC timestamp"
}
```

### Adding a provider

Provider-specific request construction belongs in `Send-Notification`. Monitoring calls supply only a title and message.

To extend the implementation:

1. Add the provider name to launcher validation.
2. Add its settings and required-field checks.
3. Add its request adapter to `Send-Notification`.
4. Validate its payload, authentication, response handling, and message limits.
5. Test successful delivery and failure paths before documenting it as supported.

## Notifications and results

### Startup

The worker checks SMB session access and attempts a startup notification. If that notification fails, the worker aborts rather than continuing without a confirmed notification channel.

### Client observation

A new client identifier generates an initial alert containing the client, dialect, and first observation timestamp. Account names are included only when `IncludeUserNames` is enabled.

Deduplication uses `ClientComputerName`, not a verified device inventory identity. The same device represented by different identifiers may appear more than once. Multiple accounts using the same client identifier do not generate separate initial alerts.

A failed initial alert is retried within the configured attempt budget. It is not continuously requeued during later polls. The client remains in memory for the end-of-run result messages.

### End of run

The worker attempts a summary, followed by a separate result message for each observed client. This can produce many notifications in a busy environment.

- `COMPLETED` means the polling loop reached its intended duration without a caught query failure.
- `INCOMPLETE` means a query failed and monitoring stopped early.
- Samples with missing dialect information are counted separately.
- The summary reports notification failures recorded before that summary was sent.

`COMPLETED` is not a clean bill of health. Review missing dialect samples and delivery failures too. Failures sending the summary or subsequent result messages cannot appear retroactively in that summary.

### Worker exit codes

| Code | Meaning in the implementation |
| --- | --- |
| `0` | Polling completed, no missing-dialect samples were recorded, and all notification calls were accepted without a recorded final delivery failure. |
| `1` | Initialization, monitoring, observation completeness, or notification delivery encountered a failure. |
| `2` | The worker detected another monitor through the named mutex and did not start polling. |

The launcher does not wait for or collect the worker's final exit code. Its process creation return value is separate from these worker exit codes.

## Checking and stopping a run

Use the process ID returned by the launcher:

```powershell
Get-Process -Id 1234
```

To stop it manually:

```powershell
Stop-Process -Id 1234
```

Replace `1234` with the recorded worker process ID. Verify the process identity and start time before stopping it because process IDs can be reused.

Forced termination does not guarantee a final notification. Findings are held in memory and may be lost.

## Security and privacy

### Credentials

The launcher encrypts serialized configuration with current-user Windows DPAPI before including it in the encoded worker command. Base64 encoding itself provides no confidentiality.

DPAPI configuration transfer does not protect against:

- Administrators or processes able to inspect the worker's memory.
- Access under the same Windows identity.
- Secrets retained in the original script, clipboard, terminal history, transcripts, or remote management logs.
- Credentials accidentally committed to source control.

Use temporary, narrowly scoped credentials where practical. Keep real secrets out of the public repository. Revoke or rotate any credential exposed in a commit, screenshot, issue, or log.

The worker must run under an identity able to decrypt the launcher's current-user DPAPI data. Validate the actual identity behavior of your management platform.

### Data leaving the server

Notifications can contain:

- Server name.
- Client identifiers, which may include IP addresses.
- SMB dialect.
- Observation timestamps.
- Account names when explicitly enabled.

Disabling account names does not anonymize the remaining information. Use only approved providers and destinations. Restrict access to channels, topics, mailboxes, and webhook receivers.

### Operational boundaries

Use only on systems you own or are authorized to administer. The tool observes sessions. It does not establish whether SMB1 can safely be removed and must not be treated as an automatic remediation decision.

## Reliability limitations

- Short connections can be missed between samples.
- Notification delivery blocks the polling loop.
- HTTP acceptance does not guarantee a person received the notification.
- A timed-out request may already have been accepted, so retries can duplicate messages.
- Only bounded, in-memory retry behavior is provided. There is no durable delivery queue.
- The worker can stop because of reboot, process termination, endpoint controls, or management-platform behavior.
- The configured observation duration is not a hard deadline for final notification delivery.
- A named mutex is not a security boundary or a universal cross-identity coordination guarantee.
- Findings are not preserved locally.
- Large client populations can create a notification burst at completion.
- Provider message truncation can omit details.

For a more complete dependency assessment, consider server-side SMB1 auditing and retained event evidence in addition to this sampling tool. That is outside this script's current scope.

## Troubleshooting

| Symptom | Checks |
| --- | --- |
| Launcher rejects configuration | Replace the selected provider's placeholders and check numeric ranges. |
| SMB access check fails | Confirm elevation, command availability, and local session-enumeration permissions. |
| Process starts but no startup notification arrives | Treat the run as unconfirmed. Check provider credentials, destination permissions, outbound connectivity, worker identity, and endpoint controls. |
| Another monitor is reported | Review the existing worker before stopping it or launching another copy. |
| No clients are observed | Confirm the correct server and observation window. Do not interpret zero samples as proof that SMB1 is unused. |
| Summary is incomplete | A query failed. Do not use the run as a completed observation window. |
| Notifications are missing or duplicated | Check provider delivery records and rate limits. Review the bounded-retry and no-durable-queue limitations. |
| Script fails to parse | The initial supplied version has malformed expressions. Correct and validate the source before use. |

Raw HTTP errors are deliberately not forwarded because they may contain sensitive endpoint details. The current implementation does not provide a local diagnostic log.

## Release validation checklist

Before publishing a working release:

- [ ] Correct all malformed PowerShell expressions in the supplied source.
- [ ] Parse the complete script with Windows PowerShell 5.1 without executing it.
- [ ] Parse the generated worker separately. It is stored as text and is not fully validated by parsing the launcher alone.
- [ ] Test launch and DPAPI decryption under the intended administrative identity.
- [ ] Test terminal closure in each claimed remote management environment.
- [ ] Test every provider advertised as supported with non-production credentials.
- [ ] Verify startup failure, authentication failure, rate limiting, timeout, and endpoint-unavailable behavior.
- [ ] Verify SMB query failure produces an incomplete result.
- [ ] Validate dialect reporting and detection against authorized, controlled test evidence. Do not enable SMB1 on a production server just to test this script.
- [ ] Test duplicate-run prevention and manual termination.
- [ ] Confirm that public source, examples, screenshots, and repository history contain no real secrets or client information.
- [ ] Document tested Windows versions and any unsupported environments.
- [ ] Choose a license and add the corresponding `LICENSE` file.

## Contributing

Bug reports and improvements are welcome. Include the Windows version, PowerShell version, selected provider, expected behavior, and sanitized reproduction details.

Do not include API keys, webhook URLs, access tokens, account names, client identifiers, or other sensitive environment information in public issues.

Useful improvements include durable result storage, a persistent notification queue, non-blocking delivery, event-based observation, and automated tests.

## License

No license has been selected in this README. Add an explicit `LICENSE` file before inviting public reuse. This document does not grant a software license.
