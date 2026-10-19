# SMB1 Sentinel

A no-install PowerShell utility for temporarily watching inbound SMB sessions and notifying you when an SMB1 client is observed.

Choose SendGrid, Slack, Discord, Pushover, ntfy, or your own HTTPS webhook. Configure the script once, then run it from an elevated Windows PowerShell terminal.

**This is a sampling tool, not a complete SMB1 audit. Zero observations does not prove SMB1 is unused.**

## Features

- Default 48-hour observation window with a configurable polling delay.
- Separate worker process for detached operation.
- Foreground mode for troubleshooting.
- Startup, first-client, per-client result, and final summary notifications.
- Provider-neutral `Send-Notification` interface.
- Bounded retries for transient HTTP failures.
- Long-message splitting without dropping message-body content.
- HTTPS-only endpoints with redirects disabled.
- Provider acknowledgement checks, not just a successful HTTP call.
- Current-user Windows DPAPI encryption for the detached launch package.
- No modules, services, scheduled tasks, or software installed by the monitor.
- No automatic SMB configuration changes.

## Requirements

Operational modes require:

- Windows PowerShell **5.1**, launched as administrator.
- A Windows SMB server exposing `Get-SmbSession` and session dialect information.
- Permission to enumerate local SMB sessions and create a process through CIM.
- Outbound HTTPS connectivity to the selected notification provider.
- Valid provider credentials and an approved destination.

PowerShell 7 is not an operational target for this version. The script explicitly checks the runtime before monitoring. No specific Windows release or RMM platform is claimed as integration-tested.

## Quick start

1. Save `SMB1-Sentinel.ps1` locally and review it.
2. Edit the configuration block at the top. Select `$Provider` and fill in that provider's settings.
3. Open an elevated **Windows PowerShell 5.1** terminal on the SMB server.
4. Run the script using your approved execution policy:

   ```powershell
   .\SMB1-Sentinel.ps1
   ```

5. Record the returned process ID and confirm the startup notification arrives.

You can also paste the complete configured script into an elevated terminal. No snippet assembly is required.

**A process ID confirms process creation only. It does not confirm worker initialization or notification delivery.** If no startup notification arrives, treat the run as unconfirmed and use foreground mode to investigate.

The worker is launched separately from the terminal. Whether a management platform permits it to continue after its session closes depends on that platform's process controls. Verify that behavior in your environment. Reboot and forced process termination stop monitoring.

Do not disable endpoint protection or weaken organizational execution controls to run the script.

## Configuration

| Setting | Default | Accepted values |
| --- | --- | --- |
| `Provider` | `SendGrid` | `SendGrid`, `Slack`, `Discord`, `Pushover`, `Ntfy`, `Webhook` |
| `DurationHours` | `48` | Greater than zero through 720. Fractions allow short test runs. |
| `PollSeconds` | `5` | Integer from 1 through 3600. |
| `NotificationTimeoutSeconds` | `15` | Integer from 1 through 120. |
| `NotificationMaxAttempts` | `3` | Integer from 1 through 5, including the first attempt. |
| `IncludeUserNames` | `$false` | Include account names in initial client alerts when `$true`. |

Only the selected provider's settings are used. Unselected provider placeholders can remain unchanged.

The polling delay is not a guaranteed sample interval. Session enumeration, notification requests, and retries add time between samples. The startup notification gives a planned end time. The actual sampling duration starts after that notification is accepted, and final result delivery can extend total process runtime.

### Provider settings

| Provider | Required settings | Behavior |
| --- | --- | --- |
| SendGrid | `ApiKey`, `From`, `To` | Uses the global Mail Send endpoint. `To` accepts an array of addresses. Configure a verified sender and a Mail Send-scoped key. |
| Slack | `Uri` | Incoming webhook. Text formatting, link-name expansion, and previews are disabled. |
| Discord | `Uri` | Incoming webhook. Adds `wait=true` for a message acknowledgement and disables allowed mentions. |
| Pushover | `Token`, `User` | Application token and recipient user or group key. |
| ntfy | `Uri`, `Token` | HTTPS topic endpoint and bearer token. Topic read/write permissions must also be restricted. |
| Generic webhook | `Uri`, optional `Token` | Sends the JSON contract below. A configured token becomes bearer authentication. |

Provider accounts, destination setup, quotas, and any fees are managed outside this script. Regional SendGrid endpoints are not configurable in this version.

All configurable endpoints must be absolute HTTPS URLs without user-info or a fragment. Certificate validation is not disabled. Redirects are not followed.

### Generic webhook contract

```json
{
  "title": "Notification title",
  "message": "Notification body",
  "host": "SERVER-NAME",
  "timeUtc": "ISO-8601 UTC timestamp",
  "notificationId": "logical-message-id-part-index"
}
```

`notificationId` stays the same when retrying the same part. A receiver can use it for deduplication. The script does not require or verify receiver-side deduplication.

This is not a universal Teams webhook payload. A workflow or other receiver must accept this schema and perform any downstream action.

## Run modes

### Detached monitoring

```powershell
.\SMB1-Sentinel.ps1
```

Returns the worker process ID, selected provider, configured duration, and a stop command. It does not wait for the worker's final exit code.

The launcher compresses the worker and configuration, encrypts the package with current-user DPAPI, and passes it to a short bootstrap command. It checks the final command length before process creation. Very large configurations can still exceed that limit and will be rejected.

### Foreground troubleshooting

```powershell
.\SMB1-Sentinel.ps1 -Foreground
```

Runs in the current terminal and reports `WorkerExitCode` on completion. Keep the terminal open. For a short test, temporarily set `DurationHours` to `0.01` in your private configured copy.

Warnings deliberately omit raw HTTP errors because those errors can expose webhook URLs or other secrets. Provider-side delivery records can help diagnose authentication, destination, and quota issues.

### Native syntax validation

```powershell
.\SMB1-Sentinel.ps1 -ValidateOnly
```

Parses the embedded worker with PowerShell's native parser. When run from a file, it also parses the launcher file. It does not require filled-in credentials, query SMB, send notifications, or start a worker.

Syntax validation is not a runtime or integration test. The operational launcher also parses the worker and generated bootstrap before starting a detached process.

## Notification lifecycle

1. **Startup** checks session access and sends a startup notification. Failure to confirm that notification aborts monitoring.
2. **First observation** sends an alert for each new client identifier reporting an SMB1 dialect.
3. **Repeat observations** update the client's last-seen timestamp without sending another initial alert.
4. **Results** send first/last observation details for each client.
5. **Final summary** reports observation status, successful polls, client count, and failed notification parts recorded before the summary.

Clients are deduplicated by `ClientComputerName`, not by a verified inventory identity. Different representations of the same device may appear separately. Multiple users on one client identifier do not generate separate initial alerts.

A failed initial alert is retried within its attempt budget. It is not continuously requeued on later polls. The client remains in memory for the end-of-run result messages.

Long message bodies are split into parts of up to 800 UTF-16 code units without splitting a surrogate pair. Titles are capped before part numbering. This can produce multiple notifications. There is no silent body truncation.

### Understanding status

- **COMPLETED** means the sampling loop reached its duration without a caught query or session-validation failure.
- **INCOMPLETE** means sampling stopped because a query or session-validation step failed.
- **FAILED** means worker initialization or execution failed outside the normal sampling completion path.

Missing or unrecognized dialect information fails the observation window rather than being treated as a clean result. Missing client identifiers also fail validation.

Observation status and delivery status are separate. A completed sampling loop can still have notification failures. The final summary includes failures from preceding client-result messages, but cannot report its own delivery failure.

### Worker result codes

| Code | Meaning |
| --- | --- |
| `0` | Sampling completed and every notification part received the expected provider acknowledgement. |
| `1` | Initialization, monitoring, or notification delivery failed. |
| `2` | Another worker owns the monitor mutex. |

Detached mode uses these as process exit codes. Foreground mode returns the code in a result object instead of closing the interactive shell. The launcher's CIM return value is separate from the worker's eventual result.

Provider acknowledgement does not guarantee inbox delivery, a phone alert, or that a person read the notification.

## Check or stop a worker

Use the process ID returned by the launcher:

```powershell
Get-Process -Id 1234 | Select-Object Id, ProcessName, StartTime
Stop-Process -Id 1234
```

Replace `1234` with your worker's ID. Verify the process before stopping it because IDs can be reused. Forced termination does not guarantee a final notification and loses in-memory findings.

A named mutex discourages overlapping runs. It is not a security boundary or a guarantee across identities with different object permissions. An access-denied mutex error is not proof that another healthy monitor exists.

## Security and privacy

- Keep real credentials and configured private copies out of public source control.
- Treat webhook URLs as secrets.
- Use narrowly scoped credentials and approved notification destinations.
- Rotate credentials exposed in commits, screenshots, transcripts, or logs.
- Protect ntfy topic read access as well as publishing access. A token alone does not make a public topic private.
- Leave `IncludeUserNames` disabled unless account names are needed and approved for the destination.

DPAPI protects the detached launch package from being passed as plaintext credentials in the child command line. It does not protect secrets in the original configured script, clipboard, terminal history, transcripts, RMM logs, or process memory. The same identity and privileged administrators remain relevant trust boundaries. The worker must be able to decrypt the launcher's current-user data.

Notifications contain server names, client identifiers, dialects, and timestamps. Account names are added only when enabled. Disabling account names does not anonymize the remaining information.

Use only on systems you are authorized to administer. The script observes sessions and does not decide whether removing SMB1 is safe.

## Limitations

- Watches current inbound sessions on the local SMB server, not outbound connections or every server in the environment.
- Short connections can occur entirely between samples.
- Notification requests and retries block sampling.
- No packet capture or server-side SMB auditing is enabled.
- No local evidence files, persistent queue, heartbeat, or automatic restart.
- Reboot, process termination, or management-platform controls can stop the worker.
- Findings are held in memory and can be lost.
- Timeouts and retries can produce duplicate messages.
- A large client population can generate many notifications and encounter provider quotas.
- Retry delays are bounded. A `Retry-After` value over 60 seconds causes that part to fail rather than blocking sampling for a longer wait.
- HTTP 408, 429, 5xx, and transport failures are retry candidates. Permanent HTTP errors are not retried.
- Invalid successful-response acknowledgements are treated as failures, not as confirmed delivery.

For a more complete dependency assessment, consider retained server-side SMB1 audit evidence alongside sampling. That configuration is outside this utility's scope.

## Testing

The package includes `Test-SMB1-Sentinel.ps1`. It uses PowerShell's native parser and in-process test doubles. It requires no Pester installation, provider credentials, or network access.

```powershell
.\Test-SMB1-Sentinel.ps1
```

It checks launcher/worker/bootstrap parsing, request construction, provider acknowledgements, message splitting, retry behavior, and session validation. On Windows it also exercises a DPAPI/compression round trip. It does not test real SMB traffic, CIM process creation, terminal closure, or live provider delivery.

The included repository workflow runs native checks with Windows PowerShell on a Windows runner when enabled in a repository. A workflow file is not evidence of a successful run. Check its actual results before making release claims.

### Validation status of this package

- Launcher, worker, bootstrap, and offline-test source passed an independent Tree-sitter PowerShell grammar check during preparation.
- Native Windows PowerShell execution and the supplied offline test suite were **not run** in the preparation environment because that runtime was unavailable.
- Live Windows SMB monitoring, detached launch, DPAPI identity behavior, and notification delivery remain integration tests for the target environment.

The earlier malformed expressions have been replaced. No claim of end-to-end production certification is made. Record actual test results and tested environments before advertising a verified release.

## Troubleshooting

| Symptom | Action |
| --- | --- |
| Placeholder error | Fill in only the selected provider's settings. |
| Runtime error | Use elevated Windows PowerShell 5.1, not a PowerShell 7 tab. |
| Session query fails | Check local SMB command availability and enumeration permissions. |
| PID returned but no startup notification | Treat the run as unconfirmed. Try foreground mode and inspect provider records and endpoint controls. |
| Incomplete result | Investigate SMB enumeration or missing/unrecognized session fields. Do not treat the window as completed. |
| Duplicate-run message | Verify the existing process before stopping it or starting another worker. |
| Delivery failures | Check credentials, destination permissions, connectivity, and provider quotas. |
| No SMB1 observations | Review server selection and observation timing. Zero observations is not proof of no usage. |

## Contributing

Include Windows and PowerShell versions, selected provider, expected behavior, and sanitized reproduction details. Do not post secrets or identifiable environment information in public issues.

Keep monitoring logic independent of provider-specific request construction. Add tests for new adapters and document their acknowledgement rules and limitations.

## References

- [Microsoft Get-SmbSession documentation](https://learn.microsoft.com/en-us/powershell/module/smbshare/get-smbsession?view=windowsserver2025-ps)
- [SendGrid Mail Send API](https://www.twilio.com/docs/sendgrid/api-reference/mail-send/mail-send)
- [Slack incoming webhooks](https://docs.slack.dev/messaging/sending-messages-using-incoming-webhooks/)
- [Discord webhook API](https://discord.com/developers/docs/resources/webhook)
- [Pushover message API](https://pushover.net/api)
- [ntfy publishing documentation](https://docs.ntfy.sh/publish/)

## License

A license has not been selected for this package. The repository owner should add a `LICENSE` file before inviting public reuse.
