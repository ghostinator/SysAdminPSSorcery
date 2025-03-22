# PowerShell Script to Update Cloudflare Dynamic DNS

## Description

This PowerShell script is designed to automatically update a specified DNS A or AAAA record in your Cloudflare account with your current dynamic external IP address. It performs an initial update when the script is first run and then continuously monitors for changes in your IP. If a change is detected, the script will update the corresponding DNS record in Cloudflare. To prevent excessive API calls, the script will only communicate with Cloudflare once per day if your IP address remains the same.

## Prerequisites

* **PowerShell:** This script requires PowerShell 5.1 or later. It should work on Windows, macOS, and Linux with PowerShell Core.
* **Cloudflare Account:** You need an active Cloudflare account with your domain and the DNS record you want to update already set up.
* **Cloudflare API Token:** You will need to generate a Cloudflare API token with the **"DNS:Edit"** permission for the specific zone (domain) you intend to update. You can create and manage API tokens in your Cloudflare dashboard under "My Profile" -> "API Tokens".
* **Cloudflare Zone ID:** You will need the unique Zone ID for your domain in Cloudflare. This can be found on the "Overview" page of your domain in the Cloudflare dashboard.
* **DNS Record ID:** You need the ID of the specific DNS record (A or AAAA) you want to update. You can find this using the Cloudflare API itself (see the "Finding Your DNS Record ID" section below) or potentially through the Cloudflare dashboard (though the API method is more reliable).

## Configuration

1.  **Download the Script:** Save the PowerShell script (e.g., `Update-Cloudflare-DDNS.ps1`) to a location on your system.

2.  **Edit the Configuration Section:** Open the script in a text editor or PowerShell ISE and modify the variables in the `# Configuration` section with your actual Cloudflare details:
    * `$cloudflareApiToken`: Replace `"YOUR_CLOUDFLARE_API_TOKEN"` with your generated Cloudflare API token.
    * `$cloudflareZoneId`: Replace `"YOUR_CLOUDFLARE_ZONE_ID"` with your Cloudflare Zone ID.
    * `$dnsRecordId`: Replace `"YOUR_DNS_RECORD_ID"` with the ID of the DNS record you want to update.
    * `$dnsRecordName`: Replace `"your.hostname.com"` with the exact name of your DNS record (e.g., `mydomain.com` or `subdomain.mydomain.com`).
    * `$dnsRecordType`: Set this to `"A"` if you are updating an IPv4 address or `"AAAA"` if you are updating an IPv6 address.
    * `$dnsRecordTTL`: Set the desired Time To Live (TTL) for your DNS record in seconds (e.g., `120`, `300`, `3600`).
    * `$dnsRecordProxied`: Set this to `$true` if you want Cloudflare to proxy traffic for this record, or `$false` otherwise.

## Usage

1.  **Open PowerShell:** Open a PowerShell window.

2.  **Navigate to the Script's Directory:** Use the `cd` command to navigate to the directory where you saved the `Update-Cloudflare-DDNS.ps1` file.

3.  **Run the Script:** Execute the script using the following command:
    ```powershell
    .\Update-Cloudflare-DDNS.ps1
    ```

    You might need to adjust PowerShell's execution policy if you haven't already. You can temporarily bypass the execution policy for a single script execution using:
    ```powershell
    powershell -ExecutionPolicy Bypass -File .\Update-Cloudflare-DDNS.ps1
    ```

4.  **Keep the Script Running:** The script will run indefinitely, checking for IP changes in the background. You might want to run this script as a scheduled task or service so it runs automatically.

## Finding Your DNS Record ID

You can use the following PowerShell script (run it once) to find the ID of your DNS record. **Remember to replace the placeholder API token and Zone ID with your actual values and adjust the `$dnsRecordName` and `$dnsRecordType` if needed:**

```powershell
# Replace with your actual values
$cloudflareApiToken = "YOUR_CLOUDFLARE_API_TOKEN"
$cloudflareZoneId = "YOUR_CLOUDFLARE_ZONE_ID"
$dnsRecordName = "your.hostname.com"
$dnsRecordType = "A" # Or "AAAA"

$apiUrl = "[https://api.cloudflare.com/client/v4/zones/$cloudflareZoneId/dns_records](https://api.cloudflare.com/client/v4/zones/$cloudflareZoneId/dns_records)"
$headers = @{
    "Authorization" = "Bearer $cloudflareApiToken"
    "Content-Type"  = "application/json"
}

try {
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    foreach ($record in $response.result) {
        if ($record.name -eq $dnsRecordName -and $record.type -eq $dnsRecordType) {
            Write-Host "DNS Record ID for $($record.name) (Type $($record.type)): $($record.id)"
            # Copy this ID and use it in the main Update-Cloudflare-DDNS.ps1 script
        }
    }
} catch {
    Write-Host "Error retrieving DNS records: $_"
}
