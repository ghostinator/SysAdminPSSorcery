# ============================================================================
# 1. Setup Local Logging & Maintenance
# ============================================================================
$LogDir = "C:\temp"
$LogRetentionDays = 14 

if (-not (Test-Path -Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

Get-ChildItem -Path $LogDir -Include "CadDiscovery_*.txt", "CadReport_*.csv" -File -Recurse | 
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } | 
    Remove-Item -Force -ErrorAction SilentlyContinue

$LogDate = Get-Date -Format "yyyyMMdd_HHmmss"
$TranscriptPath = Join-Path $LogDir "CadDiscovery_$LogDate.txt" 
Start-Transcript -Path $TranscriptPath -NoClobber

try {
    # ============================================================================
    # 2. Configuration & RMM Variable Retrieval
    # ============================================================================
    $SendGridUrl    = "https://api.sendgrid.com/v3/mail/send"
    $SendGridApiKey = $env:SENDGRID_API_KEY 
    $MailFrom       = $env:MAIL_FROM
    $MailTo         = $env:MAIL_TO
    $AppTarget      = $env:CAD_APP_TARGET

    if (-not $SendGridApiKey) {
        Write-Error "SENDGRID_API_KEY variable is missing. Exiting."
        exit
    }
    if (-not $MailFrom -or -not $MailTo) {
        Write-Error "MAIL_FROM or MAIL_TO variables are missing. Exiting."
        exit
    }
    if (-not $AppTarget) {
        $AppTarget = "Autodesk"
    }

    # ============================================================================
    # 3. File Extension Mapping
    # ============================================================================
    $ExtensionMap = @{
        "Autodesk"    = @("*.dwg", "*.dxf", "*.rvt", "*.rfa", "*.ipt", "*.iam", "*.idw", "*.nwd", "*.nwf")
        "SolidWorks"  = @("*.sldprt", "*.sldasm", "*.slddrw")
        "Catia"       = @("*.catpart", "*.catproduct")
        "MicroStation"= @("*.dgn")
        "SketchUp"    = @("*.skp")
        "Rhino"       = @("*.3dm")
        "PTCCreo"     = @("*.prt", "*.asm")
        "Universal"   = @("*.step", "*.iges", "*.stl", "*.obj", "*.jt")
        "All"         = @("*.dwg", "*.dxf", "*.rvt", "*.rfa", "*.ipt", "*.iam", "*.idw", "*.nwd", "*.nwf", 
                          "*.sldprt", "*.sldasm", "*.slddrw", "*.catpart", "*.catproduct", "*.dgn", 
                          "*.skp", "*.3dm", "*.prt", "*.asm", "*.step", "*.iges", "*.stl", "*.obj", "*.jt")
    }

    $Extensions = $ExtensionMap[$AppTarget]
    if ($null -eq $Extensions) {
        $Extensions = $ExtensionMap["Autodesk"]
        $AppTarget = "Autodesk"
    }

    # ============================================================================
    # 4. Smart Drive Discovery & Safe Folder Scanning
    # ============================================================================
    $TargetDrives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3 OR DriveType=2" | 
                    Where-Object { $_.Size -and [uint64]$_.Size -gt 50GB } | 
                    Select-Object -ExpandProperty DeviceID

    $CadFiles = @()
    
    if (-not $TargetDrives) {
        Write-Warning "No drives found matching the size criteria (> 50GB). Exiting."
        exit
    }

    foreach ($Drive in $TargetDrives) {
        Write-Output "Scanning physical drive $Drive\ for $AppTarget files..."
        
        try {
            $RootFiles = Get-ChildItem -Path "$Drive\" -Include $Extensions -File -ErrorAction SilentlyContinue
            if ($RootFiles) { $CadFiles += $RootFiles }
        } catch { }

        $TopLevelDirs = Get-ChildItem -Path "$Drive\" -Directory -ErrorAction SilentlyContinue
        
        foreach ($Dir in $TopLevelDirs) {
            try {
                $FoundFiles = Get-ChildItem -Path $Dir.FullName -Include $Extensions -Recurse -File -ErrorAction SilentlyContinue
                if ($FoundFiles) {
                    $CadFiles += $FoundFiles
                }
            } catch {
                # Silently ignore inaccessible locked folders
            }
        }
    }

    $ReportData = foreach ($File in $CadFiles) {
        [PSCustomObject]@{
            FileName = $File.Name
            Location = $File.DirectoryName
            SizeMB   = [math]::Round($File.Length / 1MB, 2)
            Modified = $File.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        }
    }

    # ============================================================================
    # 5. CSV Generation & SendGrid API Transmission
    # ============================================================================
    if ($ReportData.Count -gt 0) {
        Write-Output "Found $($ReportData.Count) files. Preparing email via SendGrid..."
        
        $CsvFileName = "CadReport_$AppTarget_$LogDate.csv"
        $CsvPath = Join-Path $LogDir $CsvFileName
        $ReportData | Sort-Object SizeMB -Descending | Export-Csv -Path $CsvPath -NoTypeInformation

        $FileBytes = [System.IO.File]::ReadAllBytes($CsvPath)
        $Base64Payload = [Convert]::ToBase64String($FileBytes)

        $Subject = "CAD File Discovery: $($ReportData.Count) $AppTarget files found on $env:COMPUTERNAME"
        $Body = @"
Automated CAD File Discovery Report

Target App: $AppTarget
Total Files Found: $($ReportData.Count)
Drives Scanned: $($TargetDrives -join ', ') (Filtered for drives > 50 GB)

A complete list of the discovered files is attached to this email as a CSV.
"@

        # SendGrid specific JSON structure
        $MessagePayload = @{
            personalizations = @(
                @{
                    to = @( @{ email = $MailTo } )
                    subject = $Subject
                }
            )
            from = @{ email = $MailFrom }
            content = @(
                @{
                    type = "text/plain"
                    value = $Body
                }
            )
            attachments = @(
                @{
                    content = $Base64Payload
                    filename = $CsvFileName
                    type = "text/csv"
                    disposition = "attachment"
                }
            )
        } | ConvertTo-Json -Depth 10

        $Headers = @{
            "Authorization" = "Bearer $SendGridApiKey"
            "Content-Type"  = "application/json"
        }

        # SendGrid returns a 202 Accepted on success, which Invoke-RestMethod handles without throwing
        Invoke-RestMethod -Uri $SendGridUrl -Method Post -Headers $Headers -Body $MessagePayload
        
        Write-Output "Email successfully queued by SendGrid API."

        Remove-Item -Path $CsvPath -Force -ErrorAction SilentlyContinue
    } else {
        Write-Output "No CAD files found for $AppTarget. No email sent."
    }

} catch {
    Write-Error "An error occurred during script execution: $_"
} finally {
    Stop-Transcript
}
