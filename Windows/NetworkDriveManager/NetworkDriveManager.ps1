<#
    Enhanced Network Drive Manager
    Features: Logging, Export, Credential Management, Advanced Diagnostics, Session Awareness
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$script:ShowAllUsers = $script:IsAdmin
$script:MappedDrives = @()
$script:SelectedDriveIndex = -1
$script:LogPath = "$env:ProgramData\NetworkDriveManager"
$script:LogFile = "$script:LogPath\NetworkDriveManager.log"

# Ensure log directory exists
if (-not (Test-Path $script:LogPath)) {
    New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null
}

function New-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enhanced Network Drive Manager"
    $form.Size = New-Object System.Drawing.Size(1000, 750)
    $form.MaximizeBox = $false
    $form.StartPosition = "CenterScreen"
    
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(980, 700)
    $panel.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($panel)

    # Title
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Enhanced Network Drive Manager"
    $lblTitle.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(10, 10)
    $lblTitle.AutoSize = $true
    $panel.Controls.Add($lblTitle)

    # Admin status and session warning
    $lblAdmin = New-Object System.Windows.Forms.Label
    $lblAdmin.AutoSize = $true
    $lblAdmin.Location = New-Object System.Drawing.Point(10, 35)
    if ($script:IsAdmin) {
        $lblAdmin.Text = "(Admin mode – all local users' drives shown)"
        $lblAdmin.ForeColor = "Orange"
    } else {
        $lblAdmin.Text = "(User mode – only current user's drives shown)"
        $lblAdmin.ForeColor = "Green"
    }
    $panel.Controls.Add($lblAdmin)

    # Session context warning
    $script:lblSessionWarning = New-Object System.Windows.Forms.Label
    $script:lblSessionWarning.AutoSize = $true
    $script:lblSessionWarning.Location = New-Object System.Drawing.Point(10, 55)
    $script:lblSessionWarning.ForeColor = "Red"
    $script:lblSessionWarning.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    if ($script:IsAdmin) {
        $script:lblSessionWarning.Text = "⚠ Warning: Running as Administrator - mapped drives may not be visible in Explorer!"
    } else {
        $script:lblSessionWarning.Text = ""
    }
    $panel.Controls.Add($script:lblSessionWarning)

    # Tech tip
    $lblTip = New-Object System.Windows.Forms.Label
    $lblTip.ForeColor = "Blue"
    $lblTip.AutoSize = $true
    $lblTip.Location = New-Object System.Drawing.Point(10, 75)
    $lblTip.Text = 'Tech Tip: Run "gpresult /h c:\temp\gpreport.html" to see if drives are deployed by GPO.'
    $panel.Controls.Add($lblTip)

    # Drive list
    $script:lbDrives = New-Object System.Windows.Forms.ListBox
    $script:lbDrives.Size = New-Object System.Drawing.Size(950, 180)
    $script:lbDrives.Location = New-Object System.Drawing.Point(10, 100)
    $script:lbDrives.SelectionMode = "One"
    $script:lbDrives.add_SelectedIndexChanged({
        if ($script:lbDrives.SelectedIndex -ge 0) {
            $script:SelectedDriveIndex = $script:lbDrives.SelectedIndex
            Update-DetailsBox
        }
    })
    $panel.Controls.Add($script:lbDrives)

    # Main buttons row
    $yBtn = 295
    $btnRefresh = New-Button "Refresh Drives" 10 $yBtn { Refresh-DriveList }
    $btnNew = New-Button "Create New" 130 $yBtn { Show-NewDialog }
    $btnEdit = New-Button "Edit Selected" 250 $yBtn { Show-EditDialog }
    $btnDelete = New-Button "Delete Selected" 370 $yBtn { Remove-Selected }
    $btnTrouble = New-Button "Troubleshoot" 490 $yBtn { Show-Troubleshoot }
    $btnTestAll = New-Button "Test All Drives" 610 $yBtn { Test-AllDrives }
    $panel.Controls.AddRange(@($btnRefresh, $btnNew, $btnEdit, $btnDelete, $btnTrouble, $btnTestAll))

    # Advanced buttons row
    $yBtn2 = 335
    $btnCredentials = New-Button "Manage Credentials" 10 $yBtn2 { Show-CredentialManager }
    $btnExportDrives = New-Button "Export Drives" 130 $yBtn2 { Export-DrivesToCSV }
    $btnExportLog = New-Button "Export Log" 250 $yBtn2 { Export-LogToFile }
    $btnOpenExplorer = New-Button "Open in Explorer" 370 $yBtn2 { Open-SelectedInExplorer }
    $btnCopyPath = New-Button "Copy UNC Path" 490 $yBtn2 { Copy-SelectedUNCPath }
    $panel.Controls.AddRange(@($btnCredentials, $btnExportDrives, $btnExportLog, $btnOpenExplorer, $btnCopyPath))

    # Details section
    $lblDetails = New-Label "Drive Details:" 10 380
    $lblDetails.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $panel.Controls.Add($lblDetails)

    $script:txtDetails = New-TextBox 10 405 950 120 $true
    $panel.Controls.Add($script:txtDetails)

    # Status bar
    $script:lblStatus = New-Label "Ready" 10 540
    $script:lblStatus.Size = New-Object System.Drawing.Size(950, 22)
    $script:lblStatus.BorderStyle = "Fixed3D"
    $panel.Controls.Add($script:lblStatus)

    # Activity log
    $lblLog = New-Label "Activity Log:" 10 570
    $lblLog.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $panel.Controls.Add($lblLog)

    $script:txtLog = New-TextBox 10 595 950 80 $true
    $panel.Controls.Add($script:txtLog)

    return $form
}

# GUI helper functions
function New-Button($text, $x, $y, $handler) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Size = New-Object System.Drawing.Size(115, 30)
    $b.Location = New-Object System.Drawing.Point($x, $y)
    $b.Add_Click($handler)
    return $b
}

function New-Label($txt, $x, $y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $txt
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.AutoSize = $true
    return $l
}

function New-TextBox($x, $y, $w, $h, $multi) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y)
    $t.Size = New-Object System.Drawing.Size($w, $h)
    $t.Multiline = $multi
    $t.ScrollBars = "Vertical"
    $t.ReadOnly = $true
    $t.Font = New-Object System.Drawing.Font("Consolas", 9)
    return $t
}

function Safe-String($v) {
    if ($null -eq $v -or $v -eq "") { return "N/A" }
    return $v.ToString()
}

# Enhanced logging with file output
function Write-Log($msg, $level = "INFO") {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$level] $msg"
    
    # GUI log
    try {
        if ($script:txtLog -and $script:txtLog.Handle -and -not $script:txtLog.IsDisposed) {
            $script:txtLog.AppendText("$logEntry`r`n")
            $script:txtLog.ScrollToCaret()
        }
    } catch {}
    
    # File log
    try {
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    } catch {}
    
    # Console
    Write-Host $logEntry
}

# Drive enumeration functions
function Get-CurrentUserDrives {
    $drives = @()
    try {
        # PSDrive enumeration
        $ps = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.DisplayRoot }
        foreach ($d in $ps) {
            $drives += [pscustomobject]@{
                User = Safe-String $env:USERNAME
                Letter = Safe-String $d.Name
                Remote = Safe-String $d.DisplayRoot
                Root = Safe-String $d.Root
                Provider = "PSDrive"
                Status = "Connected"
                Persistent = "Session"
                LastTest = "Never"
                TestResult = "Not tested"
            }
        }
        
        # WMI enumeration
        $wmiDrives = Get-WmiObject Win32_MappedLogicalDisk -ErrorAction SilentlyContinue
        if ($wmiDrives) {
            foreach ($w in $wmiDrives) {
                $l = $w.Name.Replace(":", "")
                if (-not ($drives | Where-Object { $_.Letter -eq $l })) {
                    $drives += [pscustomobject]@{
                        User = Safe-String $env:USERNAME
                        Letter = Safe-String $l
                        Remote = Safe-String $w.ProviderName
                        Root = Safe-String $w.Name
                        Provider = "WMI"
                        Status = "Connected"
                        Persistent = "Session"
                        LastTest = "Never"
                        TestResult = "Not tested"
                    }
                }
            }
        }
        
        # Net use enumeration
        try {
            $netOutput = & net use 2>$null
            foreach ($line in $netOutput) {
                if ($line -match "^\s*(\w+):\s+(.+?)\s+Microsoft Windows Network") {
                    $letter = $matches[1]
                    $path = $matches[2].Trim()
                    if (-not ($drives | Where-Object { $_.Letter -eq $letter })) {
                        $drives += [pscustomobject]@{
                            User = Safe-String $env:USERNAME
                            Letter = Safe-String $letter
                            Remote = Safe-String $path
                            Root = "${letter}:"
                            Provider = "NetUse"
                            Status = "Connected"
                            Persistent = "Session"
                            LastTest = "Never"
                            TestResult = "Not tested"
                        }
                    }
                }
            }
        } catch {}
    } catch {
        Write-Log "Error in Get-CurrentUserDrives: $($_.Exception.Message)" "ERROR"
    }
    return $drives
}

function Get-RegistryDrivesAllProfiles {
    $result = @()
    try {
        $profiles = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^S-1-5-21' }
        
        foreach ($p in $profiles) {
            try {
                $sid = $p.PSChildName
                $path = $p.GetValue('ProfileImagePath')
                if (-not $path) { continue }
                
                $userRoot = "Registry::HKEY_USERS\$sid"
                $networkKey = "$userRoot\Network"
                $loadedHive = $false
                
                # Check if user hive is loaded
                if (-not (Test-Path $userRoot)) {
                    $ntuserPath = Join-Path $path 'NTUSER.DAT'
                    if (Test-Path $ntuserPath) {
                        try {
                            $tempHive = "TempHive_$sid"
                            & reg load "HKU\$tempHive" $ntuserPath 2>$null | Out-Null
                            if ($LASTEXITCODE -eq 0) {
                                $networkKey = "Registry::HKEY_USERS\$tempHive\Network"
                                $loadedHive = $true
                            }
                        } catch { continue }
                    } else { continue }
                }
                
                if (Test-Path $networkKey) {
                    $drives = Get-ChildItem $networkKey -ErrorAction SilentlyContinue
                    foreach ($sub in $drives) {
                        try {
                            $dl = $sub.PSChildName
                            $props = Get-ItemProperty $sub.PSPath -ErrorAction SilentlyContinue
                            if ($props -and $props.RemotePath) {
                                $result += [pscustomobject]@{
                                    User = Safe-String (Split-Path $path -Leaf)
                                    Letter = Safe-String $dl
                                    Remote = Safe-String $props.RemotePath
                                    Root = "${dl}:"
                                    Provider = "Registry"
                                    Status = "Disconnected"
                                    Persistent = "Persistent"
                                    LastTest = "Never"
                                    TestResult = "Not tested"
                                }
                            }
                        } catch {}
                    }
                }
                
                # Cleanup loaded hive
                if ($loadedHive) {
                    try {
                        & reg unload "HKU\TempHive_$sid" 2>$null | Out-Null
                    } catch {}
                }
            } catch {
                Write-Log "Error processing profile $($p.PSChildName): $($_.Exception.Message)" "ERROR"
            }
        }
    } catch {
        Write-Log "Error in Get-RegistryDrivesAllProfiles: $($_.Exception.Message)" "ERROR"
    }
    return $result
}

function Get-AllDrives {
    $all = @()
    try {
        $all += Get-CurrentUserDrives
        if ($script:ShowAllUsers) {
            $all += Get-RegistryDrivesAllProfiles
        }
        
        # Remove duplicates
        $unique = @()
        foreach ($drive in $all) {
            $exists = $false
            foreach ($existing in $unique) {
                if ($existing.User -eq $drive.User -and 
                    $existing.Letter -eq $drive.Letter -and 
                    $existing.Remote -eq $drive.Remote) {
                    $exists = $true
                    break
                }
            }
            if (-not $exists) {
                $unique += $drive
            }
        }
        return $unique | Sort-Object User, Letter
    } catch {
        Write-Log "Error in Get-AllDrives: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Refresh-DriveList {
    try {
        Write-Log "Refreshing drive list..."
        $script:MappedDrives = @()
        $script:MappedDrives = Get-AllDrives
        $script:lbDrives.Items.Clear()
        
        foreach ($d in $script:MappedDrives) {
            try {
                $user = Safe-String $d.User
                $letter = Safe-String $d.Letter
                $remote = Safe-String $d.Remote
                $provider = Safe-String $d.Provider
                $persistent = Safe-String $d.Persistent
                
                $displayText = "$user`: $letter`: → $remote [$provider - $persistent]"
                $script:lbDrives.Items.Add($displayText)
            } catch {
                Write-Log "Error formatting drive entry: $($_.Exception.Message)" "ERROR"
                $script:lbDrives.Items.Add("Error displaying drive - see log")
            }
        }
        
        $script:lblStatus.Text = "Loaded $($script:MappedDrives.Count) drive records"
        Write-Log "Drive list refreshed successfully - found $($script:MappedDrives.Count) drives"
    } catch {
        Write-Log "Error refreshing drive list: $($_.Exception.Message)" "ERROR"
        $script:lblStatus.Text = "Error loading drives"
    }
}

function Update-DetailsBox {
    try {
        if ($script:SelectedDriveIndex -lt 0 -or $script:SelectedDriveIndex -ge $script:MappedDrives.Count) {
            $script:txtDetails.Text = "No drive selected"
            return
        }
        
        $d = $script:MappedDrives[$script:SelectedDriveIndex]
        $user = Safe-String $d.User
        $letter = Safe-String $d.Letter
        $remote = Safe-String $d.Remote
        $provider = Safe-String $d.Provider
        $status = Safe-String $d.Status
        $persistent = Safe-String $d.Persistent
        $lastTest = Safe-String $d.LastTest
        $testResult = Safe-String $d.TestResult
        
        # Quick accessibility test
        $accessible = "Unknown"
        try {
            if ($letter -ne "N/A" -and $letter -ne "") {
                $testPath = "${letter}:"
                $accessible = if (Test-Path $testPath) { "Yes" } else { "No" }
            }
        } catch {
            $accessible = "Error testing"
        }
        
        $script:txtDetails.Text = @"
User           : $user
Drive Letter   : ${letter}:
Remote Path    : $remote
Provider       : $provider
Status         : $status
Persistent     : $persistent
Currently      : $accessible
Last Test      : $lastTest
Test Result    : $testResult
"@
    } catch {
        $script:txtDetails.Text = "Error displaying drive details: $($_.Exception.Message)"
        Write-Log "Error updating details box: $($_.Exception.Message)" "ERROR"
    }
}

# Advanced diagnostic functions
function Test-DriveHealth($drive) {
    $results = ""
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    try {
        $path = "$($drive.Letter):"
        $server = ""
        
        # Extract server name from UNC path
        if ($drive.Remote -match "^\\\\([^\\]+)") {
            $server = $matches[1]
        }
        
        # Test 1: Ping server
        if ($server) {
            $results += "Ping $server`: "
            try {
                $pingResult = Test-Connection $server -Count 1 -Quiet -ErrorAction Stop
                $results += if ($pingResult) { "OK" } else { "FAIL" }
            } catch {
                $results += "FAIL"
            }
            $results += "`r`n"
        }
        
        # Test 2: Path accessibility
        $results += "Drive Access: "
        try {
            $accessible = Test-Path $path -ErrorAction Stop
            $results += if ($accessible) { "OK" } else { "FAIL" }
        } catch {
            $results += "FAIL"
        }
        $results += "`r`n"
        
        # Test 3: Write permissions
        $results += "Write Test: "
        try {
            if (Test-Path $path) {
                $tempFile = Join-Path $path "ndrivetest_$(Get-Random).tmp"
                Set-Content -Path $tempFile -Value "test" -ErrorAction Stop
                Remove-Item $tempFile -Force -ErrorAction Stop
                $results += "OK"
            } else {
                $results += "N/A"
            }
        } catch {
            $results += "DENIED"
        }
        $results += "`r`n"
        
        # Test 4: SMB port
        if ($server) {
            $results += "SMB Port 445: "
            try {
                $tcp = Test-NetConnection -ComputerName $server -Port 445 -InformationLevel Quiet -ErrorAction Stop
                $results += if ($tcp) { "Open" } else { "Blocked" }
            } catch {
                $results += "Unknown"
            }
            $results += "`r`n"
        }
        
        # Check for insecure protocols
        if ($server) {
            try {
                $smbConnections = Get-SmbConnection -ErrorAction SilentlyContinue | Where-Object { $_.ServerName -eq $server }
                if ($smbConnections) {
                    foreach ($conn in $smbConnections) {
                        if ($conn.Dialect -lt 2.0) {
                            $results += "⚠ WARNING: Using insecure SMB1 protocol!`r`n"
                        }
                    }
                }
            } catch {}
        }
        
        # Update drive object
        $drive.LastTest = $timestamp
        $drive.TestResult = if ($results -match "FAIL") { "Issues found" } else { "All tests passed" }
        
    } catch {
        $results += "Error during health test: $($_.Exception.Message)`r`n"
        $drive.LastTest = $timestamp
        $drive.TestResult = "Test failed"
    }
    
    return $results
}

function Test-AllDrives {
    try {
        Write-Log "Starting comprehensive drive health test..."
        $script:lblStatus.Text = "Testing all drives..."
        
        $resultsDialog = New-Object System.Windows.Forms.Form
        $resultsDialog.Text = "Drive Health Test Results"
        $resultsDialog.Size = New-Object System.Drawing.Size(800, 600)
        $resultsDialog.StartPosition = "CenterParent"
        
        $resultsText = New-Object System.Windows.Forms.TextBox
        $resultsText.Multiline = $true
        $resultsText.ScrollBars = "Vertical"
        $resultsText.ReadOnly = $true
        $resultsText.Font = New-Object System.Drawing.Font("Consolas", 9)
        $resultsText.Size = New-Object System.Drawing.Size(780, 520)
        $resultsText.Location = New-Object System.Drawing.Point(10, 10)
        $resultsDialog.Controls.Add($resultsText)
        
        $closeBtn = New-Object System.Windows.Forms.Button
        $closeBtn.Text = "Close"
        $closeBtn.Size = New-Object System.Drawing.Size(100, 30)
        $closeBtn.Location = New-Object System.Drawing.Point(690, 540)
        $closeBtn.Add_Click({ $resultsDialog.Close() })
        $resultsDialog.Controls.Add($closeBtn)
        
        $allResults = "=== COMPREHENSIVE DRIVE HEALTH TEST ===`r`n"
        $allResults += "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n`r`n"
        
        foreach ($drive in $script:MappedDrives) {
            $allResults += "Drive: $($drive.Letter): → $($drive.Remote) [$($drive.User)]`r`n"
            $allResults += "-" * 50 + "`r`n"
            $testResults = Test-DriveHealth $drive
            $allResults += $testResults + "`r`n"
        }
        
        $allResults += "=== TEST COMPLETED ===`r`n"
        $allResults += "Finished: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        
        $resultsText.Text = $allResults
        $resultsDialog.ShowDialog() | Out-Null
        
        # Refresh the main display
        Update-DetailsBox
        $script:lblStatus.Text = "Drive health test completed"
        Write-Log "Comprehensive drive health test completed"
        
    } catch {
        Write-Log "Error during drive health test: $($_.Exception.Message)" "ERROR"
        $script:lblStatus.Text = "Error during health test"
    }
}

# Credential management functions
function Show-CredentialManager {
    try {
        $credDialog = New-Object System.Windows.Forms.Form
        $credDialog.Text = "Windows Credential Manager"
        $credDialog.Size = New-Object System.Drawing.Size(700, 500)
        $credDialog.StartPosition = "CenterParent"
        
        $instructionLabel = New-Object System.Windows.Forms.Label
        $instructionLabel.Text = "Stored network credentials (select and delete problematic entries):"
        $instructionLabel.Location = New-Object System.Drawing.Point(10, 10)
        $instructionLabel.Size = New-Object System.Drawing.Size(650, 20)
        $credDialog.Controls.Add($instructionLabel)
        
        $credListBox = New-Object System.Windows.Forms.ListBox
        $credListBox.Size = New-Object System.Drawing.Size(670, 350)
        $credListBox.Location = New-Object System.Drawing.Point(10, 40)
        $credDialog.Controls.Add($credListBox)
        
        # Get stored credentials
        try {
            $credOutput = & cmdkey /list 2>$null
            $credentials = @()
            foreach ($line in $credOutput) {
                if ($line -match "Target:\s*(.+)") {
                    $target = $matches[1].Trim()
                    if ($target -match "TERMSRV|Domain|MicrosoftAccount") { continue }
                    $credentials += $target
                    $credListBox.Items.Add($target)
                }
            }
        } catch {
            $credListBox.Items.Add("Error retrieving credentials")
        }
        
        $deleteBtn = New-Object System.Windows.Forms.Button
        $deleteBtn.Text = "Delete Selected"
        $deleteBtn.Size = New-Object System.Drawing.Size(120, 30)
        $deleteBtn.Location = New-Object System.Drawing.Point(450, 410)
        $deleteBtn.Add_Click({
            if ($credListBox.SelectedItem) {
                $target = $credListBox.SelectedItem.ToString()
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "Delete credential for: $target ?",
                    "Confirm Deletion",
                    "YesNo",
                    "Question"
                )
                if ($result -eq "Yes") {
                    try {
                        & cmdkey /delete:$target
                        Write-Log "Deleted Windows credential: $target"
                        $credListBox.Items.Remove($credListBox.SelectedItem)
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("Error deleting credential", "Error", "OK", "Error")
                    }
                }
            }
        })
        $credDialog.Controls.Add($deleteBtn)
        
        $closeBtn = New-Object System.Windows.Forms.Button
        $closeBtn.Text = "Close"
        $closeBtn.Size = New-Object System.Drawing.Size(80, 30)
        $closeBtn.Location = New-Object System.Drawing.Point(580, 410)
        $closeBtn.Add_Click({ $credDialog.Close() })
        $credDialog.Controls.Add($closeBtn)
        
        $credDialog.ShowDialog() | Out-Null
        Write-Log "Credential manager accessed"
        
    } catch {
        Write-Log "Error accessing credential manager: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error accessing credential manager: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

# Export functions
function Export-DrivesToCSV {
    try {
        $csvPath = "$env:USERPROFILE\Desktop\NetworkDrives_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
        
        $exportData = @()
        foreach ($drive in $script:MappedDrives) {
            $exportData += [pscustomobject]@{
                User = $drive.User
                DriveLetter = $drive.Letter
                RemotePath = $drive.Remote
                Provider = $drive.Provider
                Status = $drive.Status
                Persistent = $drive.Persistent
                LastTest = $drive.LastTest
                TestResult = $drive.TestResult
            }
        }
        
        $exportData | Export-Csv -NoTypeInformation -Path $csvPath
        
        Write-Log "Exported $($script:MappedDrives.Count) drive records to: $csvPath"
        $script:lblStatus.Text = "Drives exported to: $csvPath"
        
        [System.Windows.Forms.MessageBox]::Show("Drive data exported to:`n$csvPath", "Export Complete", "OK", "Information")
        
    } catch {
        Write-Log "Error exporting drives: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error exporting drives: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

function Export-LogToFile {
    try {
        $exportPath = "$env:USERPROFILE\Desktop\NetworkDriveManager_Log_$(Get-Date -Format yyyyMMdd_HHmmss).txt"
        
        if (Test-Path $script:LogFile) {
            Copy-Item $script:LogFile $exportPath
        } else {
            $script:txtLog.Text | Out-File -FilePath $exportPath -Encoding UTF8
        }
        
        Write-Log "Log exported to: $exportPath"
        $script:lblStatus.Text = "Log exported to: $exportPath"
        
        [System.Windows.Forms.MessageBox]::Show("Activity log exported to:`n$exportPath", "Export Complete", "OK", "Information")
        
    } catch {
        Write-Log "Error exporting log: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error exporting log: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

# User-friendly enhancement functions
function Open-SelectedInExplorer {
    if ($script:SelectedDriveIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a drive to open.", "No Selection", "OK", "Information")
        return
    }
    
    try {
        $selectedDrive = $script:MappedDrives[$script:SelectedDriveIndex]
        $drivePath = "$($selectedDrive.Letter):"
        
        if (Test-Path $drivePath) {
            Start-Process explorer.exe $drivePath
            Write-Log "Opened drive $($selectedDrive.Letter): in Explorer"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Drive $($selectedDrive.Letter): is not accessible", "Drive Not Accessible", "OK", "Warning")
        }
    } catch {
        Write-Log "Error opening drive in Explorer: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error opening drive in Explorer: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

function Copy-SelectedUNCPath {
    if ($script:SelectedDriveIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a drive to copy its path.", "No Selection", "OK", "Information")
        return
    }
    
    try {
        $selectedDrive = $script:MappedDrives[$script:SelectedDriveIndex]
        Set-Clipboard -Value $selectedDrive.Remote
        Write-Log "Copied UNC path to clipboard: $($selectedDrive.Remote)"
        $script:lblStatus.Text = "UNC path copied to clipboard"
    } catch {
        Write-Log "Error copying UNC path: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error copying UNC path: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

# Drive management functions (Create, Edit, Remove)
function Show-NewDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Create New Network Drive"
    $dialog.Size = New-Object System.Drawing.Size(400, 300)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $driveLetterLabel = New-Object System.Windows.Forms.Label
    $driveLetterLabel.Text = "Drive Letter:"
    $driveLetterLabel.Location = New-Object System.Drawing.Point(10, 20)
    $driveLetterLabel.Size = New-Object System.Drawing.Size(80, 20)
    $dialog.Controls.Add($driveLetterLabel)

    $driveLetterCombo = New-Object System.Windows.Forms.ComboBox
    $driveLetterCombo.DropDownStyle = "DropDownList"
    $driveLetterCombo.Location = New-Object System.Drawing.Point(100, 18)
    $driveLetterCombo.Size = New-Object System.Drawing.Size(60, 25)

    $usedLetters = @("A", "B", "C")
    foreach ($drive in $script:MappedDrives) {
        if ($drive.Letter -and $drive.Letter -ne "N/A") {
            $usedLetters += $drive.Letter.ToUpper()
        }
    }

    for ($i = 68; $i -le 90; $i++) {
        $letter = [char]$i
        if ($letter -notin $usedLetters) {
            $driveLetterCombo.Items.Add($letter)
        }
    }
    if ($driveLetterCombo.Items.Count -gt 0) { $driveLetterCombo.SelectedIndex = 0 }
    $dialog.Controls.Add($driveLetterCombo)

    $remotePathLabel = New-Object System.Windows.Forms.Label
    $remotePathLabel.Text = "Remote Path:"
    $remotePathLabel.Location = New-Object System.Drawing.Point(10, 60)
    $remotePathLabel.Size = New-Object System.Drawing.Size(80, 20)
    $dialog.Controls.Add($remotePathLabel)

    $remotePathTextBox = New-Object System.Windows.Forms.TextBox
    $remotePathTextBox.Location = New-Object System.Drawing.Point(100, 58)
    $remotePathTextBox.Size = New-Object System.Drawing.Size(280, 25)
    $remotePathTextBox.Text = "\\"
    $dialog.Controls.Add($remotePathTextBox)

    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Text = "Username:"
    $usernameLabel.Location = New-Object System.Drawing.Point(10, 100)
    $usernameLabel.Size = New-Object System.Drawing.Size(80, 20)
    $dialog.Controls.Add($usernameLabel)

    $usernameTextBox = New-Object System.Windows.Forms.TextBox
    $usernameTextBox.Location = New-Object System.Drawing.Point(100, 98)
    $usernameTextBox.Size = New-Object System.Drawing.Size(280, 25)
    $dialog.Controls.Add($usernameTextBox)

    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Text = "Password:"
    $passwordLabel.Location = New-Object System.Drawing.Point(10, 140)
    $passwordLabel.Size = New-Object System.Drawing.Size(80, 20)
    $dialog.Controls.Add($passwordLabel)

    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.UseSystemPasswordChar = $true
    $passwordTextBox.Location = New-Object System.Drawing.Point(100, 138)
    $passwordTextBox.Size = New-Object System.Drawing.Size(280, 25)
    $dialog.Controls.Add($passwordTextBox)

    $persistentCheckBox = New-Object System.Windows.Forms.CheckBox
    $persistentCheckBox.Text = "Make persistent (reconnect at logon)"
    $persistentCheckBox.Location = New-Object System.Drawing.Point(100, 180)
    $persistentCheckBox.Size = New-Object System.Drawing.Size(280, 25)
    $persistentCheckBox.Checked = $true
    $dialog.Controls.Add($persistentCheckBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Create"
    $okButton.Location = New-Object System.Drawing.Point(220, 220)
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Add_Click({
        $result = Create-NetworkDrive $driveLetterCombo.SelectedItem $remotePathTextBox.Text $usernameTextBox.Text $passwordTextBox.Text $persistentCheckBox.Checked
        if ($result) {
            $dialog.DialogResult = "OK"
            $dialog.Close()
        }
    })
    $dialog.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(305, 220)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
    $cancelButton.Add_Click({
        $dialog.DialogResult = "Cancel"
        $dialog.Close()
    })
    $dialog.Controls.Add($cancelButton)

    $dialog.ShowDialog() | Out-Null
    if ($dialog.DialogResult -eq "OK") {
        Refresh-DriveList
    }
}

function Create-NetworkDrive($driveLetter, $remotePath, $username, $password, $persistent) {
    try {
        Write-Log "Creating drive ${driveLetter}: -> $remotePath (Persistent: $persistent)"
        
        $params = @{
            Name = $driveLetter
            PSProvider = "FileSystem"
            Root = $remotePath
        }
        
        if ($persistent) { 
            $params.Add("Persist", $true) 
        }
        
        if ($username -and $password) {
            $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
            $params.Add("Credential", $credential)
            Write-Log "Using provided credentials for ${driveLetter}:"
        }
        
        New-PSDrive @params -ErrorAction Stop | Out-Null
        Write-Log "Successfully created drive ${driveLetter}: -> $remotePath"
        $script:lblStatus.Text = "Drive ${driveLetter}: created successfully"
        return $true
        
    } catch {
        Write-Log "Error creating drive ${driveLetter}: $($_.Exception.Message)" "ERROR"
        $script:lblStatus.Text = "Error creating drive: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error creating drive: $($_.Exception.Message)", "Error", "OK", "Error")
        return $false
    }
}

function Show-EditDialog {
    if ($script:SelectedDriveIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a drive to edit.", "No Selection", "OK", "Information")
        return
    }
    
    $selectedDrive = $script:MappedDrives[$script:SelectedDriveIndex]
    
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Edit Network Drive - $($selectedDrive.Letter):"
    $dialog.Size = New-Object System.Drawing.Size(500, 200)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "Drive: $($selectedDrive.Letter): → $($selectedDrive.Remote)"
    $infoLabel.Location = New-Object System.Drawing.Point(10, 20)
    $infoLabel.Size = New-Object System.Drawing.Size(470, 40)
    $infoLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $dialog.Controls.Add($infoLabel)

    $reconnectButton = New-Object System.Windows.Forms.Button
    $reconnectButton.Text = "Reconnect Drive"
    $reconnectButton.Location = New-Object System.Drawing.Point(10, 80)
    $reconnectButton.Size = New-Object System.Drawing.Size(120, 30)
    $reconnectButton.Add_Click({
        $success = Reconnect-Drive $selectedDrive.Letter $selectedDrive.Remote
        if ($success) { $dialog.Close() }
    })
    $dialog.Controls.Add($reconnectButton)

    $testButton = New-Object System.Windows.Forms.Button
    $testButton.Text = "Test Connection"
    $testButton.Location = New-Object System.Drawing.Point(140, 80)
    $testButton.Size = New-Object System.Drawing.Size(120, 30)
    $testButton.Add_Click({
        $testResults = Test-DriveHealth $selectedDrive
        [System.Windows.Forms.MessageBox]::Show("Test Results:`n`n$testResults", "Connection Test", "OK", "Information")
        Update-DetailsBox
    })
    $dialog.Controls.Add($testButton)

    $openButton = New-Object System.Windows.Forms.Button
    $openButton.Text = "Open in Explorer"
    $openButton.Location = New-Object System.Drawing.Point(270, 80)
    $openButton.Size = New-Object System.Drawing.Size(120, 30)
    $openButton.Add_Click({
        try {
            if (Test-Path "$($selectedDrive.Letter):") {
                Start-Process explorer.exe "$($selectedDrive.Letter):"
            } else {
                [System.Windows.Forms.MessageBox]::Show("Drive is not accessible", "Error", "OK", "Warning")
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error opening drive: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })
    $dialog.Controls.Add($openButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(400, 80)
    $closeButton.Size = New-Object System.Drawing.Size(75, 30)
    $closeButton.Add_Click({ $dialog.Close() })
    $dialog.Controls.Add($closeButton)

    $dialog.ShowDialog() | Out-Null
    Refresh-DriveList
}

function Reconnect-Drive($driveLetter, $remotePath) {
    try {
        Write-Log "Reconnecting drive ${driveLetter}: to $remotePath"
        
        # First, try to remove existing connection
        try {
            Remove-PSDrive -Name $driveLetter -Force -ErrorAction SilentlyContinue
            & net use "${driveLetter}:" /delete /y 2>$null
        } catch {}
        
        # Reconnect using net use
        $result = & net use "${driveLetter}:" "$remotePath" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Successfully reconnected drive ${driveLetter}:"
            $script:lblStatus.Text = "Drive ${driveLetter}: reconnected successfully"
            return $true
        } else {
            throw "Net use failed: $result"
        }
        
    } catch {
        Write-Log "Error reconnecting drive ${driveLetter}: $($_.Exception.Message)" "ERROR"
        $script:lblStatus.Text = "Error reconnecting drive"
        [System.Windows.Forms.MessageBox]::Show("Error reconnecting drive: $($_.Exception.Message)", "Error", "OK", "Error")
        return $false
    }
}

function Remove-Selected {
    if ($script:SelectedDriveIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a drive to delete.", "No Selection", "OK", "Information")
        return
    }
    
    $selectedDrive = $script:MappedDrives[$script:SelectedDriveIndex]
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to remove drive $($selectedDrive.Letter): → $($selectedDrive.Remote)?`n`nUser: $($selectedDrive.User)",
        "Confirm Deletion",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        Remove-NetworkDrive $selectedDrive.Letter
        Refresh-DriveList
    }
}

function Remove-NetworkDrive($driveLetter) {
    try {
        Write-Log "Removing drive ${driveLetter}:"
        
        # Try PowerShell method first
        try {
            Remove-PSDrive -Name $driveLetter -Force -ErrorAction Stop
            Write-Log "Drive ${driveLetter}: removed using PowerShell"
        } catch {
            # Fallback to net use
            & net use "${driveLetter}:" /delete /y 2>$null
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Drive ${driveLetter}: removed using net use"
            } else {
                throw "Both removal methods failed"
            }
        }
        
        $script:lblStatus.Text = "Drive ${driveLetter}: removed successfully"
        
    } catch {
        Write-Log "Error removing drive ${driveLetter}: $($_.Exception.Message)" "ERROR"
        $script:lblStatus.Text = "Error removing drive"
        [System.Windows.Forms.MessageBox]::Show("Error removing drive: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

# Troubleshooting dialog - UPDATED WITH ANIMATED COPY LINK
function Show-Troubleshoot {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Network Drive Troubleshooting"
    $dialog.Size = New-Object System.Drawing.Size(740, 810)  # Made larger as requested
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "Sizable"

    $instructionsLabel = New-Object System.Windows.Forms.Label
    $instructionsLabel.Text = "Select troubleshooting actions to diagnose and fix network drive issues:"
    $instructionsLabel.Location = New-Object System.Drawing.Point(10, 10)
    $instructionsLabel.Size = New-Object System.Drawing.Size(700, 30)
    $dialog.Controls.Add($instructionsLabel)

    $checkBoxes = @()
    $options = @(
        "Check internet connectivity (DNS+HTTPS test)",
        "Test SMB client/protocol status and security",
        "List and manage cached credentials",
        "Check for elevation and session issues",
        "Verify remote path accessibility for all drives",
        "Check for conflicting drive mappings",
        "Restart Windows Explorer",
        "Generate comprehensive diagnostic report"
    )

    $y = 50
    foreach ($option in $options) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = $option
        $checkBox.Location = New-Object System.Drawing.Point(20, [int]$y)
        $checkBox.Size = New-Object System.Drawing.Size(650, 25)
        $checkBox.Checked = $true  # ALL BOXES CHECKED BY DEFAULT
        $dialog.Controls.Add($checkBox)
        $checkBoxes += $checkBox
        $y += 30
    }

    $resultsLabel = New-Object System.Windows.Forms.Label
    $resultsLabel.Text = "Results:"
    $resultsLabel.Location = New-Object System.Drawing.Point(10, [int]($y + 20))
    $resultsLabel.Size = New-Object System.Drawing.Size(100, 20)
    $dialog.Controls.Add($resultsLabel)

    # ANIMATED COPY LINK - RIGHT ALIGNED ABOVE RESULTS BOX
    $copyLink = New-Object System.Windows.Forms.LinkLabel
    $copyLink.Text = "Copy Results to Clipboard"
    $copyLink.AutoSize = $true
    $copyLink.Location = New-Object System.Drawing.Point(530, [int]($y + 20))  # Right aligned
    $copyLink.LinkColor = "DimGray"
    $copyLink.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $copyLink.Enabled = $false
    $dialog.Controls.Add($copyLink)

    # Animation timer setup
    $animationTimer = New-Object System.Windows.Forms.Timer
    $copyLinkBlinkState = $false
    $animationTimer.Interval = 400
    $animationTimer.add_Tick({
        if ($copyLink.Enabled) {
            $copyLinkBlinkState = -not $copyLinkBlinkState
            $copyLink.LinkColor = if ($copyLinkBlinkState) { "OrangeRed" } else { "Green" }
        } else {
            $animationTimer.Stop()
            $copyLink.LinkColor = "DimGray"
        }
    })

    $copyLink.add_Click({
        if ($resultsTextBox.Text) {
            Set-Clipboard -Value $resultsTextBox.Text
            $copyLink.LinkColor = "DimGray"
            $copyLink.Enabled = $false
            $animationTimer.Stop()
            [System.Windows.Forms.MessageBox]::Show("Results copied to clipboard.", "Copied", "OK", "Information")
        }
    })

    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = "Vertical"
    $resultsTextBox.ReadOnly = $true
    $resultsTextBox.Location = New-Object System.Drawing.Point(10, [int]($y + 45))
    $resultsTextBox.Size = New-Object System.Drawing.Size(700, 460)  # Made larger
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dialog.Controls.Add($resultsTextBox)

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "Run Selected Tests"
    $runButton.Location = New-Object System.Drawing.Point(440, [int]($y + 525))
    $runButton.Size = New-Object System.Drawing.Size(170, 30)
    $runButton.Add_Click({
        $selectedTests = @()
        for ($i = 0; $i -lt $checkBoxes.Count; $i++) {
            if ($checkBoxes[$i].Checked) {
                $selectedTests += $i
            }
        }
        if ($selectedTests.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one test to run.", "No Tests Selected", "OK", "Information")
            return
        }
        
        # Reset link state before running tests
        $resultsTextBox.Text = "Running troubleshooting tests...`r`n`r`n"
        $copyLink.Enabled = $false
        $animationTimer.Stop()
        $copyLink.LinkColor = "DimGray"
        
        # Run tests
        $results = Run-TroubleshootingTests $selectedTests
        $resultsTextBox.Text += $results
        
        # Enable and animate the copy link now that results are available
        $copyLink.Enabled = $true
        $animationTimer.Start()
    })
    $dialog.Controls.Add($runButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(620, [int]($y + 525))
    $closeButton.Size = New-Object System.Drawing.Size(75, 30)
    $closeButton.Add_Click({ $dialog.Close() })
    $dialog.Controls.Add($closeButton)

    $dialog.ShowDialog() | Out-Null
}

function Run-TroubleshootingTests($selectedTests) {
    $results = ""
    
    foreach ($testIndex in $selectedTests) {
        switch ($testIndex) {
            0 { # Internet connectivity
                $results += "=== Internet Connectivity Test ===`r`n"
                try {
                    $dns = $null
                    try { $dns = Resolve-DnsName www.microsoft.com -ErrorAction Stop } catch {}
                    
                    $ping = $false
                    try { 
                        $netTest = Test-NetConnection www.microsoft.com -Port 443 -InformationLevel Quiet -ErrorAction Stop
                        $ping = $netTest
                    } catch {}
                    
                    if ($ping) {
                        $results += "✓ Can connect to www.microsoft.com:443 (HTTPS)`r`n"
                    } else {
                        try {
                            Invoke-WebRequest -Uri "https://www.microsoft.com/favicon.ico" -UseBasicParsing -TimeoutSec 5 | Out-Null
                            $results += "✓ HTTP fallback test: Success`r`n"
                        } catch {
                            $results += "✗ HTTP test: Failed`r`n"
                        }
                    }
                    
                    if ($dns) {
                        $results += "✓ DNS resolution: Success`r`n"
                    } else {
                        $results += "✗ DNS resolution: Failed`r`n"
                    }
                } catch {
                    $results += "✗ Network test error: $($_.Exception.Message)`r`n"
                }
                $results += "`r`n"
            }
            
            1 { # SMB client/protocol status
                $results += "=== SMB Client Protocol Test ===`r`n"
                try {
                    $smbClient = Get-SmbClientConfiguration -ErrorAction SilentlyContinue
                    if ($smbClient) {
                        $results += "✓ SMB Client Available`r`n"
                        $results += "SMB1 Enabled: $($smbClient.EnableSMB1Protocol)`r`n"
                        $results += "SMB2 Enabled: $($smbClient.EnableSMB2Protocol)`r`n"
                        
                        if ($smbClient.EnableSMB1Protocol) {
                            $results += "⚠ WARNING: SMB1 is enabled - security risk!`r`n"
                        }
                        
                        # Check active connections
                        $smbConnections = Get-SmbConnection -ErrorAction SilentlyContinue
                        if ($smbConnections) {
                            $results += "Active SMB Connections: $($smbConnections.Count)`r`n"
                            foreach ($conn in $smbConnections) {
                                $results += "  → $($conn.ServerName) (SMB $($conn.Dialect))`r`n"
                                if ($conn.Dialect -lt 2.0) {
                                    $results += "    ⚠ Using insecure SMB1!`r`n"
                                }
                            }
                        }
                    } else {
                        $results += "✗ SMB Client: Not available or accessible`r`n"
                    }
                } catch {
                    $results += "✗ SMB test error: $($_.Exception.Message)`r`n"
                }
                $results += "`r`n"
            }
            
            2 { # Cached credentials
                $results += "=== Cached Credentials Analysis ===`r`n"
                try {
                    $credOutput = & cmdkey /list 2>$null
                    $networkCreds = @()
                    
                    foreach ($line in $credOutput) {
                        if ($line -match "Target:\s*(.+)" -and $line -notmatch "TERMSRV|MicrosoftAccount") {
                            $target = $matches[1].Trim()
                            if ($target -match "^\\\\") {
                                $networkCreds += $target
                            }
                        }
                    }
                    
                    if ($networkCreds.Count -gt 0) {
                        $results += "✓ Found $($networkCreds.Count) network credentials:`r`n"
                        foreach ($cred in $networkCreds) {
                            $results += "  → $cred`r`n"
                        }
                        $results += "Note: Use Credential Manager to review/remove if needed`r`n"
                    } else {
                        $results += "ℹ No network credentials stored`r`n"
                    }
                } catch {
                    $results += "✗ Credential analysis error: $($_.Exception.Message)`r`n"
                }
                $results += "`r`n"
            }
            
            3 { # Elevation and session issues
                $results += "=== Session Context Analysis ===`r`n"
                
                if ($script:IsAdmin) {
                    $results += "⚠ Running as Administrator`r`n"
                    $results += "  → Mapped drives may not be visible in File Explorer`r`n"
                    $results += "  → Recommendation: Run as regular user for drive visibility`r`n"
                } else {
                    $results += "✓ Running as regular user`r`n"
                }
                
                # Check session type
                $sessionName = $env:SESSIONNAME
                if ($sessionName) {
                    $results += "Session Type: $sessionName`r`n"
                    if ($sessionName -match "RDP") {
                        $results += "ℹ Remote Desktop session detected`r`n"
                    }
                }
                
                # Check if running in correct user context
                $currentUser = $env:USERNAME
                $results += "Current User: $currentUser`r`n"
                $results += "User Profile: $env:USERPROFILE`r`n"
                $results += "`r`n"
            }
            
            4 { # Remote path accessibility
                $results += "=== Drive Accessibility Test ===`r`n"
                if ($script:MappedDrives.Count -eq 0) {
                    $results += "ℹ No mapped drives to test`r`n"
                } else {
                    foreach ($drive in $script:MappedDrives) {
                        $results += "Testing $($drive.Letter): → $($drive.Remote)`r`n"
                        try {
                            $drivePath = "$($drive.Letter):"
                            $accessible = Test-Path $drivePath
                            
                            if ($accessible) {
                                $results += "  ✓ Drive accessible`r`n"
                                
                                # Test write permissions
                                try {
                                    $testFile = Join-Path $drivePath "writetest_$(Get-Random).tmp"
                                    Set-Content -Path $testFile -Value "test" -ErrorAction Stop
                                    Remove-Item $testFile -Force -ErrorAction Stop
                                    $results += "  ✓ Write permissions OK`r`n"
                                } catch {
                                    $results += "  ⚠ Write permissions: DENIED`r`n"
                                }
                            } else {
                                $results += "  ✗ Drive not accessible`r`n"
                            }
                        } catch {
                            $results += "  ✗ Test error: $($_.Exception.Message)`r`n"
                        }
                        $results += "`r`n"
                    }
                }
            }
            
            5 { # Conflicting drives
                $results += "=== Drive Conflict Analysis ===`r`n"
                $driveLetters = $script:MappedDrives | Group-Object { "$($_.User):$($_.Letter)" }
                $conflicts = $driveLetters | Where-Object { $_.Count -gt 1 }
                
                if ($conflicts) {
                    $results += "⚠ Found conflicting drive mappings:`r`n"
                    foreach ($conflict in $conflicts) {
                        $results += "  → $($conflict.Name): has $($conflict.Count) mappings`r`n"
                    }
                } else {
                    $results += "✓ No conflicting drive letters found`r`n"
                }
                $results += "`r`n"
            }
            
            6 { # Restart Explorer
                $results += "=== Windows Explorer Restart ===`r`n"
                try {
                    $explorerProcesses = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
                    if ($explorerProcesses) {
                        Stop-Process -Name "explorer" -Force -ErrorAction Stop
                        Start-Sleep -Seconds 3
                        Start-Process "explorer.exe"
                        $results += "✓ Windows Explorer restarted successfully`r`n"
                        $results += "ℹ Please wait a moment for Explorer to fully reload`r`n"
                    } else {
                        $results += "ℹ Explorer was not running, starting it now`r`n"
                        Start-Process "explorer.exe"
                    }
                } catch {
                    $results += "✗ Explorer restart error: $($_.Exception.Message)`r`n"
                }
                $results += "`r`n"
            }
            
            7 { # Diagnostic report
                $results += "=== Comprehensive Diagnostic Report ===`r`n"
                $results += "System Information:`r`n"
                $results += "  Computer: $($env:COMPUTERNAME)`r`n"
                $results += "  User: $($env:USERNAME)`r`n"
                $results += "  Domain: $($env:USERDOMAIN)`r`n"
                $results += "  OS Version: $([Environment]::OSVersion.VersionString)`r`n"
                $results += "  PowerShell: $($PSVersionTable.PSVersion)`r`n"
                $results += "  Session: $($env:SESSIONNAME)`r`n"
                $results += "  Admin Rights: $($script:IsAdmin)`r`n"
                $results += "`r`n"
                
                $results += "Network Drive Summary:`r`n"
                $results += "  Total Mapped Drives: $($script:MappedDrives.Count)`r`n"
                
                if ($script:MappedDrives.Count -gt 0) {
                    $providers = $script:MappedDrives | Group-Object Provider
                    foreach ($provider in $providers) {
                        $results += "  $($provider.Name): $($provider.Count) drives`r`n"
                    }
                    
                    $persistent = ($script:MappedDrives | Where-Object { $_.Persistent -eq "Persistent" }).Count
                    $session = ($script:MappedDrives | Where-Object { $_.Persistent -eq "Session" }).Count
                    $results += "  Persistent: $persistent, Session-only: $session`r`n"
                }
                
                $results += "`r`n"
                $results += "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n"
                $results += "`r`n"
            }
        }
    }
    
    return $results
}

# Main execution
try {
    $frm = New-MainForm
    Write-Log "Starting Enhanced Network Drive Manager..."
    Refresh-DriveList
    $frm.Add_Shown({ $frm.Activate() })
    [System.Windows.Forms.Application]::Run($frm)
} catch {
    $errorMsg = "Fatal error: $($_.Exception.Message)"
    Write-Host $errorMsg -ForegroundColor Red
    Write-Log $errorMsg "FATAL"
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "Fatal Error", "OK", "Error")
}
