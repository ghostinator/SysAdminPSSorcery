#Requires -Version 5.1

<#
.SYNOPSIS
    SharePoint/OneDrive Mount GUI - Easy mounting of Microsoft 365 folders using RClone
.DESCRIPTION
    This script provides a user-friendly GUI for mounting SharePoint Online and OneDrive folders
    to local drive letters using RClone. It handles authentication, configuration, and mounting operations.
.AUTHOR
    PowerShell GUI for SharePoint/OneDrive Mounting
.VERSION
    1.0
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:rcloneConfigured = $false
$script:currentMounts = @()

# Check if RClone is installed
function Test-RCloneInstalled {
    try {
        $null = Get-Command rclone -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Install RClone if not present
function Install-RClone {
    $installChoice = [System.Windows.Forms.MessageBox]::Show(
        "RClone is not installed. Would you like to download and install it automatically?`n`nThis will download RClone from the official website.",
        "RClone Not Found",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($installChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $progressForm = New-Object System.Windows.Forms.Form
            $progressForm.Text = "Installing RClone"
            $progressForm.Size = New-Object System.Drawing.Size(400, 150)
            $progressForm.StartPosition = "CenterScreen"
            $progressForm.FormBorderStyle = "FixedDialog"
            $progressForm.MaximizeBox = $false
            $progressForm.MinimizeBox = $false
            
            $progressLabel = New-Object System.Windows.Forms.Label
            $progressLabel.Text = "Downloading and installing RClone..."
            $progressLabel.Location = New-Object System.Drawing.Point(20, 30)
            $progressLabel.Size = New-Object System.Drawing.Size(350, 30)
            $progressForm.Controls.Add($progressLabel)
            
            $progressBar = New-Object System.Windows.Forms.ProgressBar
            $progressBar.Location = New-Object System.Drawing.Point(20, 70)
            $progressBar.Size = New-Object System.Drawing.Size(350, 20)
            $progressBar.Style = "Marquee"
            $progressForm.Controls.Add($progressBar)
            
            $progressForm.Show()
            $progressForm.Refresh()
            
            # Create RClone directory
            $rcloneDir = "$env:ProgramFiles\RClone"
            if (!(Test-Path $rcloneDir)) {
                New-Item -ItemType Directory -Path $rcloneDir -Force | Out-Null
            }
            
            # Download RClone
            $downloadUrl = "https://downloads.rclone.org/rclone-current-windows-amd64.zip"
            $zipPath = "$env:TEMP\rclone.zip"
            
            Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing
            
            # Extract RClone
            Expand-Archive -Path $zipPath -DestinationPath $env:TEMP -Force
            $extractedFolder = Get-ChildItem "$env:TEMP\rclone-*-windows-amd64" | Select-Object -First 1
            
            Copy-Item "$($extractedFolder.FullName)\rclone.exe" -Destination "$rcloneDir\rclone.exe" -Force
            
            # Add to PATH
            $currentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
            if ($currentPath -notlike "*$rcloneDir*") {
                [Environment]::SetEnvironmentVariable("PATH", "$currentPath;$rcloneDir", "Machine")
                $env:PATH += ";$rcloneDir"
            }
            
            # Cleanup
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
            Remove-Item $extractedFolder.FullName -Recurse -Force -ErrorAction SilentlyContinue
            
            $progressForm.Close()
            
            [System.Windows.Forms.MessageBox]::Show(
                "RClone has been installed successfully!",
                "Installation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return $true
        }
        catch {
            if ($progressForm) { $progressForm.Close() }
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install RClone: $($_.Exception.Message)",
                "Installation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
    }
    return $false
}

# Configure RClone for SharePoint/OneDrive
function Set-RCloneConfiguration {
    param(
        [string]$RemoteName,
        [string]$ServiceType,
        [string]$TenantId = "",
        [string]$SiteUrl = ""
    )
    
    try {
        # Start RClone config process
        $configArgs = @("config", "create", $RemoteName, $ServiceType)
        
        if ($ServiceType -eq "sharepoint") {
            $configArgs += "--sharepoint-url", $SiteUrl
            if ($TenantId) {
                $configArgs += "--sharepoint-tenant-id", $TenantId
            }
        }
        
        $configArgs += "--config-file", "$env:APPDATA\rclone\rclone.conf"
        
        # Run RClone config
        $process = Start-Process -FilePath "rclone" -ArgumentList $configArgs -Wait -PassThru -WindowStyle Hidden
        
        if ($process.ExitCode -eq 0) {
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        Write-Error "Configuration failed: $($_.Exception.Message)"
        return $false
    }
}

# Get available drive letters
function Get-AvailableDriveLetters {
    $usedDrives = Get-WmiObject -Class Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID | ForEach-Object { $_.Substring(0,1) }
    $allDrives = 65..90 | ForEach-Object { [char]$_ }
    return $allDrives | Where-Object { $_ -notin $usedDrives -and $_ -notin @('A', 'B', 'C') }
}

# Mount SharePoint/OneDrive
function Mount-SharePointOneDrive {
    param(
        [string]$RemoteName,
        [string]$DriveLetter,
        [string]$RemotePath = ""
    )
    
    try {
        $mountPoint = "${DriveLetter}:"
        $remotePath = if ($RemotePath) { "${RemoteName}:${RemotePath}" } else { "${RemoteName}:" }
        
        $mountArgs = @(
            "mount",
            $remotePath,
            $mountPoint,
            "--vfs-cache-mode", "writes",
            "--vfs-cache-max-age", "1h",
            "--vfs-cache-max-size", "1G",
            "--buffer-size", "32M",
            "--timeout", "1h",
            "--retries", "3",
            "--low-level-retries", "10",
            "--stats", "0"
        )
        
        $process = Start-Process -FilePath "rclone" -ArgumentList $mountArgs -PassThru
        
        # Wait a moment and check if mount was successful
        Start-Sleep -Seconds 3
        
        if (Test-Path $mountPoint) {
            $script:currentMounts += @{
                RemoteName = $RemoteName
                DriveLetter = $DriveLetter
                ProcessId = $process.Id
                RemotePath = $RemotePath
            }
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        Write-Error "Mount failed: $($_.Exception.Message)"
        return $false
    }
}

# Dismount drive
function Dismount-Drive {
    param([string]$DriveLetter)
    
    try {
        $mountPoint = "${DriveLetter}:"
        
        # Find the mount in our tracking
        $mount = $script:currentMounts | Where-Object { $_.DriveLetter -eq $DriveLetter }
        
        if ($mount) {
            # Kill the RClone process
            try {
                Stop-Process -Id $mount.ProcessId -Force -ErrorAction SilentlyContinue
            }
            catch {
                # Process might already be stopped
            }
            
            # Remove from tracking
            $script:currentMounts = $script:currentMounts | Where-Object { $_.DriveLetter -ne $DriveLetter }
        }
        
        # Force unmount if still mounted
        if (Test-Path $mountPoint) {
            Start-Process -FilePath "rclone" -ArgumentList @("umount", $mountPoint) -Wait -WindowStyle Hidden
        }
        
        return $true
    }
    catch {
        Write-Error "Unmount failed: $($_.Exception.Message)"
        return $false
    }
}

# Create the main GUI
function Show-SharePointOneDriveGUI {
    # Main Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SharePoint/OneDrive Mount Tool"
    $form.Size = New-Object System.Drawing.Size(600, 700)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.Icon = [System.Drawing.SystemIcons]::Application
    
    # Title Label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "SharePoint/OneDrive Mount Tool"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(550, 35)
    $titleLabel.TextAlign = "MiddleCenter"
    $form.Controls.Add($titleLabel)
    
    # Service Type Group
    $serviceGroupBox = New-Object System.Windows.Forms.GroupBox
    $serviceGroupBox.Text = "Service Type"
    $serviceGroupBox.Location = New-Object System.Drawing.Point(20, 70)
    $serviceGroupBox.Size = New-Object System.Drawing.Size(550, 80)
    $form.Controls.Add($serviceGroupBox)
    
    $sharePointRadio = New-Object System.Windows.Forms.RadioButton
    $sharePointRadio.Text = "SharePoint Online"
    $sharePointRadio.Location = New-Object System.Drawing.Point(20, 25)
    $sharePointRadio.Size = New-Object System.Drawing.Size(150, 20)
    $sharePointRadio.Checked = $true
    $serviceGroupBox.Controls.Add($sharePointRadio)
    
    $oneDriveRadio = New-Object System.Windows.Forms.RadioButton
    $oneDriveRadio.Text = "OneDrive for Business"
    $oneDriveRadio.Location = New-Object System.Drawing.Point(200, 25)
    $oneDriveRadio.Size = New-Object System.Drawing.Size(180, 20)
    $serviceGroupBox.Controls.Add($oneDriveRadio)
    
    $oneDrivePersonalRadio = New-Object System.Windows.Forms.RadioButton
    $oneDrivePersonalRadio.Text = "OneDrive Personal"
    $oneDrivePersonalRadio.Location = New-Object System.Drawing.Point(400, 25)
    $oneDrivePersonalRadio.Size = New-Object System.Drawing.Size(140, 20)
    $serviceGroupBox.Controls.Add($oneDrivePersonalRadio)
    
    # Configuration Group
    $configGroupBox = New-Object System.Windows.Forms.GroupBox
    $configGroupBox.Text = "Configuration"
    $configGroupBox.Location = New-Object System.Drawing.Point(20, 160)
    $configGroupBox.Size = New-Object System.Drawing.Size(550, 200)
    $form.Controls.Add($configGroupBox)
    
    # Remote Name
    $remoteNameLabel = New-Object System.Windows.Forms.Label
    $remoteNameLabel.Text = "Remote Name:"
    $remoteNameLabel.Location = New-Object System.Drawing.Point(20, 30)
    $remoteNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    $configGroupBox.Controls.Add($remoteNameLabel)
    
    $remoteNameTextBox = New-Object System.Windows.Forms.TextBox
    $remoteNameTextBox.Location = New-Object System.Drawing.Point(130, 28)
    $remoteNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $remoteNameTextBox.Text = "MySharePoint"
    $configGroupBox.Controls.Add($remoteNameTextBox)
    
    # SharePoint URL (only for SharePoint)
    $siteUrlLabel = New-Object System.Windows.Forms.Label
    $siteUrlLabel.Text = "SharePoint Site URL:"
    $siteUrlLabel.Location = New-Object System.Drawing.Point(20, 65)
    $siteUrlLabel.Size = New-Object System.Drawing.Size(130, 20)
    $configGroupBox.Controls.Add($siteUrlLabel)
    
    $siteUrlTextBox = New-Object System.Windows.Forms.TextBox
    $siteUrlTextBox.Location = New-Object System.Drawing.Point(160, 63)
    $siteUrlTextBox.Size = New-Object System.Drawing.Size(370, 20)
    $siteUrlTextBox.PlaceholderText = "https://yourtenant.sharepoint.com/sites/yoursite"
    $configGroupBox.Controls.Add($siteUrlTextBox)
    
    # Tenant ID (optional)
    $tenantIdLabel = New-Object System.Windows.Forms.Label
    $tenantIdLabel.Text = "Tenant ID (optional):"
    $tenantIdLabel.Location = New-Object System.Drawing.Point(20, 100)
    $tenantIdLabel.Size = New-Object System.Drawing.Size(130, 20)
    $configGroupBox.Controls.Add($tenantIdLabel)
    
    $tenantIdTextBox = New-Object System.Windows.Forms.TextBox
    $tenantIdTextBox.Location = New-Object System.Drawing.Point(160, 98)
    $tenantIdTextBox.Size = New-Object System.Drawing.Size(370, 20)
    $tenantIdTextBox.PlaceholderText = "your-tenant-id-guid"
    $configGroupBox.Controls.Add($tenantIdTextBox)
    
    # Remote Path
    $remotePathLabel = New-Object System.Windows.Forms.Label
    $remotePathLabel.Text = "Remote Path:"
    $remotePathLabel.Location = New-Object System.Drawing.Point(20, 135)
    $remotePathLabel.Size = New-Object System.Drawing.Size(100, 20)
    $configGroupBox.Controls.Add($remotePathLabel)
    
    $remotePathTextBox = New-Object System.Windows.Forms.TextBox
    $remotePathTextBox.Location = New-Object System.Drawing.Point(130, 133)
    $remotePathTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $remotePathTextBox.PlaceholderText = "Documents/Folder"
    $configGroupBox.Controls.Add($remotePathTextBox)
    
    # Drive Letter
    $driveLetterLabel = New-Object System.Windows.Forms.Label
    $driveLetterLabel.Text = "Drive Letter:"
    $driveLetterLabel.Location = New-Object System.Drawing.Point(350, 135)
    $driveLetterLabel.Size = New-Object System.Drawing.Size(80, 20)
    $configGroupBox.Controls.Add($driveLetterLabel)
    
    $driveLetterComboBox = New-Object System.Windows.Forms.ComboBox
    $driveLetterComboBox.Location = New-Object System.Drawing.Point(440, 133)
    $driveLetterComboBox.Size = New-Object System.Drawing.Size(60, 20)
    $driveLetterComboBox.DropDownStyle = "DropDownList"
    $availableDrives = Get-AvailableDriveLetters
    $driveLetterComboBox.Items.AddRange($availableDrives)
    if ($availableDrives.Count -gt 0) {
        $driveLetterComboBox.SelectedIndex = 0
    }
    $configGroupBox.Controls.Add($driveLetterComboBox)
    
    # Buttons
    $configureButton = New-Object System.Windows.Forms.Button
    $configureButton.Text = "Configure Remote"
    $configureButton.Location = New-Object System.Drawing.Point(20, 380)
    $configureButton.Size = New-Object System.Drawing.Size(120, 35)
    $configureButton.BackColor = [System.Drawing.Color]::LightBlue
    $form.Controls.Add($configureButton)
    
    $mountButton = New-Object System.Windows.Forms.Button
    $mountButton.Text = "Mount Drive"
    $mountButton.Location = New-Object System.Drawing.Point(160, 380)
    $mountButton.Size = New-Object System.Drawing.Size(120, 35)
    $mountButton.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($mountButton)
    
    $unmountButton = New-Object System.Windows.Forms.Button
    $unmountButton.Text = "Unmount Drive"
    $unmountButton.Location = New-Object System.Drawing.Point(300, 380)
    $unmountButton.Size = New-Object System.Drawing.Size(120, 35)
    $unmountButton.BackColor = [System.Drawing.Color]::LightCoral
    $form.Controls.Add($unmountButton)
    
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Text = "Refresh"
    $refreshButton.Location = New-Object System.Drawing.Point(440, 380)
    $refreshButton.Size = New-Object System.Drawing.Size(80, 35)
    $refreshButton.BackColor = [System.Drawing.Color]::LightYellow
    $form.Controls.Add($refreshButton)
    
    # Status Group
    $statusGroupBox = New-Object System.Windows.Forms.GroupBox
    $statusGroupBox.Text = "Current Mounts"
    $statusGroupBox.Location = New-Object System.Drawing.Point(20, 430)
    $statusGroupBox.Size = New-Object System.Drawing.Size(550, 150)
    $form.Controls.Add($statusGroupBox)
    
    $statusListBox = New-Object System.Windows.Forms.ListBox
    $statusListBox.Location = New-Object System.Drawing.Point(10, 20)
    $statusListBox.Size = New-Object System.Drawing.Size(530, 120)
    $statusGroupBox.Controls.Add($statusListBox)
    
    # Status Label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Ready"
    $statusLabel.Location = New-Object System.Drawing.Point(20, 590)
    $statusLabel.Size = New-Object System.Drawing.Size(550, 20)
    $statusLabel.ForeColor = [System.Drawing.Color]::Green
    $form.Controls.Add($statusLabel)
    
    # Event Handlers
    $sharePointRadio.Add_CheckedChanged({
        $siteUrlLabel.Visible = $sharePointRadio.Checked
        $siteUrlTextBox.Visible = $sharePointRadio.Checked
    })
    
    $oneDriveRadio.Add_CheckedChanged({
        $siteUrlLabel.Visible = $false
        $siteUrlTextBox.Visible = $false
    })
    
    $oneDrivePersonalRadio.Add_CheckedChanged({
        $siteUrlLabel.Visible = $false
        $siteUrlTextBox.Visible = $false
    })
    
    $configureButton.Add_Click({
        $statusLabel.Text = "Configuring remote..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        $form.Refresh()
        
        $remoteName = $remoteNameTextBox.Text.Trim()
        if ([string]::IsNullOrEmpty($remoteName)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a remote name.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $statusLabel.Text = "Configuration cancelled"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            return
        }


        
        # Launch RClone config in a new window for authentication
        try {
            $configProcess = Start-Process -FilePath "rclone" -ArgumentList @("config") -Wait -PassThru
            
            if ($configProcess.ExitCode -eq 0) {
                $script:rcloneConfigured = $true
                $statusLabel.Text = "Remote configured successfully"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                [System.Windows.Forms.MessageBox]::Show("Remote configuration completed. You can now mount the drive.", "Configuration Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                $statusLabel.Text = "Configuration failed"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            }
        }
        catch {
            $statusLabel.Text = "Configuration error: $($_.Exception.Message)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })
    
    $mountButton.Add_Click({
        $remoteName = $remoteNameTextBox.Text.Trim()
        $driveLetter = $driveLetterComboBox.SelectedItem
        $remotePath = $remotePathTextBox.Text.Trim()
        
        if ([string]::IsNullOrEmpty($remoteName)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a remote name.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrEmpty($driveLetter)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a drive letter.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $statusLabel.Text = "Mounting drive..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        $form.Refresh()
        
        if (Mount-SharePointOneDrive -RemoteName $remoteName -DriveLetter $driveLetter -RemotePath $remotePath) {
            $statusLabel.Text = "Drive mounted successfully to ${driveLetter}:"
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
            
            # Refresh the drive letter combo box
            $driveLetterComboBox.Items.Clear()
            $availableDrives = Get-AvailableDriveLetters
            $driveLetterComboBox.Items.AddRange($availableDrives)
            if ($availableDrives.Count -gt 0) {
                $driveLetterComboBox.SelectedIndex = 0
            }
            
            # Update status list
            $displayPath = if ($remotePath) { "$remoteName`:$remotePath" } else { "$remoteName`:" }
            $statusListBox.Items.Add("${driveLetter}: -> $displayPath")
        }
        else {
            $statusLabel.Text = "Failed to mount drive"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })
    
    $unmountButton.Add_Click({
        if ($statusListBox.SelectedItem) {
            $selectedMount = $statusListBox.SelectedItem.ToString()
            $driveLetter = $selectedMount.Substring(0, 1)
            
            $statusLabel.Text = "Unmounting drive..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            $form.Refresh()
            
            if (Dismount-Drive -DriveLetter $driveLetter) {
                $statusLabel.Text = "Drive unmounted successfully"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                $statusListBox.Items.Remove($statusListBox.SelectedItem)
                
                # Refresh the drive letter combo box
                $driveLetterComboBox.Items.Clear()
                $availableDrives = Get-AvailableDriveLetters
                $driveLetterComboBox.Items.AddRange($availableDrives)
                if ($availableDrives.Count -gt 0) {
                    $driveLetterComboBox.SelectedIndex = 0
                }
            }
            else {
                $statusLabel.Text = "Failed to unmount drive"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select a mount to unmount.", "Selection Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    
    $refreshButton.Add_Click({
        # Refresh drive letters
        $driveLetterComboBox.Items.Clear()
        $availableDrives = Get-AvailableDriveLetters
        $driveLetterComboBox.Items.AddRange($availableDrives)
        if ($availableDrives.Count -gt 0) {
            $driveLetterComboBox.SelectedIndex = 0
        }
        
        # Refresh mount status
        $statusListBox.Items.Clear()
        foreach ($mount in $script:currentMounts) {
            $displayPath = if ($mount.RemotePath) { "$($mount.RemoteName):$($mount.RemotePath)" } else { "$($mount.RemoteName):" }
            $statusListBox.Items.Add("$($mount.DriveLetter): -> $displayPath")
        }
        
        $statusLabel.Text = "Refreshed"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    })
    
    # Form closing event
    $form.Add_FormClosing({
        # Unmount all drives when closing
        foreach ($mount in $script:currentMounts) {
            try {
                Stop-Process -Id $mount.ProcessId -Force -ErrorAction SilentlyContinue
            }
            catch {
                # Process might already be stopped
            }
        }
    })
    
    # Show the form
    $form.ShowDialog()
}

# Main execution
if (-not (Test-RCloneInstalled)) {
    if (-not (Install-RClone)) {
        [System.Windows.Forms.MessageBox]::Show(
            "RClone is required but not installed. Please install RClone manually from https://rclone.org/downloads/",
            "RClone Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
}

# Show the GUI
Show-SharePointOneDriveGUI