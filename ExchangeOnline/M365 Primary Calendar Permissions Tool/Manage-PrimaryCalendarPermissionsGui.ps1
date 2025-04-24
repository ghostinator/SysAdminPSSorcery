#Requires -Modules ExchangeOnlineManagement # Still useful for static analysis/metadata
#Requires -Version 5.1
#Requires -RunAsAdministrator

# Load necessary assemblies for WPF
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms

<#
.SYNOPSIS
Provides a GUI to manage Microsoft 365 permissions for the primary calendar folder.

.DESCRIPTION
This script uses WPF to create a graphical interface allowing administrators to:
- Connect to Exchange Online.
- Check for and install the ExchangeOnlineManagement module if needed.
- Browse or search for target mailboxes (resolving aliases).
- List a user's calendar folders (only primary is actionable).
- View, add, modify, and remove delegate permissions ONLY on the primary '\Calendar' folder.

.NOTES
Version:        1.9
Author:         Brandon Cook
Copyright:      (c) 2025 Brandon Cook
GitHub:         https://github.com/ghostinator/SysAdminPSSorcery
Last Modified:  2025-04-24
Prerequisites:  PowerShell 5.1+, Admin permissions in M365. Internet access may be required for first run to install module.
Limitations:    Due to observed cmdlet limitations, permission management is restricted to the primary '\Calendar' folder.

.EXAMPLE
.\Manage-PrimaryCalendarPermissionsGui.ps1
Launches the GUI interface. Follow the prompts within the application. Installs module if needed.
#>

#region Prerequisite Module Check and Install
Write-Host "--------------------------------------------------" -ForegroundColor Yellow
Write-Host "Checking for required PowerShell module: ExchangeOnlineManagement" -ForegroundColor Yellow
$moduleName = "ExchangeOnlineManagement"
try {
    # Check if module is available
    $installedModule = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue

    if (-not $installedModule) {
        Write-Warning "Module '$moduleName' not found."
        Write-Host "Attempting to install '$moduleName' from PSGallery (Scope: CurrentUser)..." -ForegroundColor Cyan
        Install-Module -Name $moduleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        Write-Host "Module installation attempted successfully." -ForegroundColor Green
        $installedModule = Get-Module -ListAvailable -Name $moduleName -ErrorAction Stop # Use Stop now, it MUST be there
         Write-Host "Module '$moduleName' is now available." -ForegroundColor Green
    } else {
         Write-Host "Module '$moduleName' is already installed." -ForegroundColor Green
    }

    # Import the module into the current session
    Write-Host "Importing '$moduleName' module..." -ForegroundColor Cyan
    Import-Module -Name $moduleName -ErrorAction Stop
    Write-Host "'$moduleName' module imported successfully." -ForegroundColor Green
    Write-Host "--------------------------------------------------" -ForegroundColor Yellow

} catch {
    # Handle errors during module check/install/import
    Write-Error "**************************************************"
    Write-Error "CRITICAL ERROR: Failed to install or import the required '$moduleName' module."
    Write-Error "Error Details: $($_.Exception.Message)"
    Write-Error "Please ensure PowerShell can install modules from the PSGallery (check internet connection and firewall)"
    Write-Error "and that your execution policy allows script execution (e.g., 'Set-ExecutionPolicy RemoteSigned')."
    Write-Error "**************************************************"
    Read-Host -Prompt "Press ENTER to exit script"
    Exit 1 # Exit the script forcefully
}
#endregion

#region WPF GUI Definition (XAML) Main Window
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 Primary Calendar Permissions Tool (v1.9)" Height="600" Width="700" ResizeMode="CanMinimize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="*"/>    <RowDefinition Height="Auto"/> </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/> <ColumnDefinition Width="*"/>    <ColumnDefinition Width="Auto"/> </Grid.ColumnDefinitions>

        <Label Grid.Row="0" Grid.Column="0" Content="Status:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtStatus" Grid.Row="0" Grid.Column="1" IsReadOnly="True" Text="Disconnected" VerticalAlignment="Center" Background="LightGray" FontWeight="Bold"/>
        <Button x:Name="btnConnect" Grid.Row="0" Grid.Column="2" Content="Connect/Disconnect" Padding="5" Margin="5,0,0,0"/>

        <Label Grid.Row="1" Grid.Column="0" Content="Target Mailbox (Owner):" VerticalAlignment="Center" Margin="0,10,0,0"/>
        <Grid Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2">
             <Grid.ColumnDefinitions> <ColumnDefinition Width="*"/> <ColumnDefinition Width="Auto"/> <ColumnDefinition Width="Auto"/> </Grid.ColumnDefinitions>
             <TextBox x:Name="txtTargetMailbox" Grid.Column="0" VerticalAlignment="Center" Margin="0,10,5,0"/>
             <Button x:Name="btnBrowseMailboxes" Grid.Column="1" Content="..." ToolTip="Browse Mailboxes" Padding="5,0" Margin="0,10,5,0" Width="30" IsEnabled="False"/>
             <Button x:Name="btnListCalendars" Grid.Column="2" Content="List Calendars" Padding="5" Margin="0,10,0,0" IsEnabled="False"/>
        </Grid>

        <Label Grid.Row="2" Grid.Column="0" Content="Target Calendar Folder:" VerticalAlignment="Center" Margin="0,5,0,0"/>
        <ListView x:Name="lstTargetCalendar" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" VerticalAlignment="Stretch" Margin="0,5,0,0" Height="120" IsEnabled="False" SelectionMode="Single">
             <ListView.View>
                 <GridView>
                      <GridViewColumn Header="Calendar Name" DisplayMemberBinding="{Binding DisplayName}" Width="Auto"/>
                 </GridView>
             </ListView.View>
        </ListView>

        <Label Grid.Row="3" Grid.Column="0" Content="Delegate User:" VerticalAlignment="Center" Margin="0,5,0,0"/>
        <TextBox x:Name="txtDelegateUser" Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="2" VerticalAlignment="Center" Margin="0,5,0,0" IsEnabled="False"/>

        <Label Grid.Row="4" Grid.Column="0" Content="Permission Level:" VerticalAlignment="Center" Margin="0,5,0,0"/>
        <ComboBox x:Name="cmbPermissionLevel" Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2" VerticalAlignment="Center" Margin="0,5,0,0" IsEnabled="False">
             <ComboBoxItem Content="None"/> <ComboBoxItem Content="Owner"/> <ComboBoxItem Content="PublishingEditor"/> <ComboBoxItem Content="Editor"/> <ComboBoxItem Content="PublishingAuthor"/> <ComboBoxItem Content="Author"/> <ComboBoxItem Content="NoneditingAuthor"/> <ComboBoxItem Content="Reviewer"/> <ComboBoxItem Content="Contributor"/> <ComboBoxItem Content="AvailabilityOnly" IsSelected="True"/> <ComboBoxItem Content="LimitedDetails"/>
        </ComboBox>

        <StackPanel Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,15,0,5">
            <Button x:Name="btnViewPermissions" Content="View Permissions" Padding="10,5" Margin="5" IsEnabled="False"/>
            <Button x:Name="btnAddSetPermission" Content="Add / Set Permission" Padding="10,5" Margin="5" IsEnabled="False"/>
            <Button x:Name="btnRemovePermission" Content="Remove Permission" Padding="10,5" Margin="5" IsEnabled="False"/>
        </StackPanel>

        <Separator Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,10,0,10"/>

        <Label Grid.Row="7" Grid.Column="0" Content="Results:" VerticalAlignment="Top"/>
        <TextBox x:Name="txtResults" Grid.Row="7" Grid.Column="1" Grid.ColumnSpan="2" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Consolas"/>

        <Button x:Name="btnClose" Grid.Row="8" Grid.Column="2" Content="Close" Padding="10,5" Margin="0,10,0,0" IsCancel="True"/>
    </Grid>
</Window>
"@
#endregion

#region WPF GUI Definition (XAML) Mailbox Browser Pop-up
[xml]$xamlMailboxBrowser = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Find Mailbox" Height="400" Width="650" ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False">
    <Grid Margin="10">
        <Grid.RowDefinitions> <RowDefinition Height="Auto"/> <RowDefinition Height="*"/> <RowDefinition Height="Auto"/> </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <Label Content="Filter (Display Name or Email - Min 3 chars):" VerticalAlignment="Center"/>
            <TextBox x:Name="txtMailboxFilter" Width="200" VerticalAlignment="Center" Margin="5,0"/>
            <Button x:Name="btnSearchMailboxes" Content="Search" Padding="5,2" VerticalAlignment="Center"/>
            <Button x:Name="btnLoadAllMailboxes" Content="Load All" Padding="5,2" VerticalAlignment="Center" Margin="5,0,0,0" ToolTip="WARNING: Loading all mailboxes can be very slow on large systems!"/>
        </StackPanel>
        <ListView x:Name="lstMailboxes" Grid.Row="1" ItemsSource="{Binding Mailboxes}" SelectionMode="Single">
            <ListView.View> <GridView> <GridViewColumn Header="Display Name" DisplayMemberBinding="{Binding DisplayName}" Width="200"/> <GridViewColumn Header="Primary Email" DisplayMemberBinding="{Binding PrimarySmtpAddress}" Width="250"/> </GridView> </ListView.View>
        </ListView>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnSelectMailbox" Content="Select" Width="75" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="btnCancelMailboxSearch" Content="Cancel" Width="75" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
#endregion

#region Load XAML and Get Controls
try {
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Find specific controls directly
    $btnBrowseMailboxes = $window.FindName('btnBrowseMailboxes')
    Write-Host "Direct find attempt for 'btnBrowseMailboxes' returned object: $($btnBrowseMailboxes -ne $null)" -ForegroundColor Magenta
    $lstTargetCalendar = $window.FindName('lstTargetCalendar') # Find ListView by name
    Write-Host "Direct find attempt for 'lstTargetCalendar' returned object: $($lstTargetCalendar -ne $null)" -ForegroundColor Magenta

    # Get Other Control References using loop (excluding direct finds)
    $controlsToSkip = @('btnBrowseMailboxes', 'lstTargetCalendar')
    $controls = @{}
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
        if ($_.Name -notin $controlsToSkip) {
             $controls[$_.Name] = $window.FindName($_.Name)
        }
    }

    # Assign controls to more readable variables
    $btnConnect         = $controls['btnConnect']
    $txtStatus          = $controls['txtStatus']
    $txtTargetMailbox   = $controls['txtTargetMailbox']
    $btnListCalendars   = $controls['btnListCalendars']
    $txtDelegateUser    = $controls['txtDelegateUser']
    $cmbPermissionLevel = $controls['cmbPermissionLevel']
    $btnViewPermissions = $controls['btnViewPermissions']
    $btnAddSetPermission= $controls['btnAddSetPermission']
    $btnRemovePermission= $controls['btnRemovePermission']
    $txtResults         = $controls['txtResults']
    $btnClose           = $controls['btnClose']
    # Note: $btnBrowseMailboxes and $lstTargetCalendar assigned directly above

    # Validate critical controls
    if (-not $lstTargetCalendar) { throw "FATAL: Could not find control 'lstTargetCalendar' in loaded XAML!" }
    if (-not $txtStatus) { throw "FATAL: Could not find control 'txtStatus' in loaded XAML!" }
    if (-not $btnConnect) { throw "FATAL: Could not find control 'btnConnect' in loaded XAML!" }
    if (-not $txtTargetMailbox) { throw "FATAL: Could not find control 'txtTargetMailbox' in loaded XAML!" }
    if (-not $btnListCalendars) { throw "FATAL: Could not find control 'btnListCalendars' in loaded XAML!" }
    if (-not $btnBrowseMailboxes) { throw "FATAL: Could not find control 'btnBrowseMailboxes'!" }
    if (-not $txtDelegateUser) { throw "FATAL: Could not find control 'txtDelegateUser'!" }
    if (-not $cmbPermissionLevel) { throw "FATAL: Could not find control 'cmbPermissionLevel'!" }
    if (-not $btnViewPermissions) { throw "FATAL: Could not find control 'btnViewPermissions'!" }
    if (-not $btnAddSetPermission) { throw "FATAL: Could not find control 'btnAddSetPermission'!" }
    if (-not $btnRemovePermission) { throw "FATAL: Could not find control 'btnRemovePermission'!" }
    if (-not $txtResults) { throw "FATAL: Could not find control 'txtResults'!" }
    if (-not $btnClose) { throw "FATAL: Could not find control 'btnClose'!" }

    Write-Host "XAML loaded and critical controls verified." -ForegroundColor DarkGreen

} catch {
    Write-Error "Error loading XAML GUI or finding controls: $($_.Exception.Message)"
    Read-Host -Prompt "Press ENTER to exit script"
    Exit 1
}
#endregion

#region Global State Variable
$Global:IsConnected = $false
$script:CurrentUserPrimarySmtp = $null # Store resolved user email for constructing folder IDs
#endregion

#region Helper Functions

# Function to add output to the results text box
function Write-OutputToGui ($Message, $Type = "Info") {
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $FormattedMessage = "$Timestamp [$Type] : $Message`r`n"
    if ($txtResults) {
        try { $txtResults.Dispatcher.Invoke([Action]{ $txtResults.AppendText($FormattedMessage); $txtResults.ScrollToEnd() }) }
        catch { Write-Host "Fallback Log: $FormattedMessage" }
    } else { Write-Host "Fallback Log (txtResults null): $FormattedMessage" }
}

# Function to get filtered mailboxes (for browser search)
function Get-FilteredMailboxes ($FilterText) {
    if ([string]::IsNullOrWhiteSpace($FilterText) -or $FilterText.Length -lt 3) { Write-OutputToGui "Filter requires at least 3 characters." -Type Warning; Return @() }
    Write-OutputToGui "Searching mailboxes with filter: $FilterText..."
    try { $filterQuery = "(DisplayName -like '*$FilterText*') -or (PrimarySmtpAddress -like '*$FilterText*') -or (EmailAddresses -like '*smtp:$FilterText*')"; $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize 50 -Filter $filterQuery | Select-Object DisplayName, PrimarySmtpAddress | Sort-Object DisplayName; Write-OutputToGui "Found $($mailboxes.Count) matching mailboxes (max 50 shown)."; Return $mailboxes }
    catch { Write-OutputToGui "Error searching mailboxes: $($_.Exception.Message)" -Type Error; Return @() }
}

# Function to get ALL mailboxes (for browser Load All button)
function Get-AllMailboxesUnfiltered {
    Write-Host "Executing Get-Mailbox -ResultSize Unlimited. Please wait..." -ForegroundColor Yellow; Write-OutputToGui "Executing Get-Mailbox -ResultSize Unlimited (this may take a while)..."
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try { $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress | Sort-Object DisplayName; $Stopwatch.Stop(); $Duration = $Stopwatch.Elapsed.ToString("g"); Write-Host "Retrieved $($mailboxes.Count) mailboxes in $Duration." -ForegroundColor Green; Write-OutputToGui "Retrieved $($mailboxes.Count) mailboxes in $Duration."; Return $mailboxes }
    catch { $Stopwatch.Stop(); Write-OutputToGui "Error retrieving all mailboxes: $($_.Exception.Message)" -Type Error; Write-Host "Error retrieving all mailboxes: $($_.Exception.Message)" -ForegroundColor Red; Return @() }
}

# Function to show Mailbox Browser pop-up window
function Show-MailboxBrowserDialog {
    $script:SelectedMailboxFromBrowser = $null
    try {
        $browserReader = (New-Object System.Xml.XmlNodeReader $xamlMailboxBrowser); $browserWindow = [Windows.Markup.XamlReader]::Load($browserReader)
        $txtMailboxFilter = $browserWindow.FindName('txtMailboxFilter'); $btnSearchMailboxes = $browserWindow.FindName('btnSearchMailboxes'); $lstMailboxes = $browserWindow.FindName('lstMailboxes'); $btnSelectMailbox = $browserWindow.FindName('btnSelectMailbox'); $btnCancelMailboxSearch = $browserWindow.FindName('btnCancelMailboxSearch'); $btnLoadAllMailboxes = $browserWindow.FindName('btnLoadAllMailboxes')
        # Validate Pop-up Controls
        if (-not $txtMailboxFilter) { throw "Could not find 'txtMailboxFilter'!" }; if (-not $btnSearchMailboxes) { throw "Could not find 'btnSearchMailboxes'!" }; if (-not $lstMailboxes) { throw "Could not find 'lstMailboxes'!" }; if (-not $btnSelectMailbox) { throw "Could not find 'btnSelectMailbox'!" }; if (-not $btnCancelMailboxSearch) { throw "Could not find 'btnCancelMailboxSearch'!" }; if (-not $btnLoadAllMailboxes) { throw "Could not find 'btnLoadAllMailboxes'!" }
        Write-Host "Mailbox Browser pop-up controls verified." -ForegroundColor DarkGreen
        # Event Handlers for Pop-up
        $btnSearchMailboxes.Add_Click({ try { $filter = $txtMailboxFilter.Text; if ($lstMailboxes) { $results = @(Get-FilteredMailboxes -FilterText $filter); $lstMailboxes.ItemsSource = $results } else { Write-OutputToGui "Mailbox list control invalid." -Type Error } } catch { Write-OutputToGui "Error in Search: $($_.Exception.Message)" -Type Error } })
        $txtMailboxFilter.Add_KeyDown({ param($sender,$e); if ($e.Key -eq 'Enter') { try { $filter = $txtMailboxFilter.Text; if ($lstMailboxes) { $results = @(Get-FilteredMailboxes -FilterText $filter); $lstMailboxes.ItemsSource = $results } else { Write-OutputToGui "Mailbox list control invalid." -Type Error } } catch { Write-OutputToGui "Error in Filter KeyDown: $($_.Exception.Message)" -Type Error } } })
        $btnLoadAllMailboxes.Add_Click({ try { $warningTitle = "Confirm Load All"; $warningMsg = "WARNING: Loading all mailboxes can take time/resources.`n`nProceed?"; $msgBoxResult = [System.Windows.MessageBox]::Show($warningMsg, $warningTitle, [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning); if ($msgBoxResult -eq 'Yes') { Write-OutputToGui "Loading ALL mailboxes..."; $browserWindow.Cursor = [System.Windows.Input.Cursors]::Wait; $allMailboxes = @(Get-AllMailboxesUnfiltered); if ($lstMailboxes) { $lstMailboxes.ItemsSource = $allMailboxes; Write-OutputToGui "Displayed $($allMailboxes.Count) mailboxes." } else { Write-OutputToGui "Mailbox list control invalid." -Type Error } $browserWindow.Cursor = $null } else { Write-OutputToGui "Load all cancelled." } } catch { Write-OutputToGui "Error Load All: $($_.Exception.Message)" -Type Error; if ($browserWindow) {$browserWindow.Cursor = $null} } })
        $btnSelectMailbox.Add_Click({ try { if ($lstMailboxes -and $lstMailboxes.SelectedItem -ne $null) { $script:SelectedMailboxFromBrowser = $lstMailboxes.SelectedItem.PrimarySmtpAddress; $browserWindow.DialogResult = $true; $browserWindow.Close() } elseif ($lstMailboxes) { [System.Windows.MessageBox]::Show("Select mailbox first.", "Selection Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) } else { Write-OutputToGui "Mailbox list control invalid." -Type Error } } catch { Write-OutputToGui "Error Select: $($_.Exception.Message)" -Type Error } })
        $lstMailboxes.Add_MouseDoubleClick({ try { if ($lstMailboxes -and $lstMailboxes.SelectedItem -ne $null) { $script:SelectedMailboxFromBrowser = $lstMailboxes.SelectedItem.PrimarySmtpAddress; $browserWindow.DialogResult = $true; $browserWindow.Close() } elseif (-not $lstMailboxes) { Write-OutputToGui "Mailbox list control invalid." -Type Error } } catch { Write-OutputToGui "Error DblClick: $($_.Exception.Message)" -Type Error } })
        # Show Dialog
        $browserWindow.Owner = $window; $dialogResult = $browserWindow.ShowDialog()
        if ($dialogResult -eq $true) { return $script:SelectedMailboxFromBrowser } else { return $null }
    } catch { $errorMsg = "Error opening/initializing mailbox browser: $($_.Exception.Message)"; Write-OutputToGui $errorMsg -Type Error; [System.Windows.MessageBox]::Show($errorMsg, "Browser Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); Return $null }
}

# Function to resolve mailbox identity (Direct or Alias)
function Resolve-MailboxIdentity ($IdentityInput) {
    if ([string]::IsNullOrWhiteSpace($IdentityInput)) { Write-OutputToGui "Mailbox identity input empty." -Type Error; return $null }
    Write-OutputToGui "Resolving identity: $IdentityInput"
    $script:CurrentUserPrimarySmtp = $null
    try { $mailbox = Get-Mailbox -Identity $IdentityInput -ErrorAction Stop; $script:CurrentUserPrimarySmtp = $mailbox.PrimarySmtpAddress; Write-OutputToGui "Found direct match: $($script:CurrentUserPrimarySmtp)"; return $mailbox.Identity } catch { Write-OutputToGui "Direct lookup failed for '$IdentityInput'. Checking aliases..." -Type Info }
    try { $aliasFilter = "EmailAddresses -eq 'smtp:$($IdentityInput)'"; $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -Filter $aliasFilter -ErrorAction Stop; if ($mailboxes.Count -eq 1) { $resolvedIdentity = $mailboxes[0].Identity; $script:CurrentUserPrimarySmtp = $mailboxes[0].PrimarySmtpAddress; Write-OutputToGui "Found match via alias: '$IdentityInput' belongs to '$($script:CurrentUserPrimarySmtp)'" -Type Info; return $resolvedIdentity } elseif ($mailboxes.Count -gt 1) { Write-OutputToGui "Alias lookup for '$IdentityInput' returned multiple mailboxes." -Type Error; return $null } else { Write-OutputToGui "Could not find mailbox matching '$IdentityInput' directly or as alias." -Type Error; [System.Windows.MessageBox]::Show("Could not find mailbox matching '$IdentityInput'.", "Not Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); return $null } } catch { Write-OutputToGui "Error during alias lookup for '$IdentityInput': $($_.Exception.Message)" -Type Error; return $null }
    return $null
}

# Function to enable/disable controls based on connection status (Direct Access Version)
function Set-GuiState ($Connected) {
    $Global:IsConnected = $Connected
    $connectionControls = @($txtTargetMailbox, $btnListCalendars, $btnBrowseMailboxes)
    $calendarActionControls = @($txtDelegateUser, $cmbPermissionLevel, $btnViewPermissions, $btnAddSetPermission, $btnRemovePermission)
    $targetCalendarList = $lstTargetCalendar # Use correct variable name
    try {
        if ($Connected) {
            $txtStatus.Text = "Connected"; $txtStatus.Background = "LightGreen"; $btnConnect.Content = "Disconnect"
            $connectionControls | ForEach-Object { if ($_) { $_.IsEnabled = $true } }
            $calendarActionControls | ForEach-Object { if ($_) { $_.IsEnabled = $false } }
            if ($targetCalendarList) { $targetCalendarList.IsEnabled = $false }
        } else {
            $txtStatus.Text = "Disconnected"; $txtStatus.Background = "LightGray"; $btnConnect.Content = "Connect"
            $connectionControls | ForEach-Object { if ($_) { $_.IsEnabled = $false } }
            $calendarActionControls | ForEach-Object { if ($_) { $_.IsEnabled = $false } }
            if ($targetCalendarList) { $targetCalendarList.IsEnabled = $false; $targetCalendarList.ItemsSource = $null }
        }
    } catch { Write-Error "FATAL: Error updating GUI state directly for controls. Error: $($_.Exception.Message)" }
}

# Function to get and populate calendar folders (ListView version, Primary first)
function Get-CalendarFoldersForUser ($UserIdentity) {
    Write-OutputToGui "Listing calendar folders for mailbox '$($script:CurrentUserPrimarySmtp)'..."
    $targetCalendarList = $lstTargetCalendar
    if (-not $targetCalendarList) { Write-Warning "lstTargetCalendar control not found."; return }
    # Use Dispatcher as this runs after GUI is potentially shown
    $targetCalendarList.Dispatcher.Invoke([Action]{ $targetCalendarList.ItemsSource = $null; $targetCalendarList.IsEnabled = $false })
    try {
        $foldersData = Get-MailboxFolderStatistics -Identity $UserIdentity -FolderScope Calendar -ErrorAction Stop |
                       Select-Object @{Name='DisplayName';Expression={$_.FolderPath.Replace('/', '\')}},
                                     FolderPath,
                                     @{Name='IsPrimary';Expression={($_.FolderPath -eq '/Calendar') -or ($_.FolderPath -like '*/Calendar')}}

        $folders = @($foldersData) # Ensure array for single item fix
        $folders = $folders | Sort-Object -Property @{Expression={$_.IsPrimary}; Descending=$true}, DisplayName # Sort Primary first

        if ($folders.Count -gt 0) {
            Write-OutputToGui "Found $($folders.Count) calendar folder(s)."
            $targetCalendarList.Dispatcher.Invoke([Action]{
                $targetCalendarList.ItemsSource = $folders
                $targetCalendarList.IsEnabled = $true
                $targetCalendarList.SelectedIndex = 0 # Select first item (Primary)
            })
            # Enable action controls via Dispatcher
            $actionControls = @( $txtDelegateUser, $cmbPermissionLevel, $btnViewPermissions, $btnAddSetPermission, $btnRemovePermission )
            $actionControls | ForEach-Object { if ($_) { $_.Dispatcher.Invoke([Action]{ $_.IsEnabled = $true }) } }
        } else {
            Write-OutputToGui "No calendar folders found for '$($script:CurrentUserPrimarySmtp)'." -Type Warning
            [System.Windows.MessageBox]::Show("No calendar folders found for '$($script:CurrentUserPrimarySmtp)'.", "Not Found", "OK", "Information")
            $actionControls = @( $txtDelegateUser, $cmbPermissionLevel, $btnViewPermissions, $btnAddSetPermission, $btnRemovePermission )
            $actionControls | ForEach-Object { if($_) { $_.IsEnabled = $false } }
        }
    } catch { Write-OutputToGui "Error listing calendar folders for '$($script:CurrentUserPrimarySmtp)': $($_.Exception.Message)" -Type Error; [System.Windows.MessageBox]::Show("Error listing calendars for '$($script:CurrentUserPrimarySmtp)': $($_.Exception.Message)", "Error", "OK", "Error") }
}
#endregion

#region Event Handlers

# Connect/Disconnect Button
$btnConnect.Add_Click({
    if ($Global:IsConnected) {
        Write-OutputToGui "Disconnecting..."
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Set-GuiState $false
        Write-OutputToGui "Disconnected from Exchange Online."
    } else {
        Write-OutputToGui "Connecting to Exchange Online..."
        try {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Write-OutputToGui "Attempting Connect-ExchangeOnline (Disabling WAM)..."
            # Use -DisableWAM to force browser sign-in if WAM causes issues
            Connect-ExchangeOnline -DisableWAM -ShowProgress $false -ShowBanner:$false -ErrorAction Stop -Verbose
            Write-OutputToGui "Connect-ExchangeOnline command finished. Checking status..."
            if ($?) {
                Write-OutputToGui "Connection command successful (`$?` = True)."
                Write-OutputToGui "Connection successful."
                Set-GuiState $true
            } else {
                Write-OutputToGui "Connection command failed (`$?` = False)." -Type Warning
                Write-OutputToGui "Check console for verbose messages/errors." -Type Warning
                if ($Global:IsConnected) { Set-GuiState $false }
            }
        } catch {
            Write-OutputToGui "Failed during Connect process (Caught Exception): $($_.Exception.Message)" -Type Error
            if ($Global:IsConnected) { Set-GuiState $false }
        }
    }
})

# Target Mailbox Text Changed - Clear Calendar List
$txtTargetMailbox.Add_TextChanged({
    if ($Global:IsConnected) {
        if ($lstTargetCalendar) {
             try { $lstTargetCalendar.Dispatcher.Invoke([Action]{ $lstTargetCalendar.ItemsSource = $null; $lstTargetCalendar.IsEnabled = $false }) } catch {}
        }
        $actionControls = @($txtDelegateUser, $cmbPermissionLevel, $btnViewPermissions, $btnAddSetPermission, $btnRemovePermission)
        $actionControls | ForEach-Object { if($_) { try { $_.IsEnabled = $false } catch {} } }
    }
})

# List Calendars Button
$btnListCalendars.Add_Click({
    $targetUserInput = $txtTargetMailbox.Text.Trim()
    $resolvedIdentity = Resolve-MailboxIdentity -IdentityInput $targetUserInput
    if (-not $resolvedIdentity) {
        $script:CurrentUserPrimarySmtp = $null
        if ($lstTargetCalendar) { $lstTargetCalendar.ItemsSource = $null; $lstTargetCalendar.IsEnabled = $false }
        $actionControls = @($txtDelegateUser, $cmbPermissionLevel, $btnViewPermissions, $btnAddSetPermission, $btnRemovePermission)
        $actionControls | ForEach-Object { if($_) { try { $_.IsEnabled = $false } catch {} } }
        return
    }
    Get-CalendarFoldersForUser -UserIdentity $resolvedIdentity
})

# Browse Mailboxes Button
$btnBrowseMailboxes.Add_Click({
    $selectedEmail = Show-MailboxBrowserDialog
    if (-not [string]::IsNullOrWhiteSpace($selectedEmail)) {
        $txtTargetMailbox.Text = $selectedEmail
        Write-OutputToGui "Selected mailbox '$selectedEmail' from browser."
    } else {
        Write-OutputToGui "Mailbox browser cancelled or no selection made."
    }
})

# View Permissions Button
$btnViewPermissions.Add_Click({
    $targetCalendarItem = $lstTargetCalendar.SelectedItem # Get from ListView
    if (-not $targetCalendarItem) { [System.Windows.MessageBox]::Show("Please select the primary '\\Calendar' folder from the list.", "Selection Required", "OK", "Warning"); return }
    if ([string]::IsNullOrWhiteSpace($script:CurrentUserPrimarySmtp)) { Write-OutputToGui "Error: Resolved user email context missing..." -Type Error; [System.Windows.MessageBox]::Show("Target user email context missing...", "Error", "OK", "Error"); return }

    ${calendarDisplayName} = $targetCalendarItem.DisplayName
    ${isPrimaryCalendar} = $targetCalendarItem.IsPrimary

    if (-not $isPrimaryCalendar) {
         [System.Windows.MessageBox]::Show("Managing permissions on non-primary calendars ('${calendarDisplayName}') is not currently supported due to cmdlet limitations. Please select the primary '\\Calendar' folder.", "Unsupported Folder", "OK", "Information")
         Write-OutputToGui "Action skipped: Managing non-primary calendar '${calendarDisplayName}' is not supported." -Type Warning
         return
    }

    # Construct the ID for the primary calendar using DisplayName (e.g., "\Calendar")
    $permissionIdentity = "$($script:CurrentUserPrimarySmtp):$($calendarDisplayName)"
    Write-OutputToGui "Getting permissions for primary calendar: ${calendarDisplayName} (Using ID: ${permissionIdentity})"
    try {
        $permissions = Get-MailboxFolderPermission -Identity $permissionIdentity -ErrorAction Stop | Select-Object User, AccessRights, SharingPermissionFlags
        if ($permissions) {
            Write-OutputToGui "Current Permissions for ${calendarDisplayName}:"
            $permissions | Format-Table -AutoSize | Out-String | ForEach-Object { Write-OutputToGui $_ }
        } else { Write-OutputToGui "No explicit non-default permissions found for ${calendarDisplayName}." }
    } catch { Write-OutputToGui "Error getting permissions for '${calendarDisplayName}' using ID ('${permissionIdentity}'): $($_.Exception.Message)" -Type Error }
})

# Add/Set Permission Button
$btnAddSetPermission.Add_Click({
    $targetCalendarItem = $lstTargetCalendar.SelectedItem # Get from ListView
    ${delegateUser} = $txtDelegateUser.Text.Trim()
    ${permissionLevel} = $cmbPermissionLevel.SelectedItem.Content

    # Check required inputs
    if (-not $targetCalendarItem) { [System.Windows.MessageBox]::Show("Please select the primary '\\Calendar' folder first.", "Selection Required", "OK", "Warning"); return }
    if ([string]::IsNullOrWhiteSpace($script:CurrentUserPrimarySmtp)) { Write-OutputToGui "Error: Resolved user email context is missing..." -Type Error; [System.Windows.MessageBox]::Show("Target user email context missing...", "Error", "OK", "Error"); return }
    if ([string]::IsNullOrWhiteSpace(${delegateUser})) { [System.Windows.MessageBox]::Show("Please enter the Delegate User.", "Input Required", "OK", "Warning"); return }
    if ([string]::IsNullOrWhiteSpace(${permissionLevel}) -or ${permissionLevel} -eq "None") { [System.Windows.MessageBox]::Show("Please select a valid permission level (not 'None')...", "Input Required", "OK", "Warning"); return }

    # Get DisplayName and IsPrimary property
    ${calendarDisplayName} = $targetCalendarItem.DisplayName
    ${isPrimaryCalendar} = $targetCalendarItem.IsPrimary

    # Check if the selected folder is the primary calendar
    if (-not $isPrimaryCalendar) {
         [System.Windows.MessageBox]::Show("Managing permissions on non-primary calendars ('${calendarDisplayName}') is not currently supported...", "Unsupported Folder", "OK", "Information")
         Write-OutputToGui "Action skipped: Managing non-primary calendar '${calendarDisplayName}' is not supported." -Type Warning
         return
    }

    # Construct the ID for the primary calendar
    $permissionIdentity = "$($script:CurrentUserPrimarySmtp):$($calendarDisplayName)"
    Write-OutputToGui "Attempting to set permission '${permissionLevel}' for user '${delegateUser}' on primary calendar '${calendarDisplayName}' (Using ID: ${permissionIdentity})..."

    try {
        # Attempt Set first
        Write-OutputToGui "Trying Set-MailboxFolderPermission..."
        Set-MailboxFolderPermission -Identity $permissionIdentity -User ${delegateUser} -AccessRights ${permissionLevel} -ErrorAction Stop
        Write-OutputToGui "Successfully MODIFIED permission for ${delegateUser} to ${permissionLevel} on '${calendarDisplayName}'."
    } catch {
        # If Set failed, check the specific error message
        ${ErrorMessage} = $_.Exception.Message
        Write-OutputToGui "Set-MailboxFolderPermission failed: ${ErrorMessage}" -Type Warning

        # Corrected Error Check: Check if the error is specifically "no existing permission entry found"
        if (${ErrorMessage} -like "*no existing permission entry found for user*") {
            Write-OutputToGui "User permission not found, trying Add-MailboxFolderPermission..."
            try {
                # Attempt Add if Set failed because user wasn't present
                Add-MailboxFolderPermission -Identity $permissionIdentity -User ${delegateUser} -AccessRights ${permissionLevel} -ErrorAction Stop
                Write-OutputToGui "Successfully ADDED permission for ${delegateUser} as ${permissionLevel} on '${calendarDisplayName}'."
            } catch {
                # Catch errors specifically from the Add attempt
                Write-OutputToGui "Add-MailboxFolderPermission also failed: $($_.Exception.Message)" -Type Error
            }
        } else {
            # Set failed for a different reason
            Write-OutputToGui "An unexpected error occurred during Set-MailboxFolderPermission: ${ErrorMessage}" -Type Error
        }
    }
})

# Remove Permission Button
$btnRemovePermission.Add_Click({
    $targetCalendarItem = $lstTargetCalendar.SelectedItem # Get from ListView
    ${delegateUser} = $txtDelegateUser.Text.Trim()

    # Check required inputs
    if (-not $targetCalendarItem) { [System.Windows.MessageBox]::Show("Please select the primary '\\Calendar' folder first.", "Selection Required", "OK", "Warning"); return }
    if ([string]::IsNullOrWhiteSpace($script:CurrentUserPrimarySmtp)) { Write-OutputToGui "Error: Resolved user email context is missing..." -Type Error; [System.Windows.MessageBox]::Show("Target user email context missing...", "Error", "OK", "Error"); return }
    if ([string]::IsNullOrWhiteSpace(${delegateUser})) { [System.Windows.MessageBox]::Show("Please enter the Delegate User to remove.", "Input Required", "OK", "Warning"); return }

    # Get DisplayName and IsPrimary property
    ${calendarDisplayName} = $targetCalendarItem.DisplayName
    ${isPrimaryCalendar} = $targetCalendarItem.IsPrimary

    # Check if the selected folder is the primary calendar
    if (-not $isPrimaryCalendar) {
         [System.Windows.MessageBox]::Show("Managing permissions on non-primary calendars ('${calendarDisplayName}') is not currently supported...", "Unsupported Folder", "OK", "Information")
         Write-OutputToGui "Action skipped: Managing non-primary calendar '${calendarDisplayName}' is not supported." -Type Warning
         return
    }

    # Confirmation Dialog
    $confirmResult = [System.Windows.MessageBox]::Show("Are you sure you want to remove permissions for '${delegateUser}' from the primary calendar ('${calendarDisplayName}')?", "Confirm Removal", "YesNo", "Question")
    if ($confirmResult -ne 'Yes') { Write-OutputToGui "Permission removal cancelled by user."; return }

    # Construct the ID for the primary calendar
    $permissionIdentity = "$($script:CurrentUserPrimarySmtp):$($calendarDisplayName)"
    Write-OutputToGui "Attempting to remove permission for user '${delegateUser}' from primary calendar '${calendarDisplayName}' (Using ID: ${permissionIdentity})..."

    try {
        # Attempt Remove using the simple ID for the primary calendar
        Remove-MailboxFolderPermission -Identity $permissionIdentity -User ${delegateUser} -Confirm:$false -ErrorAction Stop
        Write-OutputToGui "Successfully removed permission for ${delegateUser} from '${calendarDisplayName}'."
    } catch {
         if ($_.Exception.Message -like "*couldn't be found on folder*") {
             Write-OutputToGui "User '${delegateUser}' may not have had permissions on '${calendarDisplayName}'. Error: $($_.Exception.Message)" -Type Warning
         } else {
             Write-OutputToGui "Error removing permission for '${delegateUser}' from '${calendarDisplayName}' (ID: ${permissionIdentity}): $($_.Exception.Message)" -Type Error
         }
    }
})

# Close Button
$btnClose.Add_Click({
    if ($Global:IsConnected) {
        Write-OutputToGui "Disconnecting on close..."
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    $window.Close()
})

# Window Closing Event Handler
$window.Add_Closing({
    param($sender, $e)
    if ($Global:IsConnected) {
        Write-Host "Disconnecting from Exchange Online due to window close..." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
})

#endregion

#region Initialize GUI State and Show Window
# Set initial state to disconnected
Set-GuiState $false
# Module check runs before this, GUI only shown if module is ready
Write-Host "Launching M365 Primary Calendar Permissions Manager (V4 / 1.8 Formatted)..."
# Show the main window modally
$window.ShowDialog() | Out-Null
Write-Host "GUI Closed."
#endregion