#Requires -Version 5.0

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

. (Join-Path $PSScriptRoot 'Helper.Core.ps1')

$script:guiConfigPath = Join-Path $PSScriptRoot 'yt-dlp-helper.gui.config.json'
$script:guiLogDir = Join-Path $PSScriptRoot 'logs'
$script:guiJobLogDir = Join-Path $script:guiLogDir 'jobs'
$script:guiLogPath = Join-Path $script:guiLogDir 'yt-dlp-helper-gui.log'
$script:workerScriptPath = Join-Path $PSScriptRoot 'yt-dlp-helper-worker.ps1'
$script:hostExecutable = try { (Get-Process -Id $PID -ErrorAction Stop).Path } catch { 'powershell.exe' }
$script:activeJob = $null
$script:jobPollTimer = $null
$script:jobHistory = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:currentActivityLines = New-Object System.Collections.Generic.List[string]

function Get-GuiRetentionDays {
    $retentionDays = 14
    if (-not (Test-Path -LiteralPath $script:guiConfigPath -PathType Leaf)) {
        return $retentionDays
    }

    try {
        $config = Get-Content -LiteralPath $script:guiConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        if ($config.logging -and $config.logging.retentionDays -ge 0) {
            $retentionDays = [int]$config.logging.retentionDays
        }
    }
    catch {
    }

    return $retentionDays
}

function Trim-GuiLogs {
    $retentionDays = Get-GuiRetentionDays
    if ($retentionDays -le 0) {
        return
    }

    $cutoff = (Get-Date).AddDays(-$retentionDays)
    foreach ($dir in @($script:guiLogDir, $script:guiJobLogDir)) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
            continue
        }

        Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.LastWriteTime -lt $cutoff) {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

if (-not (Test-Path -LiteralPath $script:guiLogDir -PathType Container)) {
    New-Item -ItemType Directory -Path $script:guiLogDir -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $script:guiJobLogDir -PathType Container)) {
    New-Item -ItemType Directory -Path $script:guiJobLogDir -Force | Out-Null
}

Trim-GuiLogs
Initialize-HelperCore -Mode gui -LoggingConfigPath $script:guiConfigPath -LogFilePath $script:guiLogPath
Set-HelperLogHandler -Handler { param($Entry) }
Write-HelperLog -Level 'INFO' -Message 'Helper GUI session started.'

function Get-ClipboardYoutubeUrl {
    try {
        if (-not [System.Windows.Clipboard]::ContainsText()) {
            return ''
        }

        $text = [System.Windows.Clipboard]::GetText()
        Assert-ValidYoutubeUrl -Url $text
        return $text
    }
    catch {
        return ''
    }
}

function Build-EnvironmentRows {
    param([object]$Snapshot)

    if ($null -eq $Snapshot) {
        return @()
    }

    return @(
        (New-HelperFieldRow -Field 'Download Root' -Value ([string]$Snapshot.DownloadRoot)),
        (New-HelperFieldRow -Field 'yt-dlp Version' -Value ([string]$Snapshot.YtDlpVersion)),
        (New-HelperFieldRow -Field 'yt-dlp Path' -Value ([string]$Snapshot.YtDlpPath)),
        (New-HelperFieldRow -Field 'FFmpeg Bin' -Value ([string]$Snapshot.FfmpegBinPath)),
        (New-HelperFieldRow -Field 'JS Runtime' -Value ([string]$Snapshot.JsRuntimeLabel)),
        (New-HelperFieldRow -Field 'Firefox Profile' -Value ([string]$Snapshot.FirefoxProfilePath)),
        (New-HelperFieldRow -Field 'Auth Ready' -Value ($(if ($Snapshot.AuthReady) { 'Yes' } else { 'No' }))),
        (New-HelperFieldRow -Field 'Playlist Cache Age' -Value ([string]$Snapshot.PlaylistCacheAge)),
        (New-HelperFieldRow -Field 'Treat Warnings Preferred' -Value ([string]$Snapshot.TreatWarningsPreferred))
    )
}

function Build-MyPlaylistRows {
    param([object[]]$Rows)

    if ($null -eq $Rows) {
        return @()
    }

    return @($Rows)
}

function Normalize-ActivityText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    return @((([string]$Text -replace "`r`n", "`n") -replace "`r", "`n") -split "`n")
}

function Add-ActivityText {
    param([string]$Text)

    foreach ($line in (Normalize-ActivityText -Text $Text)) {
        if ($null -eq $line) {
            continue
        }
        $script:currentActivityLines.Add($line)
    }

    while ($script:currentActivityLines.Count -gt 600) {
        $script:currentActivityLines.RemoveAt(0)
    }

    $ActivityTextBox.Text = ($script:currentActivityLines -join [Environment]::NewLine)
    $ActivityTextBox.ScrollToEnd()
}

function Reset-ActivityText {
    $script:currentActivityLines = New-Object System.Collections.Generic.List[string]
    $ActivityTextBox.Text = ''
}

function Set-StatusMessage {
    param(
        [string]$Message,
        [string]$Color = 'Black'
    )

    $StatusTextBlock.Text = $Message
    try {
        $StatusTextBlock.Foreground = [System.Windows.Media.Brushes]::$Color
    }
    catch {
        $StatusTextBlock.Foreground = [System.Windows.Media.Brushes]::Black
    }
}

function Get-GuiStateValue {
    param(
        [hashtable]$GuiState,
        [string]$Name,
        $Default = $null
    )

    if ($null -eq $GuiState -or -not $GuiState.Contains($Name)) {
        return $Default
    }

    return $GuiState[$Name]
}

$script:guiState = Get-HelperGuiState
$script:initialSnapshot = Get-HelperLocalEnvironmentSnapshot
$preferredProfilePath = Get-PreferredFirefoxProfilePath
$initialUrl = Get-ClipboardYoutubeUrl
$defaultDownloadRoot = Get-HelperDefaultDownloadRoot

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="yt-dlp Helper Workbench"
        Height="920"
        Width="1540"
        MinHeight="760"
        MinWidth="1180"
        WindowStartupLocation="CenterScreen"
        Background="#F4EEE3">
    <Window.Resources>
        <SolidColorBrush x:Key="PanelBrush" Color="#FFF9F3"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#D5C8B4"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#9F4D1D"/>
        <SolidColorBrush x:Key="AccentSoftBrush" Color="#EBD7C4"/>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,12"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Padding" Value="12,6"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="0,0,8,0"/>
        </Style>
    </Window.Resources>
    <Grid Margin="14">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="2.1*"/>
            <ColumnDefinition Width="1.1*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.ColumnSpan="2" Background="#1F4E5F" CornerRadius="14" Padding="16" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <TextBlock Text="yt-dlp Helper Workbench" FontSize="28" FontWeight="SemiBold" Foreground="White"/>
                    <TextBlock Text="Single video, playlist URL, saved playlists, environment prep, and offline playlist-folder inspection." Margin="0,6,0,0" Foreground="#DCE9EE" FontSize="14"/>
                </StackPanel>
                <Border Grid.Column="1" Background="#F1E3D3" CornerRadius="10" Padding="10,6" VerticalAlignment="Center">
                    <TextBlock x:Name="HeaderHintText" Text="One active job at a time. Live logs stay visible." Foreground="#6A3514" FontWeight="SemiBold"/>
                </Border>
            </Grid>
        </Border>

        <GroupBox Grid.Row="1" Grid.Column="0" Header="Session Context">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Text="Download Root" Margin="0,6,10,0" FontWeight="SemiBold"/>
                    <TextBox Grid.Column="1" x:Name="DownloadRootBox" Height="30"/>
                    <Button Grid.Column="2" x:Name="BrowseDownloadRootButton" Content="Browse..." Height="30"/>
                    <Button Grid.Column="3" x:Name="OpenDownloadRootButton" Content="Open" Height="30"/>
                </Grid>
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="300"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Text="Firefox Profile" Margin="0,6,10,0" FontWeight="SemiBold"/>
                    <ComboBox Grid.Column="1" x:Name="ProfileCombo" Height="30" DisplayMemberPath="Label"/>
                    <Button Grid.Column="2" x:Name="BrowseProfileButton" Content="Browse..." Height="30"/>
                    <CheckBox Grid.Column="3" x:Name="UseCookiesCheck" Content="Use cookies for URL tasks" Margin="10,7,12,0"/>
                    <CheckBox Grid.Column="4" x:Name="TreatWarningsCheck" Content="Treat non-zero exit as warning" Margin="0,7,12,0"/>
                    <TextBlock Grid.Column="5" x:Name="ProfileStatusText" Margin="0,7,0,0" Foreground="#72503A"/>
                </Grid>
            </Grid>
        </GroupBox>

        <TabControl Grid.Row="2" Grid.Column="0" x:Name="TaskTabs" Margin="0,0,12,0" Background="#FFF9F3">
            <TabItem Header="Single Video">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Text="Paste a YouTube video URL. The helper uses the current best-quality preset and your saved download context." TextWrapping="Wrap" Margin="0,0,0,12"/>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="SingleVideoUrlBox" Height="34"/>
                        <Button Grid.Column="1" x:Name="UseClipboardForSingleButton" Content="Use Clipboard" Height="34"/>
                        <Button Grid.Column="2" x:Name="StartSingleVideoButton" Content="Download Video" Height="34"/>
                    </Grid>
                </Grid>
            </TabItem>
            <TabItem Header="Playlist URL">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Text="Download a playlist by URL, optionally followed by a sidecar metadata/subtitle pass into _sidecar." TextWrapping="Wrap" Margin="0,0,0,12"/>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="PlaylistUrlBox" Height="34"/>
                        <Button Grid.Column="1" x:Name="UseClipboardForPlaylistButton" Content="Use Clipboard" Height="34"/>
                        <Button Grid.Column="2" x:Name="StartPlaylistButton" Content="Download Playlist" Height="34"/>
                    </Grid>
                    <CheckBox Grid.Row="2" x:Name="FetchPlaylistSidecarsCheck" Content="After the media pass, also fetch info json and subtitles into _sidecar" Margin="0,12,0,0"/>
                </Grid>
            </TabItem>
            <TabItem Header="My Playlists">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Text="Load your playlist feed with Firefox cookies, browse the list, and launch the selected playlist download." TextWrapping="Wrap" Margin="0,0,0,12"/>
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="LoadMyPlaylistsButton" Content="Load Cached or Fresh" Height="32"/>
                        <Button x:Name="RefreshMyPlaylistsButton" Content="Force Refresh" Height="32"/>
                        <CheckBox x:Name="FetchSelectedPlaylistSidecarsCheck" Content="Fetch sidecars after download" Margin="10,7,0,0"/>
                        <TextBlock x:Name="MyPlaylistsSourceText" Margin="16,7,0,0" Foreground="#72503A"/>
                    </StackPanel>
                    <DataGrid Grid.Row="2" x:Name="MyPlaylistsGrid" IsReadOnly="True" AutoGenerateColumns="False" HeadersVisibility="Column" CanUserAddRows="False" CanUserDeleteRows="False">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="#" Binding="{Binding Index}" Width="60"/>
                            <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="*"/>
                            <DataGridTextColumn Header="Playlist ID" Binding="{Binding PlaylistId}" Width="260"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,10,0,0">
                        <Button x:Name="DownloadSelectedPlaylistButton" Content="Download Selected Playlist" Height="34"/>
                        <TextBlock x:Name="SelectedPlaylistText" Margin="12,8,0,0" Foreground="#72503A"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            <TabItem Header="Environment">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="ReloadEnvironmentButton" Content="Reload Local Status" Height="32"/>
                        <Button x:Name="PrepareEnvironmentButton" Content="Prepare / Update Shared Tools" Height="32"/>
                        <Button x:Name="OpenHelperFolderButton" Content="Open Helper Folder" Height="32"/>
                        <Button x:Name="OpenLogsFolderButton" Content="Open Logs Folder" Height="32"/>
                    </StackPanel>
                    <DataGrid Grid.Row="1" x:Name="EnvironmentGrid" IsReadOnly="True" AutoGenerateColumns="False" HeadersVisibility="Column" CanUserAddRows="False" CanUserDeleteRows="False">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Field" Binding="{Binding Field}" Width="220"/>
                            <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>
            <TabItem Header="Folder Inspector">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="120"/>
                    </Grid.RowDefinitions>
                    <TextBlock Text="Inspect an existing playlist folder offline. The workbench reports only what current local artifacts can actually prove or infer." TextWrapping="Wrap" Margin="0,0,0,12"/>
                    <Grid Grid.Row="1" Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <ComboBox x:Name="InspectorFolderCombo" Height="34" IsEditable="True"/>
                        <Button Grid.Column="1" x:Name="BrowseInspectorFolderButton" Content="Browse..." Height="34"/>
                        <Button Grid.Column="2" x:Name="InspectFolderButton" Content="Inspect Folder" Height="34"/>
                        <Button Grid.Column="3" x:Name="OpenInspectorFolderButton" Content="Open" Height="34"/>
                    </Grid>
                    <DataGrid Grid.Row="2" x:Name="InspectorGrid" IsReadOnly="True" AutoGenerateColumns="False" HeadersVisibility="Column" CanUserAddRows="False" CanUserDeleteRows="False">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Field" Binding="{Binding Field}" Width="220"/>
                            <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <TextBlock Grid.Row="3" Text="Inspector Notes" FontWeight="SemiBold" Margin="0,10,0,6"/>
                    <TextBox Grid.Row="4" x:Name="InspectorNotesBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"/>
                </Grid>
            </TabItem>
        </TabControl>

        <GroupBox Grid.Row="1" Grid.RowSpan="2" Grid.Column="1" Header="Activity and History">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="280"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock x:Name="CurrentJobText" Text="Idle." FontWeight="SemiBold" Foreground="#6A3514"/>
                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,10">
                    <Button x:Name="CancelJobButton" Content="Cancel Current Job" Height="32" IsEnabled="False"/>
                    <TextBlock x:Name="CurrentJobStatusText" Margin="8,8,0,0" Foreground="#72503A"/>
                </StackPanel>
                <TextBox Grid.Row="2" x:Name="ActivityTextBox" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"/>
                <TextBlock Grid.Row="3" Text="Recent Jobs" FontWeight="SemiBold" Margin="0,10,0,6"/>
                <DataGrid Grid.Row="4" x:Name="RecentJobsGrid" IsReadOnly="True" AutoGenerateColumns="False" HeadersVisibility="Column" CanUserAddRows="False" CanUserDeleteRows="False">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Started" Binding="{Binding StartedDisplay}" Width="135"/>
                        <DataGridTextColumn Header="Task" Binding="{Binding Task}" Width="140"/>
                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="120"/>
                        <DataGridTextColumn Header="Summary" Binding="{Binding Summary}" Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </GroupBox>

        <TextBlock Grid.Row="3" Grid.ColumnSpan="2" x:Name="StatusTextBlock" Margin="0,12,0,0" Text="Ready. Configure the session context and start a task."/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$HeaderHintText = $window.FindName('HeaderHintText')
$DownloadRootBox = $window.FindName('DownloadRootBox')
$BrowseDownloadRootButton = $window.FindName('BrowseDownloadRootButton')
$OpenDownloadRootButton = $window.FindName('OpenDownloadRootButton')
$ProfileCombo = $window.FindName('ProfileCombo')
$BrowseProfileButton = $window.FindName('BrowseProfileButton')
$UseCookiesCheck = $window.FindName('UseCookiesCheck')
$TreatWarningsCheck = $window.FindName('TreatWarningsCheck')
$ProfileStatusText = $window.FindName('ProfileStatusText')
$TaskTabs = $window.FindName('TaskTabs')
$SingleVideoUrlBox = $window.FindName('SingleVideoUrlBox')
$UseClipboardForSingleButton = $window.FindName('UseClipboardForSingleButton')
$StartSingleVideoButton = $window.FindName('StartSingleVideoButton')
$PlaylistUrlBox = $window.FindName('PlaylistUrlBox')
$UseClipboardForPlaylistButton = $window.FindName('UseClipboardForPlaylistButton')
$StartPlaylistButton = $window.FindName('StartPlaylistButton')
$FetchPlaylistSidecarsCheck = $window.FindName('FetchPlaylistSidecarsCheck')
$LoadMyPlaylistsButton = $window.FindName('LoadMyPlaylistsButton')
$RefreshMyPlaylistsButton = $window.FindName('RefreshMyPlaylistsButton')
$FetchSelectedPlaylistSidecarsCheck = $window.FindName('FetchSelectedPlaylistSidecarsCheck')
$MyPlaylistsSourceText = $window.FindName('MyPlaylistsSourceText')
$MyPlaylistsGrid = $window.FindName('MyPlaylistsGrid')
$DownloadSelectedPlaylistButton = $window.FindName('DownloadSelectedPlaylistButton')
$SelectedPlaylistText = $window.FindName('SelectedPlaylistText')
$ReloadEnvironmentButton = $window.FindName('ReloadEnvironmentButton')
$PrepareEnvironmentButton = $window.FindName('PrepareEnvironmentButton')
$OpenHelperFolderButton = $window.FindName('OpenHelperFolderButton')
$OpenLogsFolderButton = $window.FindName('OpenLogsFolderButton')
$EnvironmentGrid = $window.FindName('EnvironmentGrid')
$InspectorFolderCombo = $window.FindName('InspectorFolderCombo')
$BrowseInspectorFolderButton = $window.FindName('BrowseInspectorFolderButton')
$InspectFolderButton = $window.FindName('InspectFolderButton')
$OpenInspectorFolderButton = $window.FindName('OpenInspectorFolderButton')
$InspectorGrid = $window.FindName('InspectorGrid')
$InspectorNotesBox = $window.FindName('InspectorNotesBox')
$CurrentJobText = $window.FindName('CurrentJobText')
$CancelJobButton = $window.FindName('CancelJobButton')
$CurrentJobStatusText = $window.FindName('CurrentJobStatusText')
$ActivityTextBox = $window.FindName('ActivityTextBox')
$RecentJobsGrid = $window.FindName('RecentJobsGrid')
$StatusTextBlock = $window.FindName('StatusTextBlock')

$RecentJobsGrid.ItemsSource = $script:jobHistory
$EnvironmentGrid.ItemsSource = @(Build-EnvironmentRows -Snapshot $script:initialSnapshot)
$DownloadRootBox.Text = $defaultDownloadRoot
$SingleVideoUrlBox.Text = $initialUrl
$PlaylistUrlBox.Text = $initialUrl
$UseCookiesCheck.IsChecked = $false
$TreatWarningsCheck.IsChecked = [bool]$script:helperState.treatYtDlpErrorsAsWarningsPreferred
$FetchPlaylistSidecarsCheck.IsChecked = $false
$FetchSelectedPlaylistSidecarsCheck.IsChecked = $false
$MyPlaylistsGrid.ItemsSource = @()
$InspectorGrid.ItemsSource = @()
$InspectorNotesBox.Text = ''

if ($script:guiState.Contains('windowWidth')) { $window.Width = [double]$script:guiState.windowWidth }
if ($script:guiState.Contains('windowHeight')) { $window.Height = [double]$script:guiState.windowHeight }
if ($script:guiState.Contains('windowLeft')) { $window.Left = [double]$script:guiState.windowLeft }
if ($script:guiState.Contains('windowTop')) { $window.Top = [double]$script:guiState.windowTop }

function Set-SelectedProfilePath {
    param([string]$TargetPath)

    $ProfileCombo.Items.Clear()
    foreach ($candidate in @(Get-FirefoxProfileCandidates)) {
        $label = if ($candidate.Name -and $candidate.Name -ne $candidate.ProfileName) {
            "$($candidate.Name) [$($candidate.ProfileName)]"
        }
        else {
            $candidate.ProfileName
        }

        [void]$ProfileCombo.Items.Add([pscustomobject]@{
            Label = $label
            Path  = [string]$candidate.Path
        })
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetPath) -and -not (@($ProfileCombo.Items) | Where-Object { $_.Path -ieq $TargetPath })) {
        [void]$ProfileCombo.Items.Add([pscustomobject]@{
            Label = Split-Path -Leaf $TargetPath
            Path  = $TargetPath
        })
    }

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        $ProfileCombo.SelectedIndex = -1
        return
    }

    foreach ($item in @($ProfileCombo.Items)) {
        if ($item.Path -ieq $TargetPath) {
            $ProfileCombo.SelectedItem = $item
            return
        }
    }
}

function Get-SelectedProfilePath {
    if ($ProfileCombo.SelectedItem -and $ProfileCombo.SelectedItem.Path) {
        return [string]$ProfileCombo.SelectedItem.Path
    }

    return ''
}

function Update-ProfileStatus {
    $selectedPath = Get-SelectedProfilePath
    $validation = Test-FirefoxProfilePath -ProfilePath $selectedPath
    if ($validation.IsValid) {
        $ProfileStatusText.Text = "Ready: $($validation.ProfilePath)"
        Save-HelperPreferences -FirefoxProfilePath $validation.ProfilePath
    }
    elseif ([string]::IsNullOrWhiteSpace($selectedPath)) {
        $ProfileStatusText.Text = 'No Firefox profile selected.'
    }
    else {
        $ProfileStatusText.Text = $validation.StatusMessage
    }
}

function Update-SelectedPlaylistText {
    $row = $MyPlaylistsGrid.SelectedItem
    if ($null -eq $row) {
        $SelectedPlaylistText.Text = 'No playlist selected.'
        return
    }

    $SelectedPlaylistText.Text = "Selected: $($row.Title)"
}

function Load-ActivityFromJob {
    param([object]$JobRecord)

    if ($null -eq $JobRecord) {
        return
    }

    Reset-ActivityText
    Add-ActivityText -Text (Read-TextFileSafe -Path $JobRecord.StdOutPath)
    Add-ActivityText -Text (Read-TextFileSafe -Path $JobRecord.StdErrPath)
}

function Set-ActiveJobUiState {
    param(
        [bool]$IsBusy,
        [string]$CurrentText,
        [string]$StatusText
    )

    $CancelJobButton.IsEnabled = $IsBusy
    $CurrentJobText.Text = $CurrentText
    $CurrentJobStatusText.Text = $StatusText
    $window.Cursor = if ($IsBusy) { [System.Windows.Input.Cursors]::Wait } else { [System.Windows.Input.Cursors]::Arrow }
}

function Build-WorkerRequest {
    param(
        [string]$Kind,
        [hashtable]$Extra
    )

    $base = [ordered]@{
        Kind                    = $Kind
        DownloadRoot            = $DownloadRootBox.Text.Trim()
        FirefoxProfilePath      = (Get-SelectedProfilePath)
        UseCookies              = [bool]$UseCookiesCheck.IsChecked
        TreatWarningsAsNonFatal = [bool]$TreatWarningsCheck.IsChecked
    }

    foreach ($key in $Extra.Keys) {
        $base[$key] = $Extra[$key]
    }

    return $base
}

function Start-HelperJob {
    param(
        [string]$TaskName,
        [hashtable]$Request
    )

    if ($null -ne $script:activeJob) {
        Set-StatusMessage -Message 'A helper job is already running.' -Color 'DarkOrange'
        return
    }

    $jobId = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $jobPrefix = Join-Path $script:guiJobLogDir $jobId
    $requestPath = "$jobPrefix.request.json"
    $resultPath = "$jobPrefix.result.json"
    $stdoutPath = "$jobPrefix.stdout.log"
    $stderrPath = "$jobPrefix.stderr.log"

    $Request | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $requestPath -Encoding UTF8 -Force

    $jobRecord = [pscustomobject]@{
        Task          = $TaskName
        Status        = 'Running'
        Summary       = 'Job started.'
        Started       = Get-Date
        StartedDisplay = (Get-Date -Format 'yyyy-MM-dd HH:mm')
        StdOutPath    = $stdoutPath
        StdErrPath    = $stderrPath
        ResultPath    = $resultPath
    }
    $script:jobHistory.Insert(0, $jobRecord)
    $RecentJobsGrid.Items.Refresh()

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $script:workerScriptPath,
        '-RequestPath', $requestPath,
        '-ResultPath', $resultPath,
        '-GuiConfigPath', $script:guiConfigPath
    )

    $process = Start-Process -FilePath $script:hostExecutable -ArgumentList $arguments -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -PassThru -WindowStyle Hidden
    $script:activeJob = [pscustomobject]@{
        TaskName              = $TaskName
        Request               = $Request
        RequestPath           = $requestPath
        ResultPath            = $resultPath
        StdOutPath            = $stdoutPath
        StdErrPath            = $stderrPath
        StdOutLength          = 0
        StdErrLength          = 0
        Process               = $process
        JobRecord             = $jobRecord
        CancellationRequested = $false
    }

    Reset-ActivityText
    Set-ActiveJobUiState -IsBusy $true -CurrentText $TaskName -StatusText 'Running'
    Set-StatusMessage -Message ("Started: {0}" -f $TaskName) -Color 'DarkBlue'
    Write-HelperLog -Level 'INFO' -Message ("Started worker job '{0}' with PID {1}." -f $TaskName, $process.Id)
    $script:jobPollTimer.Start()
}

function Read-JobOutputChunk {
    param(
        [string]$Path,
        [ref]$LengthRef
    )

    $text = Read-TextFileSafe -Path $Path
    if ([string]::IsNullOrEmpty($text)) {
        $LengthRef.Value = 0
        return ''
    }

    if ($text.Length -lt [int]$LengthRef.Value) {
        $LengthRef.Value = 0
    }

    $chunk = $text.Substring([int]$LengthRef.Value)
    $LengthRef.Value = $text.Length
    return $chunk
}

function Apply-WorkerResult {
    param([object]$Result)

    if ($null -eq $Result) {
        return
    }

    switch ([string]$Result.Kind) {
        'my-playlists' {
            $MyPlaylistsGrid.ItemsSource = @(Build-MyPlaylistRows -Rows $Result.Data.Playlists)
            $MyPlaylistsSourceText.Text = $Result.Status
            $SelectedPlaylistText.Text = 'Select a playlist row to download.'
            if ($script:helperState.lastPlaylistId) {
                foreach ($row in @($MyPlaylistsGrid.ItemsSource)) {
                    if ($row.PlaylistId -eq [string]$script:helperState.lastPlaylistId) {
                        $MyPlaylistsGrid.SelectedItem = $row
                        break
                    }
                }
            }
            Update-SelectedPlaylistText
        }
        'environment' {
            $EnvironmentGrid.ItemsSource = @(Build-EnvironmentRows -Snapshot $Result.Data.Snapshot)
            $HeaderHintText.Text = 'Shared tools checked and ready.'
        }
        'folder-inspector' {
            $InspectorGrid.ItemsSource = @($Result.Data.SummaryRows)
            $InspectorNotesBox.Text = ((@($Result.Data.Notes) + @('') + @($Result.Data.Sources | ForEach-Object { "Source: $_" })) -join [Environment]::NewLine).Trim()
            Add-HelperRecentInspectorFolder -FolderPath ([string]$Result.Data.FolderPath)
        }
    }
}

function Complete-ActiveJob {
    if ($null -eq $script:activeJob) {
        return
    }

    $result = $null
    if (Test-Path -LiteralPath $script:activeJob.ResultPath -PathType Leaf) {
        try {
            $result = Get-Content -LiteralPath $script:activeJob.ResultPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
        }
    }

    $jobRecord = $script:activeJob.JobRecord
    $jobRecord.Status = if ($script:activeJob.CancellationRequested) {
        'Cancelled'
    }
    elseif ($null -ne $result -and $result.Succeeded) {
        'Succeeded'
    }
    else {
        'Failed'
    }
    $jobRecord.Summary = if ($null -ne $result) { [string]$result.Summary } elseif ($script:activeJob.CancellationRequested) { 'Cancelled by user.' } else { 'Worker ended without a result file.' }
    $RecentJobsGrid.Items.Refresh()

    if ($null -ne $result) {
        Apply-WorkerResult -Result $result
        Set-StatusMessage -Message ([string]$result.Status) -Color $(if ($result.Succeeded) { 'DarkGreen' } else { 'Red' })
    }
    elseif ($script:activeJob.CancellationRequested) {
        Set-StatusMessage -Message 'The active helper job was cancelled.' -Color 'DarkOrange'
    }
    else {
        Set-StatusMessage -Message 'The worker exited without a usable result file.' -Color 'Red'
    }

    Set-ActiveJobUiState -IsBusy $false -CurrentText 'Idle.' -StatusText 'No active worker'
    $script:jobPollTimer.Stop()
    $script:activeJob = $null
}

function Update-ActiveJob {
    if ($null -eq $script:activeJob) {
        return
    }

    $stdoutChunk = Read-JobOutputChunk -Path $script:activeJob.StdOutPath -LengthRef ([ref]$script:activeJob.StdOutLength)
    $stderrChunk = Read-JobOutputChunk -Path $script:activeJob.StdErrPath -LengthRef ([ref]$script:activeJob.StdErrLength)
    if (-not [string]::IsNullOrWhiteSpace($stdoutChunk)) { Add-ActivityText -Text $stdoutChunk }
    if (-not [string]::IsNullOrWhiteSpace($stderrChunk)) { Add-ActivityText -Text $stderrChunk }

    try { $script:activeJob.Process.Refresh() } catch {}
    if ($script:activeJob.Process.HasExited) {
        Complete-ActiveJob
    }
}

function Cancel-ActiveJob {
    if ($null -eq $script:activeJob) {
        return
    }

    $script:activeJob.CancellationRequested = $true
    Set-StatusMessage -Message 'Cancelling active helper job...' -Color 'DarkOrange'
    try {
        & "$env:SystemRoot\System32\taskkill.exe" /PID $script:activeJob.Process.Id /T /F | Out-Null
    }
    catch {
    }
}

Set-SelectedProfilePath -TargetPath $preferredProfilePath
Update-ProfileStatus
if ($script:guiState.Contains('selectedTabIndex')) { $TaskTabs.SelectedIndex = [int]$script:guiState.selectedTabIndex }
if ($script:guiState.Contains('recentInspectorFolders')) {
    foreach ($path in @($script:guiState.recentInspectorFolders)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$path)) {
            [void]$InspectorFolderCombo.Items.Add([string]$path)
        }
    }
}
if ($InspectorFolderCombo.Items.Count -gt 0) {
    $InspectorFolderCombo.Text = [string]$InspectorFolderCombo.Items[0]
}

$script:jobPollTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:jobPollTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:jobPollTimer.Add_Tick({ Update-ActiveJob })

$BrowseDownloadRootButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select download root'
    if (Test-Path -LiteralPath $DownloadRootBox.Text -PathType Container) {
        $dialog.SelectedPath = $DownloadRootBox.Text
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $DownloadRootBox.Text = $dialog.SelectedPath
        Save-HelperPreferences -DownloadRoot $dialog.SelectedPath
        $EnvironmentGrid.ItemsSource = @(Build-EnvironmentRows -Snapshot (Get-HelperLocalEnvironmentSnapshot))
    }
})

$OpenDownloadRootButton.Add_Click({
    if (Test-Path -LiteralPath $DownloadRootBox.Text -PathType Container) {
        Start-Process explorer.exe $DownloadRootBox.Text | Out-Null
    }
})

$BrowseProfileButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select Firefox profile folder'
    $dialog.ShowNewFolderButton = $false
    $current = Get-SelectedProfilePath
    if (Test-Path -LiteralPath $current -PathType Container) {
        $dialog.SelectedPath = $current
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Set-SelectedProfilePath -TargetPath $dialog.SelectedPath
        Update-ProfileStatus
    }
})

$ProfileCombo.Add_SelectionChanged({ Update-ProfileStatus })
$MyPlaylistsGrid.Add_SelectionChanged({ Update-SelectedPlaylistText })
$UseClipboardForSingleButton.Add_Click({ $SingleVideoUrlBox.Text = Get-ClipboardYoutubeUrl })
$UseClipboardForPlaylistButton.Add_Click({ $PlaylistUrlBox.Text = Get-ClipboardYoutubeUrl })

$StartSingleVideoButton.Add_Click({
    Start-HelperJob -TaskName 'Single Video' -Request (Build-WorkerRequest -Kind 'single-video' -Extra @{
        Url = $SingleVideoUrlBox.Text.Trim()
    })
})

$StartPlaylistButton.Add_Click({
    Start-HelperJob -TaskName 'Playlist URL' -Request (Build-WorkerRequest -Kind 'playlist-url' -Extra @{
        Url           = $PlaylistUrlBox.Text.Trim()
        FetchSidecars = [bool]$FetchPlaylistSidecarsCheck.IsChecked
    })
})

$LoadMyPlaylistsButton.Add_Click({
    Start-HelperJob -TaskName 'Load My Playlists' -Request (Build-WorkerRequest -Kind 'my-playlists' -Extra @{
        UseCookies         = $true
        ForceRefreshCache  = $false
    })
})

$RefreshMyPlaylistsButton.Add_Click({
    Start-HelperJob -TaskName 'Refresh My Playlists' -Request (Build-WorkerRequest -Kind 'my-playlists' -Extra @{
        UseCookies         = $true
        ForceRefreshCache  = $true
    })
})

$DownloadSelectedPlaylistButton.Add_Click({
    $row = $MyPlaylistsGrid.SelectedItem
    if ($null -eq $row) {
        Set-StatusMessage -Message 'Select a playlist row before starting the download.' -Color 'DarkOrange'
        return
    }

    Save-HelperPreferences -PlaylistIndex ([string]$row.Index) -PlaylistId ([string]$row.PlaylistId) -DownloadRoot $DownloadRootBox.Text.Trim()
    Start-HelperJob -TaskName 'Download Selected Playlist' -Request (Build-WorkerRequest -Kind 'playlist-url' -Extra @{
        Url           = [string]$row.Url
        UseCookies    = $true
        FetchSidecars = [bool]$FetchSelectedPlaylistSidecarsCheck.IsChecked
    })
})

$ReloadEnvironmentButton.Add_Click({
    $EnvironmentGrid.ItemsSource = @(Build-EnvironmentRows -Snapshot (Get-HelperLocalEnvironmentSnapshot))
    Set-StatusMessage -Message 'Reloaded local environment status.' -Color 'DarkGreen'
})

$PrepareEnvironmentButton.Add_Click({
    Start-HelperJob -TaskName 'Prepare Environment' -Request (Build-WorkerRequest -Kind 'environment' -Extra @{ })
})

$OpenHelperFolderButton.Add_Click({ Start-Process explorer.exe $PSScriptRoot | Out-Null })
$OpenLogsFolderButton.Add_Click({ Start-Process explorer.exe $script:guiLogDir | Out-Null })

$BrowseInspectorFolderButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select a folder to inspect'
    if (Test-Path -LiteralPath $InspectorFolderCombo.Text -PathType Container) {
        $dialog.SelectedPath = $InspectorFolderCombo.Text
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if (-not (@($InspectorFolderCombo.Items) -contains $dialog.SelectedPath)) {
            [void]$InspectorFolderCombo.Items.Insert(0, $dialog.SelectedPath)
        }
        $InspectorFolderCombo.Text = $dialog.SelectedPath
    }
})

$InspectFolderButton.Add_Click({
    Start-HelperJob -TaskName 'Inspect Folder' -Request (Build-WorkerRequest -Kind 'folder-inspector' -Extra @{
        FolderPath = $InspectorFolderCombo.Text.Trim()
    })
})

$OpenInspectorFolderButton.Add_Click({
    if (Test-Path -LiteralPath $InspectorFolderCombo.Text -PathType Container) {
        Start-Process explorer.exe $InspectorFolderCombo.Text | Out-Null
    }
})

$CancelJobButton.Add_Click({ Cancel-ActiveJob })

$RecentJobsGrid.Add_SelectionChanged({
    if ($null -eq $script:activeJob -and $RecentJobsGrid.SelectedItem) {
        Load-ActivityFromJob -JobRecord $RecentJobsGrid.SelectedItem
    }
})

$window.Add_Closing({
    if ($null -ne $script:activeJob) {
        Cancel-ActiveJob
    }

    $guiState = Get-HelperGuiState
    if ($null -eq $guiState) { $guiState = [ordered]@{} }
    $guiState['windowWidth'] = [int]$window.Width
    $guiState['windowHeight'] = [int]$window.Height
    $guiState['windowLeft'] = [int]$window.Left
    $guiState['windowTop'] = [int]$window.Top
    $guiState['selectedTabIndex'] = [int]$TaskTabs.SelectedIndex
    Save-HelperGuiState -GuiState $guiState
    Write-HelperLog -Level 'INFO' -Message 'Helper GUI window closing.'
})

Set-StatusMessage -Message 'Ready. Configure the session context and start a task.' -Color 'Black'
Write-HelperLog -Level 'INFO' -Message 'Helper GUI window opening.'
$null = $window.ShowDialog()
Write-HelperLog -Level 'INFO' -Message 'Helper GUI window closed.'
