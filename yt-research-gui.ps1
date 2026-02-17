#Requires -Version 5.0
param(
    [string]$YtDlpPath = (Join-Path $PSScriptRoot 'yt-dlp.exe')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

$script:logDir = Join-Path $PSScriptRoot 'logs'
$script:logPath = $null
$script:logSync = New-Object object

function Initialize-Logger {
    try {
        if (-not (Test-Path -LiteralPath $script:logDir -PathType Container)) {
            New-Item -ItemType Directory -Path $script:logDir -Force | Out-Null
        }

        $script:logPath = Join-Path $script:logDir 'yt-research-gui.log'
        if (-not (Test-Path -LiteralPath $script:logPath -PathType Leaf)) {
            $null = New-Item -ItemType File -Path $script:logPath -Force
        }

        Write-Log -Level 'INFO' -Message '----------------------------------------'
        Write-Log -Level 'INFO' -Message "New session started. Log path: $script:logPath"
    }
    catch {
        # Logging is best-effort. If setup fails we still allow app startup.
        $script:logPath = $null
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    try {
        if ([string]::IsNullOrWhiteSpace($script:logPath)) {
            return
        }

        $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
        [System.Threading.Monitor]::Enter($script:logSync)
        try {
            Add-Content -LiteralPath $script:logPath -Value $line -Encoding UTF8
        }
        finally {
            [System.Threading.Monitor]::Exit($script:logSync)
        }
    }
    catch {
        # Never let logging failures break the app.
    }
}

function Get-ErrorDetails {
    param([object]$ErrorObject)

    if ($null -eq $ErrorObject) {
        return ''
    }

    if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) {
        $record = [System.Management.Automation.ErrorRecord]$ErrorObject
        $msg = New-Object System.Collections.Generic.List[string]
        $msg.Add("ErrorRecord: $($record.ToString())")
        if ($record.Exception) {
            $msg.Add("ExceptionType: $($record.Exception.GetType().FullName)")
            $msg.Add("ExceptionMessage: $($record.Exception.Message)")
            if ($record.Exception.StackTrace) {
                $msg.Add("ExceptionStackTrace:")
                $msg.Add($record.Exception.StackTrace)
            }
            if ($record.Exception.InnerException) {
                $msg.Add("InnerExceptionType: $($record.Exception.InnerException.GetType().FullName)")
                $msg.Add("InnerExceptionMessage: $($record.Exception.InnerException.Message)")
            }
        }
        if ($record.ScriptStackTrace) {
            $msg.Add("ScriptStackTrace:")
            $msg.Add($record.ScriptStackTrace)
        }
        return ($msg -join [Environment]::NewLine)
    }

    if ($ErrorObject -is [System.Exception]) {
        $ex = [System.Exception]$ErrorObject
        $msg = New-Object System.Collections.Generic.List[string]
        $msg.Add("ExceptionType: $($ex.GetType().FullName)")
        $msg.Add("ExceptionMessage: $($ex.Message)")
        if ($ex.StackTrace) {
            $msg.Add("ExceptionStackTrace:")
            $msg.Add($ex.StackTrace)
        }
        if ($ex.InnerException) {
            $msg.Add("InnerExceptionType: $($ex.InnerException.GetType().FullName)")
            $msg.Add("InnerExceptionMessage: $($ex.InnerException.Message)")
        }
        return ($msg -join [Environment]::NewLine)
    }

    return [string]$ErrorObject
}

Initialize-Logger
Write-Log -Level 'INFO' -Message ("PowerShell version: {0}" -f $PSVersionTable.PSVersion.ToString())
Write-Log -Level 'INFO' -Message "Script startup complete."

function Get-DefaultFirefoxProfilePath {
    $profilesRoot = Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles'
    if (-not (Test-Path -LiteralPath $profilesRoot -PathType Container)) {
        return $null
    }

    $profile = Get-ChildItem -LiteralPath $profilesRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -eq $profile) {
        return $null
    }

    return $profile.FullName
}

function Get-CookieArgs {
    param(
        [bool]$UseCookies,
        [string]$ProfilePath
    )

    if (-not $UseCookies) {
        Write-Log -Level 'DEBUG' -Message 'Cookie auth disabled for this request.'
        return @()
    }

    if ([string]::IsNullOrWhiteSpace($ProfilePath)) {
        throw 'Use cookies is enabled, but Firefox profile path is empty.'
    }

    if (-not (Test-Path -LiteralPath $ProfilePath -PathType Container)) {
        throw "Firefox profile path not found: $ProfilePath"
    }

    $profileName = Split-Path -Leaf $ProfilePath
    Write-Log -Level 'DEBUG' -Message "Cookie auth enabled with Firefox profile folder: $profileName"
    return @('--cookies-from-browser', "firefox:$profileName")
}

function Invoke-YtDlpCapture {
    param(
        [string]$ExePath,
        [string[]]$Arguments
    )

    if (-not (Test-Path -LiteralPath $ExePath -PathType Leaf)) {
        Write-Log -Level 'ERROR' -Message "yt-dlp executable not found at '$ExePath'."
        throw "yt-dlp executable not found: $ExePath"
    }

    Write-Log -Level 'INFO' -Message ("Running yt-dlp command: {0} {1}" -f $ExePath, ($Arguments -join ' '))
    $rawOutput = & $ExePath @Arguments 2>&1
    $exitCode = $LASTEXITCODE

    $lines = foreach ($item in $rawOutput) {
        if ($item -is [System.Management.Automation.ErrorRecord]) {
            $item.ToString()
        } else {
            [string]$item
        }
    }

    $outputText = ($lines -join [Environment]::NewLine).Trim()
    Write-Log -Level 'INFO' -Message "yt-dlp exit code: $exitCode"
    if (-not [string]::IsNullOrWhiteSpace($outputText)) {
        Write-Log -Level 'DEBUG' -Message ("yt-dlp output:`n{0}" -f $outputText)
    } else {
        Write-Log -Level 'DEBUG' -Message 'yt-dlp produced no output.'
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output   = $outputText
    }
}

function Convert-SecondsToClock {
    param([Nullable[double]]$Seconds)

    if ($null -eq $Seconds) {
        return ''
    }

    try {
        $span = [TimeSpan]::FromSeconds([double]$Seconds)
        if ($span.TotalHours -ge 1) {
            return ('{0:00}:{1:00}:{2:00}' -f [int]$span.TotalHours, $span.Minutes, $span.Seconds)
        }

        return ('{0:00}:{1:00}' -f $span.Minutes, $span.Seconds)
    }
    catch {
        return ''
    }
}

function Convert-UploadDate {
    param([string]$UploadDate)

    if ([string]::IsNullOrWhiteSpace($UploadDate)) {
        return ''
    }

    try {
        if ($UploadDate -match '^\d{8}$') {
            return [datetime]::ParseExact($UploadDate, 'yyyyMMdd', $null).ToString('yyyy-MM-dd')
        }

        return $UploadDate
    }
    catch {
        return $UploadDate
    }
}

function Format-Number {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    try {
        return ('{0:n0}' -f [double]$Value)
    }
    catch {
        return [string]$Value
    }
}

function Convert-SubtitleFileToText {
    param([string]$Path)

    $lines = Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
    $result = New-Object System.Collections.Generic.List[string]
    $lastWritten = $null

    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if (
            $trimmed -match '^(WEBVTT|NOTE|STYLE|REGION|Kind:|Language:|X-TIMESTAMP-MAP)' -or
            $trimmed -match '^\d+$' -or
            $trimmed -match '-->'
        ) {
            continue
        }

        $clean = $trimmed -replace '<[^>]+>', ''
        $clean = $clean -replace '\{\\an\d\}', ''
        $clean = [System.Web.HttpUtility]::HtmlDecode($clean).Trim()

        if ([string]::IsNullOrWhiteSpace($clean)) {
            continue
        }

        if ($lastWritten -ne $clean) {
            $result.Add($clean)
            $lastWritten = $clean
        }
    }

    return ($result -join [Environment]::NewLine)
}

function Get-BestSubtitleFile {
    param([System.IO.FileInfo[]]$Files)

    if ($null -eq $Files -or $Files.Count -eq 0) {
        return $null
    }

    $scored = $Files | ForEach-Object {
        $name = $_.Name.ToLowerInvariant()
        $score = 0

        if ($name -match '\.en(\.|$)') { $score += 100 }
        if ($name -match '\.en-') { $score += 90 }
        if ($name -match '\.orig\.') { $score += 15 }
        if ($name -match '\.srt$') { $score += 10 }
        if ($name -match '\.vtt$') { $score += 5 }
        if ($name -match 'live_chat') { $score -= 200 }
        if ($name -match 'description') { $score -= 50 }

        [pscustomobject]@{
            File  = $_
            Score = $score
        }
    }

    return ($scored | Sort-Object -Property Score -Descending | Select-Object -First 1).File
}

function Get-TranscriptData {
    param(
        [string]$ExePath,
        [string]$Url,
        [string[]]$CookieArgs
    )

    $tempDir = Join-Path $env:TEMP ("yt-research-subs-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    Write-Log -Level 'INFO' -Message "Transcript fetch temp directory: $tempDir"

    try {
        $outputTemplate = Join-Path $tempDir '%(id)s.%(ext)s'
        $args = @(
            '--no-update',
            '--skip-download',
            '--no-playlist',
            '--no-warnings',
            '--write-subs',
            '--write-auto-subs',
            '--sub-langs', 'en.*,en',
            '--sub-format', 'srt/vtt/best',
            '--output', $outputTemplate,
            $Url
        )

        if ($CookieArgs.Count -gt 0) {
            $args = $CookieArgs + $args
        }

        $fetchResult = Invoke-YtDlpCapture -ExePath $ExePath -Arguments $args

        $subtitleFiles = @(
            Get-ChildItem -LiteralPath $tempDir -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -in '.srt', '.vtt' }
        )
        $subtitleCount = $subtitleFiles.Count
        Write-Log -Level 'INFO' -Message "Subtitle candidates found: $subtitleCount"

        if ($subtitleFiles.Count -eq 0) {
            $status = if ($fetchResult.ExitCode -eq 0) {
                'No subtitles/transcript found for this video.'
            } else {
                "Subtitle fetch failed (exit code $($fetchResult.ExitCode))."
            }

            return [pscustomobject]@{
                Text   = ''
                Status = $status
                Source = ''
                Output = $fetchResult.Output
            }
        }

        $bestFile = Get-BestSubtitleFile -Files $subtitleFiles
        Write-Log -Level 'INFO' -Message "Selected subtitle file: $($bestFile.FullName)"
        $text = Convert-SubtitleFileToText -Path $bestFile.FullName
        Write-Log -Level 'INFO' -Message ("Transcript character count: {0}" -f $text.Length)

        $statusMessage = if ([string]::IsNullOrWhiteSpace($text)) {
            "Subtitle file found but no readable lines extracted ($($bestFile.Name))."
        } else {
            "Transcript extracted from $($bestFile.Name)."
        }

        return [pscustomobject]@{
            Text   = $text
            Status = $statusMessage
            Source = $bestFile.Name
            Output = $fetchResult.Output
        }
    }
    finally {
        Write-Log -Level 'DEBUG' -Message "Cleaning transcript temp directory: $tempDir"
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Build-MetaRows {
    param($Info)

    $channel = if ($Info.channel) { $Info.channel } elseif ($Info.uploader) { $Info.uploader } else { '' }
    $tagsValue = ''
    if ($Info.tags) {
        $tagsValue = (($Info.tags | Select-Object -First 25) -join ', ')
    }

    $rows = @(
        [pscustomobject]@{ Field = 'Title'; Value = [string]$Info.title },
        [pscustomobject]@{ Field = 'Channel'; Value = [string]$channel },
        [pscustomobject]@{ Field = 'Upload Date'; Value = (Convert-UploadDate -UploadDate ([string]$Info.upload_date)) },
        [pscustomobject]@{ Field = 'Duration'; Value = (Convert-SecondsToClock -Seconds $Info.duration) },
        [pscustomobject]@{ Field = 'Views'; Value = (Format-Number -Value $Info.view_count) },
        [pscustomobject]@{ Field = 'Likes'; Value = (Format-Number -Value $Info.like_count) },
        [pscustomobject]@{ Field = 'Uploader ID'; Value = [string]$Info.uploader_id },
        [pscustomobject]@{ Field = 'Channel ID'; Value = [string]$Info.channel_id },
        [pscustomobject]@{ Field = 'Video ID'; Value = [string]$Info.id },
        [pscustomobject]@{ Field = 'URL'; Value = [string]$Info.webpage_url },
        [pscustomobject]@{ Field = 'Tags'; Value = $tagsValue }
    )

    return $rows
}

function Build-ChapterRows {
    param($Info)

    if ($null -eq $Info.chapters) {
        return @()
    }

    $rows = foreach ($chapter in $Info.chapters) {
        [pscustomobject]@{
            Start = Convert-SecondsToClock -Seconds $chapter.start_time
            End   = Convert-SecondsToClock -Seconds $chapter.end_time
            Title = [string]$chapter.title
        }
    }

    return $rows
}

function Get-VideoResearchData {
    param(
        [string]$ExePath,
        [string]$Url,
        [bool]$UseCookies,
        [string]$ProfilePath
    )

    $cookieArgs = Get-CookieArgs -UseCookies $UseCookies -ProfilePath $ProfilePath
    Write-Log -Level 'INFO' -Message ("Get-VideoResearchData started. URL: {0} | UseCookies: {1}" -f $Url, $UseCookies)

    $metadataArgs = @(
        '--no-update',
        '--quiet',
        '--no-warnings',
        '--dump-single-json',
        '--no-download',
        '--no-playlist',
        $Url
    )

    if ($cookieArgs.Count -gt 0) {
        $metadataArgs = $cookieArgs + $metadataArgs
    }

    $metaResult = Invoke-YtDlpCapture -ExePath $ExePath -Arguments $metadataArgs
    if ($metaResult.ExitCode -ne 0) {
        throw "Metadata fetch failed (exit code $($metaResult.ExitCode)).`n$($metaResult.Output)"
    }

    if ([string]::IsNullOrWhiteSpace($metaResult.Output)) {
        throw 'Metadata fetch returned empty output.'
    }

    $info = $metaResult.Output | ConvertFrom-Json -ErrorAction Stop
    Write-Log -Level 'INFO' -Message ("Metadata parsed successfully. Video ID: {0} | Title: {1}" -f [string]$info.id, [string]$info.title)
    $metaRows = Build-MetaRows -Info $info
    $chapterRows = Build-ChapterRows -Info $info
    $transcript = Get-TranscriptData -ExePath $ExePath -Url $Url -CookieArgs $cookieArgs

    $rawJsonPretty = try {
        $info | ConvertTo-Json -Depth 30
    }
    catch {
        $metaResult.Output
    }

    return [pscustomobject]@{
        Title            = [string]$info.title
        Description      = [string]$info.description
        MetaRows         = $metaRows
        Chapters         = $chapterRows
        Transcript       = [string]$transcript.Text
        TranscriptStatus = [string]$transcript.Status
        TranscriptSource = [string]$transcript.Source
        RawJson          = [string]$rawJsonPretty
    }
}

function Assert-ValidYoutubeUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw 'Please paste a YouTube URL.'
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) {
        throw 'URL is not valid.'
    }

    if ($uri.Scheme -notin @('http', 'https')) {
        throw 'URL must start with http or https.'
    }

    if ($uri.Host -notmatch '(youtube\.com|youtu\.be)$') {
        throw 'URL must be a youtube.com or youtu.be link.'
    }
}

if (-not (Test-Path -LiteralPath $YtDlpPath -PathType Leaf)) {
    Write-Log -Level 'ERROR' -Message "Configured yt-dlp path does not exist: $YtDlpPath"
    throw "Could not find yt-dlp executable at '$YtDlpPath'."
}
Write-Log -Level 'INFO' -Message "Using yt-dlp path: $YtDlpPath"

$defaultProfilePath = Get-DefaultFirefoxProfilePath
Write-Log -Level 'INFO' -Message ("Default Firefox profile detection result: {0}" -f $(if ($defaultProfilePath) { $defaultProfilePath } else { '<none>' }))

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="YT Research GUI" Height="840" Width="1240" MinHeight="700" MinWidth="980"
        WindowStartupLocation="CenterScreen">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="YouTube URL:" VerticalAlignment="Center" Margin="0,0,8,0" FontWeight="SemiBold"/>
            <TextBox x:Name="UrlBox" Width="840" Height="30" Margin="0,0,8,0" VerticalContentAlignment="Center"/>
            <Button x:Name="FetchButton" Width="120" Height="30" Content="Fetch"/>
        </StackPanel>

        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <CheckBox x:Name="UseCookiesCheck" VerticalAlignment="Center" Content="Use Firefox cookies" Margin="0,0,16,0"/>
            <TextBlock Text="Profile folder:" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox x:Name="ProfilePathBox" Width="760" Height="28" Margin="0,0,8,0" VerticalContentAlignment="Center"/>
            <Button x:Name="BrowseProfileButton" Width="110" Height="28" Content="Browse..."/>
        </StackPanel>

        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
            <Button x:Name="CopyTitleButton" Content="Copy Title" Width="120" Height="28" Margin="0,0,8,0"/>
            <Button x:Name="CopyDescriptionButton" Content="Copy Description" Width="140" Height="28" Margin="0,0,8,0"/>
            <Button x:Name="CopyTranscriptButton" Content="Copy Transcript" Width="140" Height="28" Margin="0,0,8,0"/>
            <Button x:Name="CopyJsonButton" Content="Copy Raw JSON" Width="130" Height="28" Margin="0,0,8,0"/>
        </StackPanel>

        <TabControl Grid.Row="3" x:Name="ResultsTabs">
            <TabItem Header="Overview">
                <Grid Margin="8">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="2*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Core Metadata" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <DataGrid Grid.Row="1" x:Name="MetaGrid" IsReadOnly="True" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column" Margin="0,0,0,10">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Field" Binding="{Binding Field}" Width="220"/>
                            <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <TextBlock Grid.Row="2" Text="Chapters" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <DataGrid Grid.Row="3" x:Name="ChaptersGrid" IsReadOnly="True" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Start" Binding="{Binding Start}" Width="110"/>
                            <DataGridTextColumn Header="End" Binding="{Binding End}" Width="110"/>
                            <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>

            <TabItem Header="Description">
                <Grid Margin="8">
                    <TextBox x:Name="DescriptionTextBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                             FontFamily="Consolas" FontSize="13"/>
                </Grid>
            </TabItem>

            <TabItem Header="Transcript">
                <Grid Margin="8">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" x:Name="TranscriptStatusText" Margin="0,0,0,6"/>
                    <TextBox Grid.Row="1" x:Name="TranscriptTextBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                             FontFamily="Consolas" FontSize="13"/>
                </Grid>
            </TabItem>

            <TabItem Header="Raw JSON">
                <Grid Margin="8">
                    <TextBox x:Name="RawJsonTextBox" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                             FontFamily="Consolas" FontSize="12"/>
                </Grid>
            </TabItem>
        </TabControl>

        <TextBlock Grid.Row="4" x:Name="StatusTextBlock" Margin="0,8,0,0" Text="Ready. Paste a YouTube URL and click Fetch."/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$UrlBox = $window.FindName('UrlBox')
$FetchButton = $window.FindName('FetchButton')
$UseCookiesCheck = $window.FindName('UseCookiesCheck')
$ProfilePathBox = $window.FindName('ProfilePathBox')
$BrowseProfileButton = $window.FindName('BrowseProfileButton')
$CopyTitleButton = $window.FindName('CopyTitleButton')
$CopyDescriptionButton = $window.FindName('CopyDescriptionButton')
$CopyTranscriptButton = $window.FindName('CopyTranscriptButton')
$CopyJsonButton = $window.FindName('CopyJsonButton')
$MetaGrid = $window.FindName('MetaGrid')
$ChaptersGrid = $window.FindName('ChaptersGrid')
$DescriptionTextBox = $window.FindName('DescriptionTextBox')
$TranscriptStatusText = $window.FindName('TranscriptStatusText')
$TranscriptTextBox = $window.FindName('TranscriptTextBox')
$RawJsonTextBox = $window.FindName('RawJsonTextBox')
$StatusTextBlock = $window.FindName('StatusTextBlock')

$script:lastResult = $null
$ProfilePathBox.Text = if ($defaultProfilePath) { $defaultProfilePath } else { '' }
$UseCookiesCheck.IsChecked = [bool]$defaultProfilePath

function Set-Status {
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

function Set-ClipboardText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    [System.Windows.Clipboard]::SetText($Text)
}

$script:isFetching = $false

$FetchButton.Add_Click({
    if ($script:isFetching) {
        return
    }

    try {
        $url = $UrlBox.Text.Trim()
        Assert-ValidYoutubeUrl -Url $url

        $useCookies = [bool]$UseCookiesCheck.IsChecked
        $profilePath = $ProfilePathBox.Text.Trim()
        Write-Log -Level 'INFO' -Message ("Fetch clicked. URL: {0} | UseCookies: {1} | ProfilePath: {2}" -f $url, $useCookies, $(if ($profilePath) { $profilePath } else { '<empty>' }))

        $script:isFetching = $true
        $FetchButton.IsEnabled = $false
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Set-Status -Message 'Fetching metadata and transcript...' -Color 'DarkBlue'
        Write-Log -Level 'INFO' -Message "Fetch execution started on UI thread."

        $result = Get-VideoResearchData -ExePath $YtDlpPath -Url $url -UseCookies $useCookies -ProfilePath $profilePath
        $script:lastResult = $result

        $MetaGrid.ItemsSource = $script:lastResult.MetaRows
        $ChaptersGrid.ItemsSource = $script:lastResult.Chapters
        $DescriptionTextBox.Text = $script:lastResult.Description
        $TranscriptTextBox.Text = $script:lastResult.Transcript
        $RawJsonTextBox.Text = $script:lastResult.RawJson

        if ([string]::IsNullOrWhiteSpace($script:lastResult.TranscriptSource)) {
            $TranscriptStatusText.Text = $script:lastResult.TranscriptStatus
        } else {
            $TranscriptStatusText.Text = "$($script:lastResult.TranscriptStatus) Source: $($script:lastResult.TranscriptSource)"
        }

        $titleForStatus = if ([string]::IsNullOrWhiteSpace($script:lastResult.Title)) { 'video' } else { $script:lastResult.Title }
        Set-Status -Message "Loaded metadata for: $titleForStatus" -Color 'Green'
        Write-Log -Level 'INFO' -Message ("Fetch execution completed. UI populated for title: {0}" -f $titleForStatus)
    }
    catch {
        Write-Log -Level 'ERROR' -Message ("Fetch failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($_.Exception.Message + $logHint) -Color 'Red'
        [System.Windows.MessageBox]::Show(($_.Exception.Message + $logHint), 'Fetch Failed', 'OK', 'Error') | Out-Null
    }
    finally {
        $script:isFetching = $false
        $FetchButton.IsEnabled = $true
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        Write-Log -Level 'DEBUG' -Message 'Fetch attempt finished.'
    }
})

$BrowseProfileButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select Firefox profile folder'
    $dialog.ShowNewFolderButton = $false

    if (-not [string]::IsNullOrWhiteSpace($ProfilePathBox.Text) -and (Test-Path -LiteralPath $ProfilePathBox.Text -PathType Container)) {
        $dialog.SelectedPath = $ProfilePathBox.Text
    }

    $result = $dialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $ProfilePathBox.Text = $dialog.SelectedPath
        $UseCookiesCheck.IsChecked = $true
        Write-Log -Level 'INFO' -Message ("Firefox profile selected via UI: {0}" -f $dialog.SelectedPath)
    }
})

$CopyTitleButton.Add_Click({
    if ($script:lastResult -and -not [string]::IsNullOrWhiteSpace($script:lastResult.Title)) {
        Set-ClipboardText -Text $script:lastResult.Title
        Set-Status -Message 'Title copied to clipboard.' -Color 'DarkGreen'
    }
})

$CopyDescriptionButton.Add_Click({
    if ($script:lastResult -and -not [string]::IsNullOrWhiteSpace($script:lastResult.Description)) {
        Set-ClipboardText -Text $script:lastResult.Description
        Set-Status -Message 'Description copied to clipboard.' -Color 'DarkGreen'
    }
})

$CopyTranscriptButton.Add_Click({
    if ($script:lastResult -and -not [string]::IsNullOrWhiteSpace($script:lastResult.Transcript)) {
        Set-ClipboardText -Text $script:lastResult.Transcript
        Set-Status -Message 'Transcript copied to clipboard.' -Color 'DarkGreen'
    }
})

$CopyJsonButton.Add_Click({
    if ($script:lastResult -and -not [string]::IsNullOrWhiteSpace($script:lastResult.RawJson)) {
        Set-ClipboardText -Text $script:lastResult.RawJson
        Set-Status -Message 'Raw JSON copied to clipboard.' -Color 'DarkGreen'
    }
})

Write-Log -Level 'INFO' -Message 'GUI window opening.'
$null = $window.ShowDialog()
Write-Log -Level 'INFO' -Message 'GUI window closed.'
