Set-StrictMode -Version Latest

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
}
catch {
}

$script:helperCoreRoot = $PSScriptRoot
$script:helperRepoRoot = Split-Path -Parent $script:helperCoreRoot
$script:helperPrefsPath = Join-Path $script:helperCoreRoot 'user_preferences.json'
$script:helperLegacyPrefsPath = Join-Path $script:helperRepoRoot 'user_preferences.json'
$script:helperCacheDir = Join-Path $script:helperCoreRoot 'cache'
$script:helperLegacyCacheDir = Join-Path $script:helperRepoRoot 'cache'
$script:helperLogsDir = Join-Path $script:helperCoreRoot 'logs'
$script:helperCurrentLogLevel = 'INFO'
$script:helperLogLevels = @{
    DEBUG = 10
    INFO  = 20
    WARN  = 30
    ERROR = 40
}
$script:helperLogFilePath = $null
$script:helperLogHandler = $null
$script:helperRuntimeMode = 'cli'
$script:helperTreatWarningsAsNonFatal = $false
$script:helperState = [ordered]@{
    lastChoice                          = $null
    lastPlaylistIndex                   = $null
    lastPlaylistId                      = $null
    currentYtDlpVersion                 = $null
    lastDownloadRootPath                = $null
    firefoxProfilePath                  = $null
    treatYtDlpErrorsAsWarningsPreferred = $false
    gui                                 = [ordered]@{}
}

function ConvertTo-HelperHashtable {
    param([object]$InputObject)

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $result[[string]$key] = ConvertTo-HelperHashtable -InputObject $InputObject[$key]
        }
        return $result
    }

    $properties = @()
    if ($InputObject.PSObject) {
        $properties = @($InputObject.PSObject.Properties)
    }

    if (($InputObject -is [pscustomobject] -or $properties.Count -gt 0) -and -not ($InputObject -is [string])) {
        $result = [ordered]@{}
        foreach ($prop in $properties) {
            if ($prop.MemberType -notin @('NoteProperty', 'Property', 'AliasProperty', 'ScriptProperty')) {
                continue
            }
            $result[$prop.Name] = ConvertTo-HelperHashtable -InputObject $prop.Value
        }
        return $result
    }

    if (($InputObject -is [System.Collections.IEnumerable]) -and -not ($InputObject -is [string])) {
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($item in $InputObject) {
            $items.Add((ConvertTo-HelperHashtable -InputObject $item))
        }
        return @($items)
    }

    return $InputObject
}

function Get-HelperDefaultDownloadRoot {
    if (-not [string]::IsNullOrWhiteSpace([string]$script:helperState.lastDownloadRootPath)) {
        return [string]$script:helperState.lastDownloadRootPath
    }

    return (Join-Path $script:helperCoreRoot 'Downloads')
}

function Read-TextFileSafe {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return ''
    }

    try {
        return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop)
    }
    catch {
        return ''
    }
}

function ConvertTo-ProcessArgument {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    if ($Value.Length -eq 0) {
        return '""'
    }

    if ($Value -notmatch '[\s"]') {
        return $Value
    }

    $escaped = $Value -replace '(\\*)"', '$1$1\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    return '"' + $escaped + '"'
}

function Join-ProcessArguments {
    param([string[]]$Arguments)

    if ($null -eq $Arguments -or $Arguments.Count -eq 0) {
        return ''
    }

    return (($Arguments | ForEach-Object { ConvertTo-ProcessArgument -Value ([string]$_) }) -join ' ')
}

function New-HelperFieldRow {
    param(
        [string]$Field,
        [string]$Value
    )

    return [pscustomobject]@{
        Field = $Field
        Value = if ($null -eq $Value) { '' } else { [string]$Value }
    }
}

function Should-WriteHelperLog {
    param([string]$Level)

    if (-not $script:helperLogLevels.ContainsKey($script:helperCurrentLogLevel)) {
        $script:helperCurrentLogLevel = 'INFO'
    }

    if (-not $script:helperLogLevels.ContainsKey($Level)) {
        return $true
    }

    return ($script:helperLogLevels[$Level] -ge $script:helperLogLevels[$script:helperCurrentLogLevel])
}

function Set-HelperLogHandler {
    param([scriptblock]$Handler)

    $script:helperLogHandler = $Handler
}

function Set-HelperLogFilePath {
    param([string]$Path)

    $script:helperLogFilePath = $null
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    $logDir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -LiteralPath $logDir -PathType Container)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $null = New-Item -ItemType File -Path $Path -Force
    }

    $script:helperLogFilePath = $Path
}

function Load-HelperLoggingSettings {
    param([string]$ConfigPath)

    $script:helperCurrentLogLevel = 'INFO'

    if ([string]::IsNullOrWhiteSpace($ConfigPath) -or -not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        return
    }

    try {
        $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $config -or -not ($config.PSObject.Properties.Name -contains 'logging')) {
            return
        }

        $logging = $config.logging
        if ($null -eq $logging) {
            return
        }

        if ($logging.PSObject.Properties.Name -contains 'level') {
            $candidateLevel = ([string]$logging.level).Trim().ToUpperInvariant()
            if ($script:helperLogLevels.ContainsKey($candidateLevel)) {
                $script:helperCurrentLogLevel = $candidateLevel
            }
        }
    }
    catch {
    }
}

function Write-HelperLog {
    param(
        [string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if ([string]::IsNullOrWhiteSpace($Message) -or -not (Should-WriteHelperLog -Level $Level)) {
        return
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

    if (-not [string]::IsNullOrWhiteSpace($script:helperLogFilePath)) {
        try {
            Add-Content -LiteralPath $script:helperLogFilePath -Value $line -Encoding UTF8
        }
        catch {
        }
    }

    if ($script:helperLogHandler) {
        & $script:helperLogHandler ([pscustomobject]@{
            Timestamp = $timestamp
            Level     = $Level
            Message   = $Message
            Line      = $line
        })
        return
    }

    switch ($Level) {
        'DEBUG' { Write-Host $Message -ForegroundColor DarkGray }
        'INFO'  { Write-Host $Message }
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Error $Message }
    }
}

function Move-HelperLegacyStateIfNeeded {
    if ((-not (Test-Path -LiteralPath $script:helperPrefsPath -PathType Leaf)) -and (Test-Path -LiteralPath $script:helperLegacyPrefsPath -PathType Leaf)) {
        try {
            Move-Item -LiteralPath $script:helperLegacyPrefsPath -Destination $script:helperPrefsPath -Force -ErrorAction Stop
        }
        catch {
        }
    }

    if ((-not (Test-Path -LiteralPath $script:helperCacheDir -PathType Container)) -and (Test-Path -LiteralPath $script:helperLegacyCacheDir -PathType Container)) {
        try {
            Move-Item -LiteralPath $script:helperLegacyCacheDir -Destination $script:helperCacheDir -Force -ErrorAction Stop
        }
        catch {
        }
    }
}

function Load-HelperPreferences {
    $script:helperState = [ordered]@{
        lastChoice                          = $null
        lastPlaylistIndex                   = $null
        lastPlaylistId                      = $null
        currentYtDlpVersion                 = $null
        lastDownloadRootPath                = $null
        firefoxProfilePath                  = $null
        treatYtDlpErrorsAsWarningsPreferred = $false
        gui                                 = [ordered]@{}
    }

    if (-not (Test-Path -LiteralPath $script:helperPrefsPath -PathType Leaf)) {
        return
    }

    try {
        $prefs = Get-Content -LiteralPath $script:helperPrefsPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $prefs) {
            return
        }

        if ($prefs.PSObject.Properties.Name -contains 'lastMenuChoice') {
            $script:helperState.lastChoice = [string]$prefs.lastMenuChoice
        }
        if ($prefs.PSObject.Properties.Name -contains 'lastPlaylistIndex') {
            $script:helperState.lastPlaylistIndex = [string]$prefs.lastPlaylistIndex
        }
        if ($prefs.PSObject.Properties.Name -contains 'lastPlaylistId') {
            $script:helperState.lastPlaylistId = [string]$prefs.lastPlaylistId
        }
        if ($prefs.PSObject.Properties.Name -contains 'currentYtDlpVersion') {
            $script:helperState.currentYtDlpVersion = [string]$prefs.currentYtDlpVersion
        }
        if ($prefs.PSObject.Properties.Name -contains 'lastDownloadRootPath') {
            $script:helperState.lastDownloadRootPath = [string]$prefs.lastDownloadRootPath
        }
        if ($prefs.PSObject.Properties.Name -contains 'firefoxProfilePath') {
            $script:helperState.firefoxProfilePath = [string]$prefs.firefoxProfilePath
        }
        if ($prefs.PSObject.Properties.Name -contains 'treatYtDlpErrorsAsWarningsPreferred') {
            $script:helperState.treatYtDlpErrorsAsWarningsPreferred = [bool]$prefs.treatYtDlpErrorsAsWarningsPreferred
        }
        if ($prefs.PSObject.Properties.Name -contains 'gui') {
            $script:helperState.gui = (ConvertTo-HelperHashtable -InputObject $prefs.gui)
        }
    }
    catch {
        Write-HelperLog -Level 'WARN' -Message ("Could not load helper preferences from '{0}': {1}" -f $script:helperPrefsPath, $_.Exception.Message)
    }
}

function Save-HelperPreferences {
    param(
        [string]$MenuChoice = $script:helperState.lastChoice,
        [string]$PlaylistIndex = $script:helperState.lastPlaylistIndex,
        [string]$PlaylistId = $script:helperState.lastPlaylistId,
        [string]$YtDlpVersion = $script:helperState.currentYtDlpVersion,
        [string]$DownloadRoot = $script:helperState.lastDownloadRootPath,
        [string]$FirefoxProfilePath = $script:helperState.firefoxProfilePath,
        [Nullable[bool]]$TreatWarningsPreferred = $null,
        [hashtable]$GuiState = $null
    )

    if ([string]::IsNullOrWhiteSpace($DownloadRoot)) {
        $DownloadRoot = Get-HelperDefaultDownloadRoot
    }

    $script:helperState.lastChoice = $MenuChoice
    $script:helperState.lastPlaylistIndex = $PlaylistIndex
    $script:helperState.lastPlaylistId = $PlaylistId
    $script:helperState.currentYtDlpVersion = $YtDlpVersion
    $script:helperState.lastDownloadRootPath = $DownloadRoot
    $script:helperState.firefoxProfilePath = $FirefoxProfilePath
    if ($null -ne $TreatWarningsPreferred) {
        $script:helperState.treatYtDlpErrorsAsWarningsPreferred = [bool]$TreatWarningsPreferred
    }
    if ($null -ne $GuiState) {
        $script:helperState.gui = (ConvertTo-HelperHashtable -InputObject $GuiState)
    }

    $prefs = [ordered]@{}
    if (Test-Path -LiteralPath $script:helperPrefsPath -PathType Leaf) {
        try {
            $existing = Get-Content -LiteralPath $script:helperPrefsPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($null -ne $existing) {
                foreach ($prop in $existing.PSObject.Properties) {
                    $prefs[$prop.Name] = $prop.Value
                }
            }
        }
        catch {
        }
    }

    $prefs['lastMenuChoice'] = $script:helperState.lastChoice
    $prefs['lastPlaylistIndex'] = $script:helperState.lastPlaylistIndex
    $prefs['lastPlaylistId'] = $script:helperState.lastPlaylistId
    $prefs['currentYtDlpVersion'] = $script:helperState.currentYtDlpVersion
    $prefs['lastDownloadRootPath'] = $script:helperState.lastDownloadRootPath
    $prefs['firefoxProfilePath'] = $script:helperState.firefoxProfilePath
    $prefs['treatYtDlpErrorsAsWarningsPreferred'] = $script:helperState.treatYtDlpErrorsAsWarningsPreferred
    $prefs['gui'] = $script:helperState.gui

    try {
        $prefs | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:helperPrefsPath -Encoding UTF8 -Force
    }
    catch {
        Write-HelperLog -Level 'WARN' -Message ("Could not save helper preferences to '{0}': {1}" -f $script:helperPrefsPath, $_.Exception.Message)
    }
}

function Get-HelperGuiState {
    return (ConvertTo-HelperHashtable -InputObject $script:helperState.gui)
}

function Save-HelperGuiState {
    param([hashtable]$GuiState)

    Save-HelperPreferences -GuiState $GuiState
}

function Add-HelperRecentInspectorFolder {
    param([string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        return
    }

    $gui = Get-HelperGuiState
    if ($null -eq $gui) {
        $gui = [ordered]@{}
    }

    $recent = New-Object System.Collections.Generic.List[string]
    if ($gui.Contains('recentInspectorFolders') -and $null -ne $gui.recentInspectorFolders) {
        foreach ($path in @($gui.recentInspectorFolders)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$path) -and $path -ne $FolderPath -and -not $recent.Contains([string]$path)) {
                $recent.Add([string]$path)
            }
        }
    }

    $recent.Insert(0, $FolderPath)
    while ($recent.Count -gt 8) {
        $recent.RemoveAt($recent.Count - 1)
    }

    $gui['recentInspectorFolders'] = @($recent)
    Save-HelperGuiState -GuiState $gui
}

function Initialize-HelperCore {
    param(
        [ValidateSet('cli', 'gui', 'worker')]
        [string]$Mode = 'cli',
        [string]$LoggingConfigPath,
        [string]$LogFilePath,
        [switch]$TreatYtDlpErrorsAsWarnings
    )

    $script:helperRuntimeMode = $Mode
    $script:helperTreatWarningsAsNonFatal = [bool]$TreatYtDlpErrorsAsWarnings

    Move-HelperLegacyStateIfNeeded
    Load-HelperPreferences
    Load-HelperLoggingSettings -ConfigPath $LoggingConfigPath
    if (-not [string]::IsNullOrWhiteSpace($LogFilePath)) {
        Set-HelperLogFilePath -Path $LogFilePath
    }
}

function Test-FirefoxProfilePath {
    param([AllowNull()][string]$ProfilePath)

    $result = [ordered]@{
        IsValid       = $false
        ProfilePath   = $ProfilePath
        ProfileName   = $null
        StatusMessage = $null
    }

    if ([string]::IsNullOrWhiteSpace($ProfilePath)) {
        $result.StatusMessage = 'profile path is empty or not set'
        return [pscustomobject]$result
    }

    try {
        $resolvedPath = (Resolve-Path -LiteralPath $ProfilePath -ErrorAction Stop).Path
    }
    catch {
        $result.StatusMessage = "profile path does not exist: $ProfilePath"
        return [pscustomobject]$result
    }

    if (-not (Test-Path -LiteralPath $resolvedPath -PathType Container)) {
        $result.StatusMessage = "profile path does not exist: $resolvedPath"
        return [pscustomobject]$result
    }

    try {
        $items = @(Get-ChildItem -LiteralPath $resolvedPath -Force -ErrorAction Stop)
        if ($items.Count -le 0) {
            $result.StatusMessage = "profile path exists but is empty: $resolvedPath"
            return [pscustomobject]$result
        }
    }
    catch {
        $result.StatusMessage = "profile path is not accessible: $($_.Exception.Message)"
        return [pscustomobject]$result
    }

    $result.IsValid = $true
    $result.ProfilePath = $resolvedPath
    $result.ProfileName = Split-Path -Leaf $resolvedPath
    return [pscustomobject]$result
}

function Get-FirefoxProfileCandidates {
    $candidatesByPath = @{}
    $firefoxAppDataRoot = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Mozilla\Firefox'
    $profilesIniPath = Join-Path $firefoxAppDataRoot 'profiles.ini'
    $profilesRoot = Join-Path $firefoxAppDataRoot 'Profiles'

    function Add-Candidate {
        param(
            [string]$CandidatePath,
            [string]$CandidateName,
            [string]$Source,
            [bool]$IsDefault = $false
        )

        if ([string]::IsNullOrWhiteSpace($CandidatePath)) {
            return
        }

        $validation = Test-FirefoxProfilePath -ProfilePath $CandidatePath
        if (-not $validation.IsValid) {
            return
        }

        $key = $validation.ProfilePath.ToLowerInvariant()
        if ($candidatesByPath.ContainsKey($key)) {
            $existing = $candidatesByPath[$key]
            if ($IsDefault) {
                $existing.IsDefault = $true
            }
            if ([string]::IsNullOrWhiteSpace([string]$existing.Name) -and -not [string]::IsNullOrWhiteSpace($CandidateName)) {
                $existing.Name = $CandidateName
            }
            if ([string]::IsNullOrWhiteSpace([string]$existing.Source)) {
                $existing.Source = $Source
            }
            return
        }

        $candidatesByPath[$key] = [pscustomobject]@{
            Name        = if ([string]::IsNullOrWhiteSpace($CandidateName)) { $validation.ProfileName } else { $CandidateName }
            ProfileName = $validation.ProfileName
            Path        = $validation.ProfilePath
            Source      = $Source
            IsDefault   = $IsDefault
        }
    }

    function Add-ProfileSection {
        param(
            [string]$SectionName,
            [hashtable]$SectionData
        )

        if ($SectionName -notlike 'Profile*' -or -not $SectionData.ContainsKey('Path')) {
            return
        }

        $candidatePath = [string]$SectionData['Path']
        $isRelative = $true
        if ($SectionData.ContainsKey('IsRelative') -and $SectionData['IsRelative'] -eq '0') {
            $isRelative = $false
        }

        if ($isRelative) {
            $candidatePath = Join-Path $firefoxAppDataRoot $candidatePath
        }

        Add-Candidate -CandidatePath $candidatePath -CandidateName ([string]$SectionData['Name']) -Source 'profiles.ini' -IsDefault:($SectionData.ContainsKey('Default') -and $SectionData['Default'] -eq '1')
    }

    if (Test-Path -LiteralPath $profilesIniPath -PathType Leaf) {
        try {
            $currentSectionName = $null
            $currentSectionData = @{}
            foreach ($line in Get-Content -LiteralPath $profilesIniPath -ErrorAction Stop) {
                $trimmed = $line.Trim()
                if ($trimmed -match '^\[(.+)\]$') {
                    Add-ProfileSection -SectionName $currentSectionName -SectionData $currentSectionData
                    $currentSectionName = $matches[1]
                    $currentSectionData = @{}
                    continue
                }

                if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith(';') -or -not $currentSectionName) {
                    continue
                }

                $parts = $trimmed -split '=', 2
                if ($parts.Count -eq 2) {
                    $currentSectionData[$parts[0].Trim()] = $parts[1].Trim()
                }
            }

            Add-ProfileSection -SectionName $currentSectionName -SectionData $currentSectionData
        }
        catch {
            Write-HelperLog -Level 'WARN' -Message ("Could not inspect Firefox profiles.ini: {0}" -f $_.Exception.Message)
        }
    }

    if (Test-Path -LiteralPath $profilesRoot -PathType Container) {
        try {
            foreach ($profileDir in Get-ChildItem -LiteralPath $profilesRoot -Directory -ErrorAction Stop) {
                Add-Candidate -CandidatePath $profileDir.FullName -CandidateName $profileDir.Name -Source 'Profiles folder'
            }
        }
        catch {
            Write-HelperLog -Level 'WARN' -Message ("Could not inspect Firefox profile folders: {0}" -f $_.Exception.Message)
        }
    }

    return @(
        $candidatesByPath.Values |
            Sort-Object -Property @{ Expression = { if ($_.IsDefault) { 0 } else { 1 } } }, Name, Path
    )
}

function Get-PreferredFirefoxProfilePath {
    $saved = Test-FirefoxProfilePath -ProfilePath $script:helperState.firefoxProfilePath
    if ($saved.IsValid) {
        return $saved.ProfilePath
    }

    $candidates = @(Get-FirefoxProfileCandidates)
    if ($candidates.Count -gt 0) {
        $defaultCandidate = $candidates | Where-Object { $_.IsDefault } | Select-Object -First 1
        if ($defaultCandidate) {
            return [string]$defaultCandidate.Path
        }

        return [string]$candidates[0].Path
    }

    return $null
}

function Resolve-HelperAuthState {
    param(
        [bool]$UseCookies,
        [string]$FirefoxProfilePath
    )

    $preferredPath = if ([string]::IsNullOrWhiteSpace($FirefoxProfilePath)) {
        if ([string]::IsNullOrWhiteSpace([string]$script:helperState.firefoxProfilePath)) {
            Get-PreferredFirefoxProfilePath
        }
        else {
            $script:helperState.firefoxProfilePath
        }
    }
    else {
        $FirefoxProfilePath
    }

    $validation = Test-FirefoxProfilePath -ProfilePath $preferredPath
    $authReady = $validation.IsValid

    if ($authReady -and $validation.ProfilePath -ne $script:helperState.firefoxProfilePath) {
        Save-HelperPreferences -FirefoxProfilePath $validation.ProfilePath
    }

    if ($UseCookies -and -not $authReady) {
        throw "Firefox cookies were requested, but no usable Firefox profile is available ($($validation.StatusMessage))."
    }

    return [pscustomobject]@{
        UseCookies  = $UseCookies -and $authReady
        AuthReady   = $authReady
        ProfilePath = if ($authReady) { $validation.ProfilePath } else { $preferredPath }
        ProfileName = if ($authReady) { $validation.ProfileName } else { $null }
        AuthValue   = if ($authReady) { "firefox:$($validation.ProfileName)" } else { $null }
        Status      = if ($authReady) { "profile ready: $($validation.ProfilePath)" } else { $validation.StatusMessage }
    }
}

function Get-LocalYtDlpVersion {
    param([string]$ExePath)

    if ([string]::IsNullOrWhiteSpace($ExePath) -or -not (Test-Path -LiteralPath $ExePath -PathType Leaf)) {
        return $null
    }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ExePath
        $psi.Arguments = '--version'
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $process = [System.Diagnostics.Process]::Start($psi)
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -eq 0 -and $stdout -match '^\d{4}\.\d{2}\.\d{2}(?:\.\d{6})?') {
            return $matches[0]
        }

        Write-HelperLog -Level 'WARN' -Message ("Could not get yt-dlp version from '{0}'. ExitCode: {1} | Stdout: {2} | Stderr: {3}" -f $ExePath, $process.ExitCode, $stdout.Trim(), $stderr.Trim())
        return $null
    }
    catch {
        Write-HelperLog -Level 'WARN' -Message ("Error running yt-dlp --version: {0}" -f $_.Exception.Message)
        return $null
    }
}

function Get-AbsoluteFfmpegBinPath {
    param([string]$BaseInstallPath)

    if ([string]::IsNullOrWhiteSpace($BaseInstallPath) -or -not (Test-Path -LiteralPath $BaseInstallPath -PathType Container)) {
        return $null
    }

    $extractedFolder = Get-ChildItem -LiteralPath $BaseInstallPath -Directory -ErrorAction SilentlyContinue |
        Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'bin') -PathType Container } |
        Select-Object -First 1

    if ($extractedFolder) {
        return (Join-Path $extractedFolder.FullName 'bin')
    }

    $directBin = Join-Path $BaseInstallPath 'bin'
    if (Test-Path -LiteralPath $directBin -PathType Container) {
        return $directBin
    }

    return $null
}

function Get-YtDlpJsRuntimeArgument {
    $runtimeCandidates = @(
        @{ Command = 'deno'; Runtime = 'deno'; Label = 'Deno' },
        @{ Command = 'node'; Runtime = 'node'; Label = 'Node.js' },
        @{ Command = 'bun'; Runtime = 'bun'; Label = 'Bun' },
        @{ Command = 'qjs'; Runtime = 'quickjs'; Label = 'QuickJS' }
    )

    foreach ($candidate in $runtimeCandidates) {
        $commandInfo = Get-Command $candidate.Command -ErrorAction SilentlyContinue
        if ($commandInfo) {
            return [pscustomobject]@{
                Runtime = $candidate.Runtime
                Label   = $candidate.Label
                Command = $commandInfo.Source
            }
        }
    }

    return $null
}

function Ensure-YtDlpNightly {
    param([string]$ExpectedExePath)

    $resolvedExePath = $null
    $localVersion = $null
    $exeExists = Test-Path -LiteralPath $ExpectedExePath -PathType Leaf

    if ($exeExists) {
        $localVersion = Get-LocalYtDlpVersion -ExePath $ExpectedExePath
        if (-not $localVersion -and $script:helperState.currentYtDlpVersion) {
            $localVersion = $script:helperState.currentYtDlpVersion
        }
        $resolvedExePath = $ExpectedExePath
    }

    $nightlyApiUrl = 'https://api.github.com/repos/yt-dlp/yt-dlp-nightly-builds/releases/latest'
    Write-HelperLog -Level 'INFO' -Message ("Checking GitHub for latest yt-dlp nightly release: {0}" -f $nightlyApiUrl)

    $latestRelease = $null
    try {
        $latestRelease = Invoke-RestMethod -Uri $nightlyApiUrl -UseBasicParsing -TimeoutSec 20
    }
    catch {
        Write-HelperLog -Level 'WARN' -Message ("Failed to fetch latest nightly release info: {0}" -f $_.Exception.Message)
        if ($exeExists) {
            return [pscustomobject]@{
                ExePath = $resolvedExePath
                Version = if ($localVersion) { $localVersion } else { $script:helperState.currentYtDlpVersion }
                Updated = $false
            }
        }
        throw 'yt-dlp.exe is missing and the nightly release check failed.'
    }

    $latestVersion = [string]$latestRelease.tag_name
    if ($latestVersion -like 'v*') {
        $latestVersion = $latestVersion.Substring(1)
    }

    $updateNeeded = (-not $exeExists) -or (-not $localVersion) -or ($latestVersion -gt $localVersion)
    if (-not $updateNeeded) {
        if ($localVersion -and $localVersion -ne $script:helperState.currentYtDlpVersion) {
            Save-HelperPreferences -YtDlpVersion $localVersion
        }

        return [pscustomobject]@{
            ExePath = $resolvedExePath
            Version = $localVersion
            Updated = $false
        }
    }

    $asset = $latestRelease.assets | Where-Object { $_.name -eq 'yt-dlp.exe' } | Select-Object -First 1
    if (-not $asset) {
        if ($exeExists) {
            return [pscustomobject]@{
                ExePath = $resolvedExePath
                Version = if ($localVersion) { $localVersion } else { $script:helperState.currentYtDlpVersion }
                Updated = $false
            }
        }

        throw 'Latest nightly release did not contain yt-dlp.exe.'
    }

    Write-HelperLog -Level 'INFO' -Message ("Downloading yt-dlp nightly {0} to '{1}'." -f $latestVersion, $ExpectedExePath)
    try {
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $ExpectedExePath -UseBasicParsing -ErrorAction Stop
    }
    catch {
        if ($exeExists) {
            Write-HelperLog -Level 'WARN' -Message ("Failed to update yt-dlp. Continuing with existing executable. Error: {0}" -f $_.Exception.Message)
            return [pscustomobject]@{
                ExePath = $resolvedExePath
                Version = if ($localVersion) { $localVersion } else { $script:helperState.currentYtDlpVersion }
                Updated = $false
            }
        }

        throw
    }

    $newVersion = Get-LocalYtDlpVersion -ExePath $ExpectedExePath
    if (-not $newVersion) {
        throw "Downloaded yt-dlp to '$ExpectedExePath', but could not read its version."
    }

    Save-HelperPreferences -YtDlpVersion $newVersion
    return [pscustomobject]@{
        ExePath = $ExpectedExePath
        Version = $newVersion
        Updated = $true
    }
}

function Ensure-SharedFfmpeg {
    $ffmpegRoot = Join-Path $script:helperRepoRoot 'ffmpeg_yt-dlp'
    $binPath = Get-AbsoluteFfmpegBinPath -BaseInstallPath $ffmpegRoot
    if ($binPath) {
        return [pscustomobject]@{
            RootPath   = $ffmpegRoot
            BinPath    = $binPath
            Downloaded = $false
        }
    }

    $downloadUrl = 'https://github.com/yt-dlp/FFmpeg-Builds/releases/latest/download/ffmpeg-master-latest-win64-gpl.zip'
    if (-not (Test-Path -LiteralPath $ffmpegRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $ffmpegRoot -Force | Out-Null
    }

    $zipPath = Join-Path $ffmpegRoot 'ffmpeg_download.zip'
    Write-HelperLog -Level 'INFO' -Message ("Downloading FFmpeg build to '{0}'." -f $zipPath)

    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        Expand-Archive -Path $zipPath -DestinationPath $ffmpegRoot -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $zipPath -PathType Leaf) {
            Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
        }
    }

    $binPath = Get-AbsoluteFfmpegBinPath -BaseInstallPath $ffmpegRoot
    if (-not $binPath) {
        throw "FFmpeg was downloaded, but no bin directory was found under '$ffmpegRoot'."
    }

    return [pscustomobject]@{
        RootPath   = $ffmpegRoot
        BinPath    = $binPath
        Downloaded = $true
    }
}

function Get-HelperLocalEnvironmentSnapshot {
    $ytDlpPath = Join-Path $script:helperRepoRoot 'yt-dlp.exe'
    $ffmpegRoot = Join-Path $script:helperRepoRoot 'ffmpeg_yt-dlp'
    $ffmpegBin = Get-AbsoluteFfmpegBinPath -BaseInstallPath $ffmpegRoot
    $jsRuntime = Get-YtDlpJsRuntimeArgument
    $savedProfileValidation = Test-FirefoxProfilePath -ProfilePath $script:helperState.firefoxProfilePath
    $cachePath = Join-Path $script:helperCacheDir 'playlists_cache.json'

    $cacheAgeText = ''
    if (Test-Path -LiteralPath $cachePath -PathType Leaf) {
        try {
            $age = (Get-Date) - (Get-Item -LiteralPath $cachePath).LastWriteTime
            $cacheAgeText = ("{0}h {1}m" -f [Math]::Floor($age.TotalHours), $age.Minutes)
        }
        catch {
        }
    }

    return [pscustomobject]@{
        DownloadRoot           = Get-HelperDefaultDownloadRoot
        YtDlpPath              = $ytDlpPath
        YtDlpExists            = (Test-Path -LiteralPath $ytDlpPath -PathType Leaf)
        YtDlpVersion           = Get-LocalYtDlpVersion -ExePath $ytDlpPath
        FfmpegRoot             = $ffmpegRoot
        FfmpegBinPath          = $ffmpegBin
        FfmpegReady            = -not [string]::IsNullOrWhiteSpace([string]$ffmpegBin)
        JsRuntime              = if ($jsRuntime) { $jsRuntime.Runtime } else { $null }
        JsRuntimeLabel         = if ($jsRuntime) { $jsRuntime.Label } else { $null }
        FirefoxProfilePath     = if ($savedProfileValidation.IsValid) { $savedProfileValidation.ProfilePath } else { $script:helperState.firefoxProfilePath }
        AuthReady              = $savedProfileValidation.IsValid
        AuthStatus             = if ($savedProfileValidation.IsValid) { 'ready' } else { $savedProfileValidation.StatusMessage }
        PlaylistCachePath      = $cachePath
        PlaylistCacheAge       = $cacheAgeText
        LastChoice             = $script:helperState.lastChoice
        LastPlaylistId         = $script:helperState.lastPlaylistId
        TreatWarningsPreferred = [bool]$script:helperState.treatYtDlpErrorsAsWarningsPreferred
    }
}

function Resolve-HelperDownloadRootPath {
    param([string]$DownloadRoot)

    $candidate = if ([string]::IsNullOrWhiteSpace($DownloadRoot)) { Get-HelperDefaultDownloadRoot } else { $DownloadRoot.Trim() }
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $candidate = Join-Path $script:helperCoreRoot $candidate
    }

    if (-not (Test-Path -LiteralPath $candidate -PathType Container)) {
        New-Item -ItemType Directory -Path $candidate -Force -ErrorAction Stop | Out-Null
    }

    Save-HelperPreferences -DownloadRoot $candidate
    return $candidate
}

function Resolve-HelperEnvironment {
    param(
        [string]$DownloadRoot,
        [bool]$UseCookies,
        [string]$FirefoxProfilePath,
        [switch]$TreatWarningsAsNonFatal
    )

    $resolvedDownloadRoot = Resolve-HelperDownloadRootPath -DownloadRoot $DownloadRoot
    $ytDlp = Ensure-YtDlpNightly -ExpectedExePath (Join-Path $script:helperRepoRoot 'yt-dlp.exe')
    $ffmpeg = Ensure-SharedFfmpeg
    $jsRuntime = Get-YtDlpJsRuntimeArgument
    $authState = Resolve-HelperAuthState -UseCookies:$UseCookies -FirefoxProfilePath $FirefoxProfilePath

    if ($TreatWarningsAsNonFatal.IsPresent) {
        Save-HelperPreferences -TreatWarningsPreferred $true
    }

    return [pscustomobject]@{
        DownloadRoot            = $resolvedDownloadRoot
        YtDlpPath               = $ytDlp.ExePath
        YtDlpVersion            = $ytDlp.Version
        FfmpegRoot              = $ffmpeg.RootPath
        FfmpegBinPath           = $ffmpeg.BinPath
        JsRuntime               = if ($jsRuntime) { $jsRuntime.Runtime } else { $null }
        JsRuntimeLabel          = if ($jsRuntime) { $jsRuntime.Label } else { $null }
        UseCookies              = $authState.UseCookies
        AuthReady               = $authState.AuthReady
        FirefoxProfilePath      = $authState.ProfilePath
        FirefoxProfileName      = $authState.ProfileName
        CookiesFromBrowserValue = $authState.AuthValue
        TreatWarningsAsNonFatal = [bool]$TreatWarningsAsNonFatal
    }
}

function Invoke-ExternalProcess {
    param(
        [string]$FilePath,
        [string[]]$Arguments,
        [string]$DisplayName = 'process',
        [string]$WorkingDirectory,
        [bool]$LogStdOut = $true,
        [bool]$LogStdErr = $true
    )

    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        throw "No executable path was supplied for $DisplayName."
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = Join-ProcessArguments -Arguments $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
        $psi.WorkingDirectory = $WorkingDirectory
    }

    $stdoutBuilder = New-Object System.Text.StringBuilder
    $stderrBuilder = New-Object System.Text.StringBuilder

    Write-HelperLog -Level 'INFO' -Message ("Running {0}: {1} {2}" -f $DisplayName, $FilePath, $psi.Arguments)

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    try {
        $null = $process.Start()

        while (-not $process.HasExited) {
            while ($process.StandardOutput.Peek() -ge 0) {
                $line = $process.StandardOutput.ReadLine()
                [void]$stdoutBuilder.AppendLine($line)
                if ($LogStdOut) {
                    Write-HelperLog -Level 'INFO' -Message ("[{0}] {1}" -f $DisplayName, $line)
                }
            }

            while ($process.StandardError.Peek() -ge 0) {
                $line = $process.StandardError.ReadLine()
                [void]$stderrBuilder.AppendLine($line)
                if ($LogStdErr) {
                    Write-HelperLog -Level 'INFO' -Message ("[{0}:stderr] {1}" -f $DisplayName, $line)
                }
            }

            Start-Sleep -Milliseconds 120
        }

        while ($process.StandardOutput.Peek() -ge 0) {
            $line = $process.StandardOutput.ReadLine()
            [void]$stdoutBuilder.AppendLine($line)
            if ($LogStdOut) {
                Write-HelperLog -Level 'INFO' -Message ("[{0}] {1}" -f $DisplayName, $line)
            }
        }

        while ($process.StandardError.Peek() -ge 0) {
            $line = $process.StandardError.ReadLine()
            [void]$stderrBuilder.AppendLine($line)
            if ($LogStdErr) {
                Write-HelperLog -Level 'INFO' -Message ("[{0}:stderr] {1}" -f $DisplayName, $line)
            }
        }

        return [pscustomobject]@{
            ExitCode       = $process.ExitCode
            StandardOut    = $stdoutBuilder.ToString()
            StandardError  = $stderrBuilder.ToString()
            CombinedOutput = ($stdoutBuilder.ToString().TrimEnd() + [Environment]::NewLine + $stderrBuilder.ToString().TrimEnd()).Trim()
        }
    }
    finally {
        try {
            $process.Dispose()
        }
        catch {
        }
    }
}

function Get-HelperCommonFlags {
    return @(
        '-f', 'bv*+ba/b',
        '--sub-langs', 'en.*,en',
        '--write-subs',
        '--write-auto-subs',
        '--convert-subs', 'srt',
        '--embed-metadata',
        '--embed-subs',
        '--merge-output-format', 'mkv',
        '--no-write-description',
        '--no-write-info-json',
        '--no-write-thumbnail',
        '--progress-delta', '2'
    )
}

function Assert-ValidYoutubeUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw 'Please provide a YouTube URL.'
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) {
        throw 'The provided URL is not valid.'
    }

    if ($uri.Scheme -notin @('http', 'https')) {
        throw 'The URL must start with http or https.'
    }

    if ($uri.Host -notmatch '(youtube\.com|youtu\.be)$') {
        throw 'The URL must be a youtube.com or youtu.be link.'
    }
}

function Get-HelperCookieArguments {
    param([object]$EnvironmentState)

    if ($null -eq $EnvironmentState -or -not [bool]$EnvironmentState.UseCookies) {
        return @()
    }

    return @('--cookies-from-browser', [string]$EnvironmentState.CookiesFromBrowserValue)
}

function Sanitize-HelperFileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ''
    }

    $result = $Name
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $result = $result.Replace($char, '_')
    }

    $result = $result -replace '_+', '_'
    return $result.Trim('_')
}

function Convert-SrtFileToText {
    param([string]$Path)

    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    $stripped = New-Object System.Collections.Generic.List[string]

    foreach ($line in $lines) {
        if ($line -match '^[\t\s]*\d+[\t\s]*$') { continue }
        if ($line -match '\-\-\>') { continue }
        $clean = ($line -replace '<[^>]+>', '')
        $stripped.Add($clean)
    }

    $final = New-Object System.Collections.Generic.List[string]
    $previousBlank = $false
    foreach ($line in $stripped) {
        $isBlank = [string]::IsNullOrWhiteSpace($line)
        if ($isBlank) {
            if (-not $previousBlank) {
                $final.Add('')
                $previousBlank = $true
            }
        }
        else {
            $final.Add($line)
            $previousBlank = $false
        }
    }

    return ($final -join [Environment]::NewLine)
}

function Get-PlaylistNamingInfo {
    param(
        [string]$YtDlpPath,
        [string]$PlaylistUrl,
        [object]$EnvironmentState
    )

    $prefetchArgs = New-Object System.Collections.Generic.List[string]
    if ($EnvironmentState.JsRuntime) {
        $prefetchArgs.Add('--js-runtimes')
        $prefetchArgs.Add([string]$EnvironmentState.JsRuntime)
    }
    foreach ($arg in (Get-HelperCookieArguments -EnvironmentState $EnvironmentState)) {
        $prefetchArgs.Add([string]$arg)
    }
    $prefetchArgs.Add('-J')
    $prefetchArgs.Add('--flat-playlist')
    $prefetchArgs.Add('--playlist-items')
    $prefetchArgs.Add('0')
    $prefetchArgs.Add($PlaylistUrl)

    $response = Invoke-ExternalProcess -FilePath $YtDlpPath -Arguments @($prefetchArgs) -DisplayName 'yt-dlp prefetch' -LogStdOut:$false -LogStdErr:$true
    $playlistTitle = $null
    $playlistId = $null

    if ($response.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($response.StandardOut)) {
        try {
            $info = $response.StandardOut | ConvertFrom-Json -ErrorAction Stop
            if ($null -ne $info) {
                if ($info.PSObject.Properties.Name -contains 'title') {
                    $playlistTitle = [string]$info.title
                }
                if ($info.PSObject.Properties.Name -contains 'id') {
                    $playlistId = [string]$info.id
                }
            }
        }
        catch {
            Write-HelperLog -Level 'WARN' -Message ("Could not parse playlist naming metadata: {0}" -f $_.Exception.Message)
        }
    }

    $folderName = if ($playlistTitle) { Sanitize-HelperFileName -Name $playlistTitle } else { 'Playlist_UnknownTitle' }
    if ($playlistId) {
        $folderName = '{0} [{1}]' -f $folderName, $playlistId
    }

    return [pscustomobject]@{
        Title      = $playlistTitle
        PlaylistId = $playlistId
        FolderName = $folderName
        Prefetch   = $response
    }
}

function Convert-PlaylistSidecarsToText {
    param([string]$SidecarOutputDir)

    $textOutDir = Join-Path $SidecarOutputDir 'text'
    if (-not (Test-Path -LiteralPath $textOutDir -PathType Container)) {
        New-Item -ItemType Directory -Path $textOutDir -Force | Out-Null
    }

    $createdFiles = New-Object System.Collections.Generic.List[string]
    $srtFiles = @(Get-ChildItem -LiteralPath $SidecarOutputDir -Filter '*.srt' -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($srt in $srtFiles) {
        $txtPath = Join-Path $textOutDir ("{0}.txt" -f $srt.BaseName)
        try {
            $content = Convert-SrtFileToText -Path $srt.FullName
            Set-Content -LiteralPath $txtPath -Value $content -Encoding UTF8 -Force
            $createdFiles.Add($txtPath)
        }
        catch {
            Write-HelperLog -Level 'WARN' -Message ("Failed to convert subtitle '{0}' to text: {1}" -f $srt.FullName, $_.Exception.Message)
        }
    }

    return @($createdFiles)
}

function New-HelperOperationResult {
    param(
        [string]$Kind,
        [bool]$Succeeded,
        [string]$Status,
        [string]$Summary,
        [object]$Data,
        [string[]]$Warnings = @()
    )

    return [pscustomobject]@{
        Kind      = $Kind
        Succeeded = $Succeeded
        Status    = $Status
        Summary   = $Summary
        Data      = $Data
        Warnings  = @($Warnings)
    }
}

function Invoke-SingleVideoDownloadOperation {
    param([hashtable]$Request)

    Assert-ValidYoutubeUrl -Url ([string]$Request.Url)
    $environment = Resolve-HelperEnvironment -DownloadRoot ([string]$Request.DownloadRoot) -UseCookies:([bool]$Request.UseCookies) -FirefoxProfilePath ([string]$Request.FirefoxProfilePath) -TreatWarningsAsNonFatal:([bool]$Request.TreatWarningsAsNonFatal)
    $commonArgs = @(Get-HelperCommonFlags)

    $singlesRoot = Join-Path $environment.DownloadRoot 'Singles'
    if (-not (Test-Path -LiteralPath $singlesRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $singlesRoot -Force | Out-Null
    }

    $archivePath = Join-Path $singlesRoot 'download_archive.txt'
    $outputTemplate = Join-Path $singlesRoot '%(title)s [%(id)s].%(ext)s'

    $arguments = New-Object System.Collections.Generic.List[string]
    if ($environment.FfmpegBinPath) {
        $arguments.Add('--ffmpeg-location')
        $arguments.Add([string]$environment.FfmpegBinPath)
    }
    if ($environment.JsRuntime) {
        $arguments.Add('--js-runtimes')
        $arguments.Add([string]$environment.JsRuntime)
    }
    foreach ($arg in $commonArgs) { $arguments.Add([string]$arg) }
    $arguments.Add('--download-archive')
    $arguments.Add($archivePath)
    foreach ($arg in (Get-HelperCookieArguments -EnvironmentState $environment)) { $arguments.Add([string]$arg) }
    $arguments.Add('-o')
    $arguments.Add($outputTemplate)
    $arguments.Add([string]$Request.Url)

    $response = Invoke-ExternalProcess -FilePath $environment.YtDlpPath -Arguments @($arguments) -DisplayName 'yt-dlp single download'
    $succeeded = ($response.ExitCode -eq 0) -or [bool]$Request.TreatWarningsAsNonFatal
    $status = if ($response.ExitCode -eq 0) {
        'Download completed successfully.'
    }
    else {
        "yt-dlp finished with exit code $($response.ExitCode)."
    }

    return (New-HelperOperationResult -Kind 'single-video' -Succeeded:$succeeded -Status $status -Summary ('Single video -> {0}' -f $singlesRoot) -Data ([pscustomobject]@{
        DownloadRoot = $environment.DownloadRoot
        OutputRoot   = $singlesRoot
        ArchivePath  = $archivePath
        ExitCode     = $response.ExitCode
        YtDlpVersion = $environment.YtDlpVersion
    }))
}

function Invoke-PlaylistDownloadOperation {
    param([hashtable]$Request)

    Assert-ValidYoutubeUrl -Url ([string]$Request.Url)
    $environment = Resolve-HelperEnvironment -DownloadRoot ([string]$Request.DownloadRoot) -UseCookies:([bool]$Request.UseCookies) -FirefoxProfilePath ([string]$Request.FirefoxProfilePath) -TreatWarningsAsNonFatal:([bool]$Request.TreatWarningsAsNonFatal)
    $commonArgs = @(Get-HelperCommonFlags)
    $namingInfo = Get-PlaylistNamingInfo -YtDlpPath $environment.YtDlpPath -PlaylistUrl ([string]$Request.Url) -EnvironmentState $environment
    $playlistOutputDir = Join-Path $environment.DownloadRoot $namingInfo.FolderName

    if (-not (Test-Path -LiteralPath $playlistOutputDir -PathType Container)) {
        New-Item -ItemType Directory -Path $playlistOutputDir -Force | Out-Null
    }

    $archivePath = Join-Path $playlistOutputDir 'download_archive.txt'
    $arguments = New-Object System.Collections.Generic.List[string]
    if ($environment.FfmpegBinPath) {
        $arguments.Add('--ffmpeg-location')
        $arguments.Add([string]$environment.FfmpegBinPath)
    }
    if ($environment.JsRuntime) {
        $arguments.Add('--js-runtimes')
        $arguments.Add([string]$environment.JsRuntime)
    }
    foreach ($arg in $commonArgs) { $arguments.Add([string]$arg) }
    $arguments.Add('--download-archive')
    $arguments.Add($archivePath)
    foreach ($arg in (Get-HelperCookieArguments -EnvironmentState $environment)) { $arguments.Add([string]$arg) }
    $arguments.Add('--output')
    $arguments.Add((Join-Path $playlistOutputDir '%(title)s [%(id)s].%(ext)s'))
    $arguments.Add([string]$Request.Url)

    $response = Invoke-ExternalProcess -FilePath $environment.YtDlpPath -Arguments @($arguments) -DisplayName 'yt-dlp playlist download'
    $succeeded = ($response.ExitCode -eq 0) -or [bool]$Request.TreatWarningsAsNonFatal
    $warnings = New-Object System.Collections.Generic.List[string]
    if ($response.ExitCode -ne 0) {
        $warnings.Add("yt-dlp exit code: $($response.ExitCode)")
    }

    $sidecarOutputDir = $null
    $textFiles = @()
    if ([bool]$Request.FetchSidecars) {
        $sidecarOutputDir = Join-Path $playlistOutputDir '_sidecar'
        if (-not (Test-Path -LiteralPath $sidecarOutputDir -PathType Container)) {
            New-Item -ItemType Directory -Path $sidecarOutputDir -Force | Out-Null
        }

        $sidecarArgs = New-Object System.Collections.Generic.List[string]
        if ($environment.FfmpegBinPath) {
            $sidecarArgs.Add('--ffmpeg-location')
            $sidecarArgs.Add([string]$environment.FfmpegBinPath)
        }
        if ($environment.JsRuntime) {
            $sidecarArgs.Add('--js-runtimes')
            $sidecarArgs.Add([string]$environment.JsRuntime)
        }
        foreach ($arg in (Get-HelperCookieArguments -EnvironmentState $environment)) { $sidecarArgs.Add([string]$arg) }
        $sidecarArgs.Add('--skip-download')
        $sidecarArgs.Add('--write-info-json')
        $sidecarArgs.Add('--write-subs')
        $sidecarArgs.Add('--write-auto-subs')
        $sidecarArgs.Add('--convert-subs')
        $sidecarArgs.Add('srt')
        $sidecarArgs.Add('--sub-langs')
        $sidecarArgs.Add('en.*,en')
        $sidecarArgs.Add('--output')
        $sidecarArgs.Add((Join-Path $sidecarOutputDir '%(title)s [%(id)s].%(ext)s'))
        $sidecarArgs.Add([string]$Request.Url)

        $sidecarResponse = Invoke-ExternalProcess -FilePath $environment.YtDlpPath -Arguments @($sidecarArgs) -DisplayName 'yt-dlp playlist sidecars'
        if ($sidecarResponse.ExitCode -ne 0) {
            $warnings.Add("Sidecar fetch exit code: $($sidecarResponse.ExitCode)")
            if (-not [bool]$Request.TreatWarningsAsNonFatal) {
                $succeeded = $false
            }
        }

        $textFiles = @(Convert-PlaylistSidecarsToText -SidecarOutputDir $sidecarOutputDir)
    }

    return (New-HelperOperationResult -Kind 'playlist-url' -Succeeded:$succeeded -Status (if ($succeeded) { 'Playlist processing finished.' } else { 'Playlist processing finished with errors.' }) -Summary ('Playlist -> {0}' -f $playlistOutputDir) -Data ([pscustomobject]@{
        DownloadRoot     = $environment.DownloadRoot
        OutputRoot       = $playlistOutputDir
        ArchivePath      = $archivePath
        PlaylistTitle    = $namingInfo.Title
        PlaylistId       = $namingInfo.PlaylistId
        ExitCode         = $response.ExitCode
        SidecarOutputDir = $sidecarOutputDir
        TextFiles        = @($textFiles)
        YtDlpVersion     = $environment.YtDlpVersion
    }) -Warnings @($warnings))
}

function Get-HelperPlaylistCache {
    param(
        [object]$EnvironmentState,
        [switch]$ForceRefresh
    )

    if (-not (Test-Path -LiteralPath $script:helperCacheDir -PathType Container)) {
        New-Item -ItemType Directory -Path $script:helperCacheDir -Force | Out-Null
    }

    $cachePath = Join-Path $script:helperCacheDir 'playlists_cache.json'
    $useCache = $false
    if (-not $ForceRefresh.IsPresent -and (Test-Path -LiteralPath $cachePath -PathType Leaf)) {
        try {
            $age = (Get-Date) - (Get-Item -LiteralPath $cachePath).LastWriteTime
            if ($age.TotalHours -lt 24) {
                $useCache = $true
            }
        }
        catch {
        }
    }

    if ($useCache) {
        try {
            $cached = Get-Content -LiteralPath $cachePath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $playlists = @($cached.entries | Where-Object { $_.url -and $_.title })
            if ($playlists.Count -gt 0) {
                return [pscustomobject]@{
                    Source    = 'cache'
                    CachePath = $cachePath
                    Playlists = @($playlists)
                }
            }
        }
        catch {
            Write-HelperLog -Level 'WARN' -Message ("Could not read playlist cache: {0}" -f $_.Exception.Message)
        }
    }

    $args = New-Object System.Collections.Generic.List[string]
    if ($EnvironmentState.JsRuntime) {
        $args.Add('--js-runtimes')
        $args.Add([string]$EnvironmentState.JsRuntime)
    }
    foreach ($arg in (Get-HelperCookieArguments -EnvironmentState $EnvironmentState)) { $args.Add([string]$arg) }
    $args.Add('-J')
    $args.Add('--flat-playlist')
    $args.Add('https://www.youtube.com/feed/playlists')

    $response = Invoke-ExternalProcess -FilePath $EnvironmentState.YtDlpPath -Arguments @($args) -DisplayName 'yt-dlp my playlists' -LogStdOut:$false -LogStdErr:$true
    if ($response.ExitCode -ne 0) {
        throw "Failed to list playlists (exit code $($response.ExitCode))."
    }

    $feed = $response.StandardOut | ConvertFrom-Json -ErrorAction Stop
    try {
        $feed | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $cachePath -Encoding UTF8 -Force
    }
    catch {
        Write-HelperLog -Level 'WARN' -Message ("Could not write playlist cache: {0}" -f $_.Exception.Message)
    }

    return [pscustomobject]@{
        Source    = 'network'
        CachePath = $cachePath
        Playlists = @($feed.entries | Where-Object { $_.url -and $_.title })
    }
}

function Invoke-ListMyPlaylistsOperation {
    param([hashtable]$Request)

    $environment = Resolve-HelperEnvironment -DownloadRoot ([string]$Request.DownloadRoot) -UseCookies:$true -FirefoxProfilePath ([string]$Request.FirefoxProfilePath) -TreatWarningsAsNonFatal:([bool]$Request.TreatWarningsAsNonFatal)
    $playlistResult = Get-HelperPlaylistCache -EnvironmentState $environment -ForceRefresh:([bool]$Request.ForceRefreshCache)
    $rows = New-Object System.Collections.Generic.List[object]
    $index = 1
    foreach ($item in @($playlistResult.Playlists)) {
        $playlistId = ''
        if ($item.PSObject.Properties.Name -contains 'id' -and $item.id) {
            $playlistId = [string]$item.id
        }
        elseif ($item.url -match 'list=([^&]+)') {
            $playlistId = $matches[1]
        }

        $rows.Add([pscustomobject]@{
            Index      = $index
            Title      = [string]$item.title
            PlaylistId = $playlistId
            Url        = [string]$item.url
        })
        $index++
    }

    return (New-HelperOperationResult -Kind 'my-playlists' -Succeeded:$true -Status ('Loaded {0} playlists from {1}.' -f $rows.Count, $playlistResult.Source) -Summary 'My playlists loaded.' -Data ([pscustomobject]@{
        Source    = $playlistResult.Source
        CachePath = $playlistResult.CachePath
        Playlists = @($rows)
    }))
}

function Get-PlaylistMetafileSummary {
    param([string[]]$InfoJsonPaths)

    foreach ($path in $InfoJsonPaths) {
        try {
            $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $isPlaylist = $false
            if ($json.PSObject.Properties.Name -contains '_type' -and [string]$json._type -eq 'playlist') {
                $isPlaylist = $true
            }
            elseif ($json.PSObject.Properties.Name -contains 'entries' -and $json.entries) {
                $isPlaylist = $true
            }

            if (-not $isPlaylist) {
                continue
            }

            $entryCount = $null
            if ($json.PSObject.Properties.Name -contains 'playlist_count' -and $json.playlist_count) {
                $entryCount = [int]$json.playlist_count
            }
            elseif ($json.PSObject.Properties.Name -contains 'n_entries' -and $json.n_entries) {
                $entryCount = [int]$json.n_entries
            }
            elseif ($json.PSObject.Properties.Name -contains 'entries' -and $json.entries) {
                $entryCount = @($json.entries).Count
            }

            return [pscustomobject]@{
                Path       = $path
                Title      = if ($json.PSObject.Properties.Name -contains 'title') { [string]$json.title } else { '' }
                PlaylistId = if ($json.PSObject.Properties.Name -contains 'id') { [string]$json.id } else { '' }
                EntryCount = $entryCount
            }
        }
        catch {
        }
    }

    return $null
}

function Get-PlaylistFolderInspection {
    param([string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        throw 'Please select a folder to inspect.'
    }

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        throw "Folder not found: $FolderPath"
    }

    $resolvedPath = (Resolve-Path -LiteralPath $FolderPath).Path
    $archivePath = Join-Path $resolvedPath 'download_archive.txt'
    $sidecarDir = Join-Path $resolvedPath '_sidecar'
    $infoJsonFiles = @(Get-ChildItem -LiteralPath $resolvedPath -Filter '*.info.json' -File -Recurse -ErrorAction SilentlyContinue)
    $infoJsonPaths = @($infoJsonFiles | ForEach-Object { $_.FullName })
    $playlistMeta = Get-PlaylistMetafileSummary -InfoJsonPaths $infoJsonPaths
    $playlistMetaPath = if ($playlistMeta) { [string]$playlistMeta.Path } else { '' }

    $videoInfoJsonCount = @($infoJsonFiles | Where-Object { $_.FullName -ne $playlistMetaPath }).Count
    $subtitleFiles = @(Get-ChildItem -LiteralPath $resolvedPath -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.srt', '.vtt', '.ass', '.lrc' })
    $textTranscripts = @(Get-ChildItem -LiteralPath $resolvedPath -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq '.txt' -and $_.Name -notin @('download_archive.txt') })
    $mediaFiles = @(Get-ChildItem -LiteralPath $resolvedPath -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension.ToLowerInvariant() -in @('.mkv', '.mp4', '.webm', '.m4a', '.mp3', '.opus', '.wav', '.flac') })

    $archiveCount = $null
    if (Test-Path -LiteralPath $archivePath -PathType Leaf) {
        try {
            $archiveIds = @(Get-Content -LiteralPath $archivePath -ErrorAction Stop | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $archiveCount = $archiveIds.Count
        }
        catch {
        }
    }

    $folderName = Split-Path -Leaf $resolvedPath
    $folderPatternId = ''
    $folderPatternTitle = $folderName
    if ($folderName -match '^(?<title>.+?) \[(?<id>[^\]]+)\]$') {
        $folderPatternTitle = $matches['title']
        $folderPatternId = $matches['id']
    }

    $sources = New-Object System.Collections.Generic.List[string]
    if ($playlistMeta) { $sources.Add('playlist metadata file') }
    if (Test-Path -LiteralPath $archivePath -PathType Leaf) { $sources.Add('download archive') }
    if (Test-Path -LiteralPath $sidecarDir -PathType Container) { $sources.Add('sidecar folder') }
    if ($videoInfoJsonCount -gt 0) { $sources.Add('video info json files') }
    if ($mediaFiles.Count -gt 0) { $sources.Add('media files') }

    $recognized = $sources.Count -gt 0
    $certainty = 'Unknown'
    if ($playlistMeta -and $archiveCount -ne $null) {
        $certainty = 'Confirmed'
    }
    elseif ((Test-Path -LiteralPath $archivePath -PathType Leaf) -and ($mediaFiles.Count -gt 0 -or $videoInfoJsonCount -gt 0 -or (Test-Path -LiteralPath $sidecarDir -PathType Container))) {
        $certainty = 'Confirmed'
    }
    elseif ($recognized -or -not [string]::IsNullOrWhiteSpace($folderPatternId)) {
        $certainty = 'Inferred'
    }

    $notes = New-Object System.Collections.Generic.List[string]
    $completionState = 'Unknown'
    if ($playlistMeta -and $playlistMeta.EntryCount -ne $null -and $archiveCount -ne $null) {
        if ($archiveCount -ge $playlistMeta.EntryCount) {
            $completionState = 'Confirmed complete'
            $notes.Add(('Archive count matches or exceeds playlist metadata count ({0}/{1}).' -f $archiveCount, $playlistMeta.EntryCount))
        }
        else {
            $completionState = 'Inferred incomplete'
            $notes.Add(('Archive count is below playlist metadata count ({0}/{1}).' -f $archiveCount, $playlistMeta.EntryCount))
        }
    }
    elseif ($archiveCount -ne $null) {
        $completionState = 'Unknown'
        $notes.Add('A download archive is present, but no total playlist entry count metadata was found.')
    }
    else {
        $completionState = 'Unknown'
        $notes.Add('No download archive was found at the folder root.')
    }

    if (-not $playlistMeta) {
        $notes.Add('No explicit playlist metafile was found; title and playlist id may be inferred from folder naming only.')
    }

    $playlistTitle = if ($playlistMeta -and $playlistMeta.Title) { $playlistMeta.Title } else { $folderPatternTitle }
    $playlistId = if ($playlistMeta -and $playlistMeta.PlaylistId) { $playlistMeta.PlaylistId } else { $folderPatternId }
    $detectedSidecarCount = $videoInfoJsonCount + $subtitleFiles.Count + $textTranscripts.Count

    $summaryRows = @(
        (New-HelperFieldRow -Field 'Folder' -Value $resolvedPath),
        (New-HelperFieldRow -Field 'Recognized playlist folder' -Value ($(if ($recognized) { 'Yes' } else { 'No' }))),
        (New-HelperFieldRow -Field 'Certainty' -Value $certainty),
        (New-HelperFieldRow -Field 'Completion state' -Value $completionState),
        (New-HelperFieldRow -Field 'Playlist title' -Value $playlistTitle),
        (New-HelperFieldRow -Field 'Playlist id' -Value $playlistId),
        (New-HelperFieldRow -Field 'Archive entries' -Value ($(if ($archiveCount -ne $null) { [string]$archiveCount } else { '' }))),
        (New-HelperFieldRow -Field 'Media files' -Value ([string]$mediaFiles.Count)),
        (New-HelperFieldRow -Field 'Detected sidecars' -Value ([string]$detectedSidecarCount)),
        (New-HelperFieldRow -Field 'Video info json files' -Value ([string]$videoInfoJsonCount)),
        (New-HelperFieldRow -Field 'Subtitle files' -Value ([string]$subtitleFiles.Count)),
        (New-HelperFieldRow -Field 'Text transcripts' -Value ([string]$textTranscripts.Count)),
        (New-HelperFieldRow -Field 'Last modified' -Value ([string](Get-Item -LiteralPath $resolvedPath).LastWriteTime))
    )

    return (New-HelperOperationResult -Kind 'folder-inspector' -Succeeded:$true -Status 'Folder inspection completed.' -Summary ("Folder inspected -> {0}" -f $folderName) -Data ([pscustomobject]@{
        IsRecognizedPlaylistFolder = $recognized
        Sources                    = @($sources)
        Certainty                  = $certainty
        ArchiveEntryCount          = $archiveCount
        DetectedMediaCount         = $mediaFiles.Count
        DetectedSidecarCount       = $detectedSidecarCount
        PlaylistTitle              = $playlistTitle
        PlaylistId                 = $playlistId
        CompletionState            = $completionState
        Notes                      = @($notes)
        SummaryRows                = @($summaryRows)
        FolderPath                 = $resolvedPath
        ArchivePath                = if (Test-Path -LiteralPath $archivePath -PathType Leaf) { $archivePath } else { '' }
        PlaylistMetadataPath       = if ($playlistMeta) { $playlistMeta.Path } else { '' }
    }))
}

function Invoke-EnvironmentOperation {
    param([hashtable]$Request)

    $environment = Resolve-HelperEnvironment -DownloadRoot ([string]$Request.DownloadRoot) -UseCookies:([bool]$Request.UseCookies) -FirefoxProfilePath ([string]$Request.FirefoxProfilePath) -TreatWarningsAsNonFatal:([bool]$Request.TreatWarningsAsNonFatal)
    $local = Get-HelperLocalEnvironmentSnapshot

    return (New-HelperOperationResult -Kind 'environment' -Succeeded:$true -Status 'Environment prepared successfully.' -Summary 'Environment ready.' -Data ([pscustomobject]@{
        DownloadRoot       = $environment.DownloadRoot
        YtDlpPath          = $environment.YtDlpPath
        YtDlpVersion       = $environment.YtDlpVersion
        FfmpegBinPath      = $environment.FfmpegBinPath
        JsRuntime          = $environment.JsRuntime
        JsRuntimeLabel     = $environment.JsRuntimeLabel
        AuthReady          = $environment.AuthReady
        FirefoxProfilePath = $environment.FirefoxProfilePath
        Snapshot           = $local
    }))
}

function Invoke-HelperOperation {
    param([object]$Request)

    $requestData = ConvertTo-HelperHashtable -InputObject $Request
    if ($null -eq $requestData -or -not $requestData.Contains('Kind')) {
        throw 'Operation request is missing the Kind property.'
    }

    switch ([string]$requestData.Kind) {
        'single-video' { return (Invoke-SingleVideoDownloadOperation -Request $requestData) }
        'playlist-url' { return (Invoke-PlaylistDownloadOperation -Request $requestData) }
        'my-playlists' { return (Invoke-ListMyPlaylistsOperation -Request $requestData) }
        'folder-inspector' { return (Get-PlaylistFolderInspection -FolderPath ([string]$requestData.FolderPath)) }
        'environment' { return (Invoke-EnvironmentOperation -Request $requestData) }
        default { throw "Unknown helper operation kind: $($requestData.Kind)" }
    }
}
