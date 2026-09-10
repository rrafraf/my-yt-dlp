#Requires -Version 5.0
param(
    [string]$YtDlpPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'yt-dlp.exe')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

$script:repoRoot = Split-Path -Parent $PSScriptRoot
$script:logDir = Join-Path $PSScriptRoot 'logs'
$script:logPath = $null
$script:logSync = New-Object object
$script:loggingConfigPath = Join-Path $PSScriptRoot 'yt-research-gui.config.json'
$script:logLevels = @{
    DEBUG = 10
    INFO  = 20
    WARN  = 30
    ERROR = 40
}
$script:utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($false)
$script:currentLogLevel = 'INFO'
$script:logRetentionDays = 14
$script:guiDataRoot = Join-Path $PSScriptRoot 'data'
$script:guiAudioDir = Join-Path $script:guiDataRoot 'audio'
$script:guiTranscriptDir = Join-Path $script:guiDataRoot 'transcripts'
$script:guiLlmResultDir = Join-Path $script:guiDataRoot 'llm-results'
$script:guiTranscriptStudioDir = Join-Path $script:guiDataRoot 'transcript-studio'
$script:guiBundleDir = Join-Path $script:guiDataRoot 'video-bundles'
$script:guiPythonProjectRoot = $PSScriptRoot
$script:guiWhisperModule = 'yt_research_gui_whisper'
$script:guiWhisperApp = 'yt-research-gui-whisper'
$script:guiWhisperModel = 'turbo'
$script:guiWhisperLanguage = 'en'
$script:ollamaChunkPrefix = '[[OLLAMA_CHUNK_BASE64]]'
$script:ollamaCacheVersion = 'v2'
$script:localLlmToolRoot = Join-Path $script:repoRoot 'tools\local_llm_text'
$script:localLlmToolScript = Join-Path $script:localLlmToolRoot 'cli.py'
$script:transcriptStudioProjectRoot = Join-Path $script:repoRoot 'tools\transcript_studio'
$script:transcriptStudioModule = 'transcript_studio.cli'
$script:transcriptStudioApp = 'transcript-studio'
$script:ollamaModel = 'gemma4'
$script:ollamaTimeoutSeconds = 180
$script:ollamaModels = @()
$script:activeWhisperProcess = $null
$script:activeWhisperStdOutPath = ''
$script:activeWhisperStdErrPath = ''
$script:activeWhisperTranscriptPath = ''
$script:activeWhisperTimingPath = ''
$script:activeWhisperAudioPath = ''
$script:activeWhisperVideoId = ''
$script:activeWhisperOutputSummary = ''
$script:activeWhisperBundleRunId = ''
$script:whisperPollTimer = $null
$script:activeGemmaProcess = $null
$script:activeGemmaStdOutPath = ''
$script:activeGemmaStdErrPath = ''
$script:activeGemmaResultPath = ''
$script:activeGemmaInputPath = ''
$script:activeGemmaVideoId = ''
$script:activeGemmaPresetId = ''
$script:activeGemmaModel = ''
$script:activeGemmaTranscriptHash = ''
$script:activeGemmaBundleRunId = ''
$script:gemmaPollTimer = $null
$script:gemmaPresets = @()
$script:gemmaHelperAvailable = $false

. (Join-Path $PSScriptRoot 'processing_bundle.ps1')

function Load-LoggingSettings {
    $script:currentLogLevel = 'INFO'
    $script:logRetentionDays = 14
    $script:ollamaModel = 'gemma4'
    $script:ollamaTimeoutSeconds = 180

    if (-not (Test-Path -LiteralPath $script:loggingConfigPath -PathType Leaf)) {
        return
    }

    try {
        $config = Get-Content -LiteralPath $script:loggingConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        return
    }

    if ($null -eq $config) {
        return
    }

    $logging = $null
    if ($config.PSObject.Properties.Name -contains 'logging') {
        $logging = $config.logging
    }

    if ($null -ne $logging) {
        if ($logging.PSObject.Properties.Name -contains 'level') {
            $candidateLevel = ([string]$logging.level).Trim().ToUpperInvariant()
            if ($script:logLevels.ContainsKey($candidateLevel)) {
                $script:currentLogLevel = $candidateLevel
            }
        }

        if ($logging.PSObject.Properties.Name -contains 'retentionDays') {
            $candidateRetention = 0
            if ([int]::TryParse([string]$logging.retentionDays, [ref]$candidateRetention) -and $candidateRetention -ge 0) {
                $script:logRetentionDays = $candidateRetention
            }
        }
    }

    $ollama = $null
    if ($config.PSObject.Properties.Name -contains 'ollama') {
        $ollama = $config.ollama
    }

    if ($null -ne $ollama) {
        if ($ollama.PSObject.Properties.Name -contains 'model') {
            $candidateModel = ([string]$ollama.model).Trim()
            if (-not [string]::IsNullOrWhiteSpace($candidateModel)) {
                $script:ollamaModel = $candidateModel
            }
        }

        if ($ollama.PSObject.Properties.Name -contains 'timeoutSeconds') {
            $candidateTimeout = 0
            if ([int]::TryParse([string]$ollama.timeoutSeconds, [ref]$candidateTimeout) -and $candidateTimeout -gt 0) {
                $script:ollamaTimeoutSeconds = $candidateTimeout
            }
        }
    }
}

function Should-WriteLog {
    param(
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$Level
    )

    if (-not $script:logLevels.ContainsKey($script:currentLogLevel)) {
        $script:currentLogLevel = 'INFO'
    }

    return ($script:logLevels[$Level] -ge $script:logLevels[$script:currentLogLevel])
}

function Trim-LogFile {
    if ([string]::IsNullOrWhiteSpace($script:logPath) -or -not (Test-Path -LiteralPath $script:logPath -PathType Leaf)) {
        return
    }

    if ($script:logRetentionDays -le 0) {
        return
    }

    try {
        $content = Get-Content -LiteralPath $script:logPath -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($content)) {
            return
        }

        $normalized = $content -replace "`r`n", "`n"
        $pattern = '(?ms)^\[(?<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3})\] \[(?:DEBUG|INFO|WARN|ERROR)\] .*?(?=^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}\] \[(?:DEBUG|INFO|WARN|ERROR)\] |\z)'
        $matches = [System.Text.RegularExpressions.Regex]::Matches($normalized, $pattern)
        if ($matches.Count -eq 0) {
            return
        }

        $cutoff = (Get-Date).AddDays(-$script:logRetentionDays)
        $entriesToKeep = New-Object System.Collections.Generic.List[string]

        foreach ($entry in $matches) {
            $timestamp = [datetime]::MinValue
            $timestampText = $entry.Groups['ts'].Value
            $isParsed = [datetime]::TryParseExact(
                $timestampText,
                'yyyy-MM-dd HH:mm:ss.fff',
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$timestamp
            )

            if (-not $isParsed -or $timestamp -ge $cutoff) {
                $entriesToKeep.Add($entry.Value.TrimEnd("`n"))
            }
        }

        if ($entriesToKeep.Count -eq $matches.Count) {
            return
        }

        if ($entriesToKeep.Count -eq 0) {
            Set-Content -LiteralPath $script:logPath -Value '' -Encoding UTF8 -Force
            return
        }

        $newContent = ($entriesToKeep -join "`r`n") + "`r`n"
        Set-Content -LiteralPath $script:logPath -Value $newContent -Encoding UTF8 -Force
    }
    catch {
        # Log trimming is best-effort.
    }
}

function Initialize-Logger {
    try {
        if (-not (Test-Path -LiteralPath $script:logDir -PathType Container)) {
            New-Item -ItemType Directory -Path $script:logDir -Force | Out-Null
        }

        $script:logPath = Join-Path $script:logDir 'yt-research-gui.log'
        if (-not (Test-Path -LiteralPath $script:logPath -PathType Leaf)) {
            $null = New-Item -ItemType File -Path $script:logPath -Force
        }

        Trim-LogFile
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
        if (-not (Should-WriteLog -Level $Level)) {
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

Load-LoggingSettings
Initialize-Logger
Write-Log -Level 'INFO' -Message ("PowerShell version: {0}" -f $PSVersionTable.PSVersion.ToString())
Write-Log -Level 'INFO' -Message ("Logging configuration loaded. Level: {0} | RetentionDays: {1}" -f $script:currentLogLevel, $script:logRetentionDays)
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

function Test-YtDlpCookieClientError {
    param(
        [object]$Result,
        [string[]]$CookieArgs
    )

    if ($null -eq $Result -or $null -eq $CookieArgs -or $CookieArgs.Count -eq 0) {
        return $false
    }

    if ([int]$Result.ExitCode -eq 0) {
        return $false
    }

    return ([string]$Result.Output) -match [regex]::Escape('The following content is not available on this app')
}

function Invoke-YtDlpCaptureWithCookieFallback {
    param(
        [string]$ExePath,
        [string[]]$BaseArguments,
        [string[]]$CookieArgs,
        [string]$OperationName = 'yt-dlp command'
    )

    $firstAttemptArgs = if ($null -ne $CookieArgs -and $CookieArgs.Count -gt 0) {
        $CookieArgs + $BaseArguments
    }
    else {
        @($BaseArguments)
    }

    $initialResult = Invoke-YtDlpCapture -ExePath $ExePath -Arguments $firstAttemptArgs
    if (-not (Test-YtDlpCookieClientError -Result $initialResult -CookieArgs $CookieArgs)) {
        return [pscustomobject]@{
            Result             = $initialResult
            UsedCookieFallback = $false
        }
    }

    Write-Log -Level 'WARN' -Message ("{0} failed with Firefox cookies and YouTube returned the app-client error. Retrying without cookies." -f $OperationName)
    $retryResult = Invoke-YtDlpCapture -ExePath $ExePath -Arguments $BaseArguments
    if ([int]$retryResult.ExitCode -eq 0) {
        Write-Log -Level 'INFO' -Message ("{0} succeeded after retrying without Firefox cookies." -f $OperationName)
    }
    else {
        Write-Log -Level 'WARN' -Message ("{0} still failed after retrying without Firefox cookies. Retry exit code: {1}" -f $OperationName, $retryResult.ExitCode)
    }

    return [pscustomobject]@{
        Result             = $retryResult
        UsedCookieFallback = $true
    }
}

function Ensure-DirectoryExists {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Read-TextFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return ''
    }

    try {
        return [string](Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop)
    }
    catch {
        return ''
    }
}

function Get-GuiWhisperAudioFilePath {
    param([string]$VideoId)

    return (Join-Path $script:guiAudioDir ($VideoId + '.wav'))
}

function Get-GuiWhisperTranscriptPath {
    param([string]$VideoId)

    return (Join-Path $script:guiTranscriptDir ($VideoId + '.txt'))
}

function Get-GuiWhisperTimingPath {
    param([string]$VideoId)

    return (Join-Path $script:guiTranscriptDir ($VideoId + '.timings.json'))
}

function Get-GuiWhisperExistingAudioFile {
    param([string]$VideoId)

    $preferred = Get-GuiWhisperAudioFilePath -VideoId $VideoId
    if (Test-Path -LiteralPath $preferred -PathType Leaf) {
        return (Get-Item -LiteralPath $preferred -ErrorAction Stop)
    }

    return (
        Get-ChildItem -LiteralPath $script:guiAudioDir -File -Filter ($VideoId + '.*') -ErrorAction SilentlyContinue |
            Sort-Object -Property Name |
            Select-Object -First 1
    )
}

function Test-WhisperTranscriptCached {
    param([string]$VideoId)

    if ([string]::IsNullOrWhiteSpace($VideoId)) {
        return $false
    }

    $transcriptPath = Get-GuiWhisperTranscriptPath -VideoId $VideoId
    return (-not [string]::IsNullOrWhiteSpace((Read-TextFile -Path $transcriptPath)))
}

function Ensure-GuiWhisperStorage {
    Ensure-DirectoryExists -Path $script:guiDataRoot
    Ensure-DirectoryExists -Path $script:guiAudioDir
    Ensure-DirectoryExists -Path $script:guiTranscriptDir
    Ensure-DirectoryExists -Path $script:guiLlmResultDir
}

function Normalize-TranscriptSourceText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $normalized = (([string]$Text -replace "`r`n", "`n") -replace "`r", "`n")
    $lines = @($normalized -split "`n")
    $trimmedLines = foreach ($line in $lines) {
        $line.TrimEnd()
    }

    return (($trimmedLines -join "`n").Trim())
}

function Get-TextSha256 {
    param([string]$Text)

    $normalized = Normalize-TranscriptSourceText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalized)
        $hashBytes = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function ConvertTo-SafeCacheSegment {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 'default'
    }

    $safe = [regex]::Replace([string]$Value, '[^A-Za-z0-9._-]+', '-')
    $safe = $safe.Trim('-')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return 'default'
    }

    return $safe
}

function Get-GuiGemmaCacheDirectory {
    param([string]$VideoId)

    $videoSegment = if ([string]::IsNullOrWhiteSpace($VideoId)) { '_no-video-id' } else { (ConvertTo-SafeCacheSegment -Value $VideoId) }
    return (Join-Path $script:guiLlmResultDir $videoSegment)
}

function Get-GuiGemmaResultPath {
    param(
        [string]$VideoId,
        [string]$Model,
        [string]$PresetId,
        [string]$TranscriptHash
    )

    $cacheDir = Get-GuiGemmaCacheDirectory -VideoId $VideoId
    $fileName = '{0}--{1}--{2}--{3}.json' -f $script:ollamaCacheVersion, (ConvertTo-SafeCacheSegment -Value $Model), (ConvertTo-SafeCacheSegment -Value $PresetId), $TranscriptHash
    return (Join-Path $cacheDir $fileName)
}

function Get-GuiGemmaInputPath {
    param(
        [string]$VideoId,
        [string]$Model,
        [string]$PresetId,
        [string]$TranscriptHash
    )

    $cacheDir = Get-GuiGemmaCacheDirectory -VideoId $VideoId
    $fileName = '{0}--{1}--{2}--{3}.input.txt' -f $script:ollamaCacheVersion, (ConvertTo-SafeCacheSegment -Value $Model), (ConvertTo-SafeCacheSegment -Value $PresetId), $TranscriptHash
    return (Join-Path $cacheDir $fileName)
}

function New-OllamaModelEntry {
    param(
        [string]$Model,
        [string]$Display,
        [string]$Family = '',
        [string]$ParameterSize = '',
        [string]$QuantizationLevel = '',
        [string]$Source = 'ollama'
    )

    $resolvedModel = ([string]$Model).Trim()
    if ([string]::IsNullOrWhiteSpace($resolvedModel)) {
        return $null
    }

    $resolvedDisplay = ([string]$Display).Trim()
    if ([string]::IsNullOrWhiteSpace($resolvedDisplay)) {
        $resolvedDisplay = $resolvedModel
    }

    return [pscustomobject]@{
        model             = $resolvedModel
        display           = $resolvedDisplay
        family            = ([string]$Family).Trim()
        parameterSize     = ([string]$ParameterSize).Trim()
        quantizationLevel = ([string]$QuantizationLevel).Trim()
        source            = ([string]$Source).Trim()
    }
}

function Write-TextFileUtf8 {
    param(
        [string]$Path,
        [string]$Text
    )

    Ensure-DirectoryExists -Path (Split-Path -Parent $Path)
    [System.IO.File]::WriteAllText($Path, [string]$Text, $script:utf8NoBomEncoding)
}

function Get-ReadyGemmaStatusText {
    param($Result)

    if (-not $script:gemmaHelperAvailable) {
        return 'The Ollama helper is unavailable in this environment.'
    }

    if ($null -eq $Result) {
        return 'Load a video to use Ollama transcript post-processing.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$Result.Transcript)) {
        return 'No transcript available for Ollama post-processing yet.'
    }

    $selectedModel = Get-SelectedOllamaModel
    if ([string]::IsNullOrWhiteSpace($selectedModel)) {
        return 'No Ollama model is selected.'
    }

    return ("Ready to run Ollama model '{0}' on the current transcript." -f $selectedModel)
}

function Reset-ResultGemmaState {
    param(
        $Result,
        [string]$Status
    )

    if ($null -eq $Result) {
        return
    }

    $nextStatus = if ([string]::IsNullOrWhiteSpace($Status)) { Get-ReadyGemmaStatusText -Result $Result } else { $Status }
    $Result.GemmaDisplayText = ''
    $Result.GemmaActivityText = ''
    $Result.GemmaStatus = $nextStatus
    $Result.GemmaModel = ''
    $Result.GemmaPresetId = ''
    $Result.GemmaSourceTextHash = ''
    $Result.GemmaResultPath = ''
    if ($Result.PSObject.Properties.Name -contains 'BundleLlmResultPath') {
        $Result.BundleLlmResultPath = ''
    }
    if ($Result.PSObject.Properties.Name -contains 'BundlePromptPath') {
        $Result.BundlePromptPath = ''
    }
    if ($Result.PSObject.Properties.Name -contains 'BundleDisplayPath') {
        $Result.BundleDisplayPath = ''
    }
}

function Get-RepoFfmpegBinDirectory {
    $candidates = @(
        (Join-Path $script:repoRoot 'ffmpeg_yt-dlp\ffmpeg-master-latest-win64-gpl\bin'),
        (Join-Path $script:repoRoot 'ffmpeg_yt-dlp\bin')
    )

    foreach ($dir in $candidates) {
        if (Test-Path -LiteralPath (Join-Path $dir 'ffmpeg.exe') -PathType Leaf) {
            return $dir
        }
    }

    $ffmpegRoot = Join-Path $script:repoRoot 'ffmpeg_yt-dlp'
    if (Test-Path -LiteralPath $ffmpegRoot -PathType Container) {
        $ffmpegExe = Get-ChildItem -LiteralPath $ffmpegRoot -Filter 'ffmpeg.exe' -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($ffmpegExe) {
            return $ffmpegExe.Directory.FullName
        }
    }

    throw "FFmpeg was not found under '$ffmpegRoot'. Run the helper setup first."
}

function Get-LocalGuiPythonPath {
    $localPython = Join-Path $script:guiPythonProjectRoot '.venv\Scripts\python.exe'
    if (Test-Path -LiteralPath $localPython -PathType Leaf) {
        return $localPython
    }

    return $null
}

function Get-RepoPythonExecutablePath {
    $localPython = Get-LocalGuiPythonPath
    if ($localPython) {
        return $localPython
    }

    foreach ($name in @('py', 'python')) {
        $command = Get-Command -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) {
            if ($command.Source) {
                return $command.Source
            }

            return $command.Name
        }
    }

    return $null
}

function Get-UvExecutablePath {
    $uv = Get-Command -Name uv -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($uv) {
        if ($uv.Source) {
            return $uv.Source
        }

        return $uv.Name
    }

    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Links\uv.exe'),
        (Join-Path $env:USERPROFILE '.local\bin\uv.exe'),
        (Join-Path $env:USERPROFILE '.cargo\bin\uv.exe')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    return $null
}

function ConvertTo-ProcessArgumentString {
    param([string[]]$Arguments)

    $escapedArgs = foreach ($arg in $Arguments) {
        if ($arg -match '[\s"]') {
            '"' + ($arg -replace '"', '\"') + '"'
        }
        else {
            $arg
        }
    }

    return ($escapedArgs -join ' ')
}

function Get-GuiPythonInvocation {
    param(
        [string[]]$Arguments
    )

    $workingDirectory = $script:guiPythonProjectRoot
    $runnerDescription = ''
    $executable = ''
    $commandArgs = @()

    $localPython = Get-LocalGuiPythonPath
    if ($localPython) {
        $executable = $localPython
        $commandArgs = @('-u', '-m', $script:guiWhisperModule) + @($Arguments)
        $runnerDescription = "{0} {1}" -f $localPython, ($commandArgs -join ' ')
    }
    else {
        $uvPath = Get-UvExecutablePath
        if (-not $uvPath) {
            throw "uv was not found. Install uv or run 'uv sync' in '$workingDirectory' first."
        }

        $executable = $uvPath
        $commandArgs = @('run', '--project', $workingDirectory, $script:guiWhisperApp) + @($Arguments)
        $runnerDescription = "{0} {1}" -f $uvPath, ($commandArgs -join ' ')
    }

    return [pscustomobject]@{
        WorkingDirectory  = $workingDirectory
        Executable        = $executable
        CommandArgs       = $commandArgs
        RunnerDescription = $runnerDescription
        ArgumentString    = ConvertTo-ProcessArgumentString -Arguments $commandArgs
    }
}

function Get-LocalLlmToolInvocation {
    param(
        [string[]]$Arguments
    )

    if (-not (Test-Path -LiteralPath $script:localLlmToolScript -PathType Leaf)) {
        throw "Local LLM helper script was not found at '$($script:localLlmToolScript)'."
    }

    $pythonPath = Get-RepoPythonExecutablePath
    if (-not $pythonPath) {
        throw 'Python was not found. Install Python 3.10+ or run uv sync in the GUI folder first.'
    }

    $fileName = [System.IO.Path]::GetFileName([string]$pythonPath).ToLowerInvariant()
    $isPyLauncher = ($fileName -eq 'py' -or $fileName -eq 'py.exe')
    $commandArgs = if ($isPyLauncher) {
        @('-3', '-u', $script:localLlmToolScript) + @($Arguments)
    }
    else {
        @('-u', $script:localLlmToolScript) + @($Arguments)
    }

    return [pscustomobject]@{
        WorkingDirectory  = $script:repoRoot
        Executable        = $pythonPath
        CommandArgs       = $commandArgs
        RunnerDescription = "{0} {1}" -f $pythonPath, ($commandArgs -join ' ')
        ArgumentString    = ConvertTo-ProcessArgumentString -Arguments $commandArgs
    }
}

function Invoke-GuiPythonAppCapture {
    param(
        [string[]]$Arguments
    )

    $invocation = Get-GuiPythonInvocation -Arguments $Arguments
    Write-Log -Level 'INFO' -Message ("Running Whisper helper: {0}" -f $invocation.RunnerDescription)

    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $invocation.Executable
    $processInfo.Arguments = $invocation.ArgumentString
    $processInfo.WorkingDirectory = $invocation.WorkingDirectory
    $processInfo.UseShellExecute = $false
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.CreateNoWindow = $true
    $processInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $processInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8

    $process = [System.Diagnostics.Process]::Start($processInfo)
    if ($null -eq $process) {
        throw 'Failed to start the Whisper helper process.'
    }

    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    $exitCode = $process.ExitCode

    $rawOutput = @()
    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $rawOutput += ($stdout -split "`r?`n")
    }
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $rawOutput += ($stderr -split "`r?`n")
    }

    $lines = foreach ($item in $rawOutput) {
        if ($item -is [System.Management.Automation.ErrorRecord]) {
            $item.ToString()
        }
        else {
            [string]$item
        }
    }

    $outputText = ($lines -join [Environment]::NewLine).Trim()
    Write-Log -Level 'INFO' -Message "Whisper helper exit code: $exitCode"
    if (-not [string]::IsNullOrWhiteSpace($outputText)) {
        Write-Log -Level 'INFO' -Message ("Whisper helper output:`n{0}" -f $outputText)
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output   = $outputText
    }
}

function Start-GuiPythonApp {
    param(
        [string[]]$Arguments,
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    $invocation = Get-GuiPythonInvocation -Arguments $Arguments
    Ensure-DirectoryExists -Path (Split-Path -Parent $StdOutPath)
    Ensure-DirectoryExists -Path (Split-Path -Parent $StdErrPath)
    Write-Log -Level 'INFO' -Message ("Starting Whisper helper in background: {0}" -f $invocation.RunnerDescription)

    $process = Start-Process `
        -FilePath $invocation.Executable `
        -ArgumentList $invocation.ArgumentString `
        -WorkingDirectory $invocation.WorkingDirectory `
        -RedirectStandardOutput $StdOutPath `
        -RedirectStandardError $StdErrPath `
        -WindowStyle Hidden `
        -PassThru

    if ($null -eq $process) {
        throw 'Failed to start the Whisper helper process.'
    }

    return [pscustomobject]@{
        Process           = $process
        StdOutPath        = $StdOutPath
        StdErrPath        = $StdErrPath
        RunnerDescription = $invocation.RunnerDescription
    }
}

function Get-LocalTranscriptStudioPythonPath {
    foreach ($candidate in @(
        (Join-Path $script:transcriptStudioProjectRoot '.venv\Scripts\python.exe'),
        (Join-Path $script:transcriptStudioProjectRoot '.venv\Scripts\pythonw.exe')
    )) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    return $null
}

function Get-TranscriptStudioInvocation {
    param(
        [string[]]$Arguments
    )

    $workingDirectory = $script:transcriptStudioProjectRoot
    $runnerDescription = ''
    $executable = ''
    $commandArgs = @()

    $localPython = Get-LocalTranscriptStudioPythonPath
    if ($localPython) {
        $executable = $localPython
        $commandArgs = @('-m', $script:transcriptStudioModule) + @($Arguments)
        $runnerDescription = "{0} {1}" -f $localPython, ($commandArgs -join ' ')
    }
    else {
        $uvPath = Get-UvExecutablePath
        if (-not $uvPath) {
            throw "Transcript Studio dependencies are not installed. Run 'uv sync' in '$workingDirectory' first."
        }

        $executable = $uvPath
        $commandArgs = @('run', '--project', $workingDirectory, $script:transcriptStudioApp) + @($Arguments)
        $runnerDescription = "{0} {1}" -f $uvPath, ($commandArgs -join ' ')
    }

    return [pscustomobject]@{
        WorkingDirectory  = $workingDirectory
        Executable        = $executable
        CommandArgs       = $commandArgs
        RunnerDescription = $runnerDescription
        ArgumentString    = ConvertTo-ProcessArgumentString -Arguments $commandArgs
    }
}

function New-TranscriptStudioLaunchLogs {
    param([string]$SessionPath)

    Ensure-DirectoryExists -Path $script:logDir
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $sessionLabel = [System.IO.Path]::GetFileNameWithoutExtension([string]$SessionPath)
    if ([string]::IsNullOrWhiteSpace($sessionLabel)) {
        $sessionLabel = 'session'
    }

    $safeLabel = [System.Text.RegularExpressions.Regex]::Replace($sessionLabel, '[^A-Za-z0-9._-]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($safeLabel)) {
        $safeLabel = 'session'
    }

    return [pscustomobject]@{
        StdOutPath = Join-Path $script:logDir ("transcript-studio-{0}-{1}.stdout.log" -f $safeLabel, $timestamp)
        StdErrPath = Join-Path $script:logDir ("transcript-studio-{0}-{1}.stderr.log" -f $safeLabel, $timestamp)
    }
}

function Get-TranscriptStudioLaunchDiagnosticText {
    param(
        [string]$SessionPath,
        [string]$StdOutPath,
        [string]$StdErrPath,
        [object]$Process
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Transcript Studio exited before it opened a visible window.')
    if (-not [string]::IsNullOrWhiteSpace($SessionPath)) {
        $lines.Add("Session: $SessionPath")
    }

    if ($null -ne $Process) {
        try {
            $lines.Add(("Exit code: {0}" -f $Process.ExitCode))
        }
        catch {
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($StdOutPath)) {
        $lines.Add("stdout log: $StdOutPath")
    }
    if (-not [string]::IsNullOrWhiteSpace($StdErrPath)) {
        $lines.Add("stderr log: $StdErrPath")
    }

    $combinedOutput = Get-CombinedToolOutput -StdOutPath $StdOutPath -StdErrPath $StdErrPath
    $excerpt = Get-ActivityExcerptText -Text $combinedOutput -MaxLines 24 -MaxCharacters 2200
    if (-not [string]::IsNullOrWhiteSpace($excerpt)) {
        $lines.Add('')
        $lines.Add('Recent output:')
        $lines.Add($excerpt)
    }

    return (($lines.ToArray()) -join [Environment]::NewLine).Trim()
}

function Start-TranscriptStudioApp {
    param(
        [string]$SessionPath
    )

    if ([string]::IsNullOrWhiteSpace($SessionPath)) {
        throw 'Transcript Studio launch requires a session file path.'
    }

    $invocation = Get-TranscriptStudioInvocation -Arguments @('--session', $SessionPath)
    $launchLogs = New-TranscriptStudioLaunchLogs -SessionPath $SessionPath
    Write-Log -Level 'INFO' -Message ("Starting Transcript Studio: {0} | stdout: {1} | stderr: {2}" -f $invocation.RunnerDescription, $launchLogs.StdOutPath, $launchLogs.StdErrPath)

    $process = Start-Process `
        -FilePath $invocation.Executable `
        -ArgumentList $invocation.ArgumentString `
        -WorkingDirectory $invocation.WorkingDirectory `
        -RedirectStandardOutput $launchLogs.StdOutPath `
        -RedirectStandardError $launchLogs.StdErrPath `
        -PassThru

    if ($null -eq $process) {
        throw 'Failed to start Transcript Studio.'
    }

    Start-Sleep -Milliseconds 1200
    try {
        $process.Refresh()
    }
    catch {
    }

    if ($process.HasExited) {
        $diagnostic = Get-TranscriptStudioLaunchDiagnosticText -SessionPath $SessionPath -StdOutPath $launchLogs.StdOutPath -StdErrPath $launchLogs.StdErrPath -Process $process
        Write-Log -Level 'ERROR' -Message $diagnostic
        throw $diagnostic
    }

    return [pscustomobject]@{
        Process           = $process
        RunnerDescription = $invocation.RunnerDescription
        SessionPath       = $SessionPath
        StdOutPath        = $launchLogs.StdOutPath
        StdErrPath        = $launchLogs.StdErrPath
    }
}

function Ensure-TranscriptStudioStorage {
    Ensure-DirectoryExists -Path $script:guiTranscriptStudioDir
}

function ConvertFrom-JsonSafe {
    param([string]$JsonText)

    if ([string]::IsNullOrWhiteSpace($JsonText)) {
        return $null
    }

    try {
        return ($JsonText | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        return $null
    }
}

function Write-TranscriptStudioSessionFile {
    param($Result)

    if ($null -eq $Result) {
        throw 'Transcript Studio requires a loaded result.'
    }

    $videoId = ([string]$Result.VideoId).Trim()
    $bundleRoot = [string]$Result.BundleRoot
    if (-not [string]::IsNullOrWhiteSpace($bundleRoot) -and (Test-Path -LiteralPath $bundleRoot -PathType Container)) {
        $sessionPath = Join-Path $bundleRoot 'exports\transcript-studio.session.json'
    }
    else {
        Ensure-TranscriptStudioStorage
        $sessionFileName = if ([string]::IsNullOrWhiteSpace($videoId)) { 'transcript-session.json' } else { "$videoId.session.json" }
        $sessionPath = Join-Path $script:guiTranscriptStudioDir $sessionFileName
    }

    $sessionDirectory = Split-Path -Parent $sessionPath
    Ensure-DirectoryExists -Path $sessionDirectory
    $timingsPayload = ConvertFrom-JsonSafe -JsonText ([string]$Result.TranscriptTimingJson)
    $rawJsonPayload = ConvertFrom-JsonSafe -JsonText ([string]$Result.RawJson)
    $sessionAudioPath = if (-not [string]::IsNullOrWhiteSpace([string]$Result.BundleAudioPath)) { [string]$Result.BundleAudioPath } elseif (-not [string]::IsNullOrWhiteSpace([string]$Result.AudioPath)) { [string]$Result.AudioPath } else { '' }
    $sessionTranscriptPath = if (-not [string]::IsNullOrWhiteSpace([string]$Result.BundleTranscriptPath)) { [string]$Result.BundleTranscriptPath } elseif (-not [string]::IsNullOrWhiteSpace([string]$Result.TranscriptPath)) { [string]$Result.TranscriptPath } else { '' }
    $sessionTimingPath = if (-not [string]::IsNullOrWhiteSpace([string]$Result.BundleTimingPath)) { [string]$Result.BundleTimingPath } elseif (-not [string]::IsNullOrWhiteSpace([string]$Result.TranscriptTimingPath)) { [string]$Result.TranscriptTimingPath } else { '' }
    $sessionLlmResultPath = if (-not [string]::IsNullOrWhiteSpace([string]$Result.BundleLlmResultPath)) { [string]$Result.BundleLlmResultPath } else { [string]$Result.GemmaResultPath }
    $sessionPromptPath = [string]$Result.BundlePromptPath
    $sessionDisplayPath = [string]$Result.BundleDisplayPath
    $sessionMetadataPath = [string]$Result.BundleMetadataPath

    $audioPathValue = if ([string]::IsNullOrWhiteSpace($sessionAudioPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath $sessionAudioPath) }
    $transcriptPathValue = if ([string]::IsNullOrWhiteSpace($sessionTranscriptPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath $sessionTranscriptPath) }
    $timingPathValue = if ([string]::IsNullOrWhiteSpace($sessionTimingPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath $sessionTimingPath) }
    $metadataPathValue = if ([string]::IsNullOrWhiteSpace($sessionMetadataPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath $sessionMetadataPath) }
    $bundleManifestValue = if ([string]::IsNullOrWhiteSpace([string]$Result.BundleManifestPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath ([string]$Result.BundleManifestPath)) }

    $metadata = [ordered]@{
        transcriptStatus = [string]$Result.TranscriptStatus
        transcriptSource = [string]$Result.TranscriptSource
        transcriptMode   = [string]$Result.TranscriptMode
        metaRows         = @($Result.MetaRows)
        chapters         = @($Result.Chapters)
        rawJson          = if ($null -ne $rawJsonPayload) { $rawJsonPayload } else { [string]$Result.RawJson }
        bundleManifest   = $bundleManifestValue
        metadataJsonPath = $metadataPathValue
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Result.GemmaDisplayText) -or -not [string]::IsNullOrWhiteSpace([string]$Result.GemmaStatus)) {
        $metadata.ollama = [ordered]@{
            model          = [string]$Result.GemmaModel
            presetId       = [string]$Result.GemmaPresetId
            displayText    = [string]$Result.GemmaDisplayText
            status         = [string]$Result.GemmaStatus
            resultPath     = if ([string]::IsNullOrWhiteSpace($sessionLlmResultPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath $sessionLlmResultPath) }
            promptPath     = if ([string]::IsNullOrWhiteSpace($sessionPromptPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath $sessionPromptPath) }
            displayPath    = if ([string]::IsNullOrWhiteSpace($sessionDisplayPath)) { '' } else { (Get-RelativePathFromBase -BasePath $sessionDirectory -TargetPath $sessionDisplayPath) }
            sourceTextHash = [string]$Result.GemmaSourceTextHash
        }
    }

    $payload = [ordered]@{
        sessionVersion = 1
        title          = [string]$Result.Title
        description    = [string]$Result.Description
        subtitle       = [string]$Result.TranscriptStatus
        sourceKind     = 'youtube'
        sourceId       = $videoId
        sourceUrl      = [string]$Result.VideoUrl
        audioPath      = $audioPathValue
        transcriptPath = $transcriptPathValue
        timingsPath    = $timingPathValue
        transcriptText = [string]$Result.Transcript
        timings        = if ($null -ne $timingsPayload) { $timingsPayload } else { $null }
        metadata       = $metadata
    }

    $sessionJson = $payload | ConvertTo-Json -Depth 30
    Write-TextFileUtf8 -Path $sessionPath -Text $sessionJson
    if (-not [string]::IsNullOrWhiteSpace($videoId)) {
        $manifest = Get-VideoBundleManifest -VideoId $videoId -Title ([string]$Result.Title) -SourceUrl ([string]$Result.VideoUrl)
        $manifest.counters.transcriptStudioLaunches = [int]$manifest.counters.transcriptStudioLaunches + 1
        Set-ManifestLatestPath -Manifest $manifest -PropertyName 'transcriptStudioSession' -VideoId $videoId -AbsolutePath $sessionPath
        Upsert-BundleArtifact -Manifest $manifest -Artifact (New-BundleArtifactRecord -VideoId $videoId -Path $sessionPath -Kind 'transcript-studio-session' -Role 'export')
        Save-VideoBundleManifest -VideoId $videoId -Manifest $manifest
        Add-BundleEvent -VideoId $videoId -Type 'transcript_studio_session_exported' -Outputs ([ordered]@{
            session = (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $videoId) -TargetPath $sessionPath)
        }) -Details ([ordered]@{
            transcriptMode = [string]$Result.TranscriptMode
        })
    }
    $Result.BundleSessionPath = [string]$sessionPath
    return $sessionPath
}

function Invoke-LocalLlmToolCapture {
    param(
        [string[]]$Arguments
    )

    $invocation = Get-LocalLlmToolInvocation -Arguments $Arguments
    Write-Log -Level 'INFO' -Message ("Running local LLM helper: {0}" -f $invocation.RunnerDescription)

    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $invocation.Executable
    $processInfo.Arguments = $invocation.ArgumentString
    $processInfo.WorkingDirectory = $invocation.WorkingDirectory
    $processInfo.UseShellExecute = $false
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.CreateNoWindow = $true
    $processInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $processInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8

    $process = [System.Diagnostics.Process]::Start($processInfo)
    if ($null -eq $process) {
        throw 'Failed to start the local LLM helper process.'
    }

    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    $exitCode = $process.ExitCode
    $outputText = @($stdout, $stderr) -join [Environment]::NewLine
    $outputText = (([string]$outputText -replace "`r`n", "`n") -replace "`r", "`n").Trim()

    Write-Log -Level 'INFO' -Message "Local LLM helper exit code: $exitCode"
    if (-not [string]::IsNullOrWhiteSpace($outputText)) {
        Write-Log -Level 'INFO' -Message ("Local LLM helper output:`n{0}" -f $outputText)
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output   = $outputText
    }
}

function Start-LocalLlmTool {
    param(
        [string[]]$Arguments,
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    $invocation = Get-LocalLlmToolInvocation -Arguments $Arguments
    Ensure-DirectoryExists -Path (Split-Path -Parent $StdOutPath)
    Ensure-DirectoryExists -Path (Split-Path -Parent $StdErrPath)
    Write-Log -Level 'INFO' -Message ("Starting local LLM helper in background: {0}" -f $invocation.RunnerDescription)

    $process = Start-Process `
        -FilePath $invocation.Executable `
        -ArgumentList $invocation.ArgumentString `
        -WorkingDirectory $invocation.WorkingDirectory `
        -RedirectStandardOutput $StdOutPath `
        -RedirectStandardError $StdErrPath `
        -WindowStyle Hidden `
        -PassThru

    if ($null -eq $process) {
        throw 'Failed to start the local LLM helper process.'
    }

    return [pscustomobject]@{
        Process           = $process
        StdOutPath        = $StdOutPath
        StdErrPath        = $StdErrPath
        RunnerDescription = $invocation.RunnerDescription
    }
}

function Get-GemmaPresetDefinitions {
    $response = Invoke-LocalLlmToolCapture -Arguments @('list-presets')
    if ([int]$response.ExitCode -ne 0) {
        throw "Failed to load Gemma presets.`n$($response.Output)"
    }

    if ([string]::IsNullOrWhiteSpace([string]$response.Output)) {
        throw 'Gemma preset listing returned empty output.'
    }

    $presets = $response.Output | ConvertFrom-Json -ErrorAction Stop
    return @($presets)
}

function Get-OllamaModelDefinitions {
    $response = Invoke-LocalLlmToolCapture -Arguments @('list-models', '--timeout-seconds', '15')
    if ([int]$response.ExitCode -ne 0) {
        throw "Failed to load installed Ollama models.`n$($response.Output)"
    }

    if ([string]::IsNullOrWhiteSpace([string]$response.Output)) {
        throw 'Ollama model listing returned empty output.'
    }

    $models = $response.Output | ConvertFrom-Json -ErrorAction Stop
    return @($models)
}

function Get-SelectedOllamaModel {
    if ($null -eq $OllamaModelComboBox) {
        return ''
    }

    $selected = $OllamaModelComboBox.SelectedItem
    if ($null -eq $selected) {
        return ''
    }

    $model = ''
    if ($selected.PSObject.Properties.Name -contains 'model') {
        $model = [string]$selected.model
    }
    elseif ($selected.PSObject.Properties.Name -contains 'name') {
        $model = [string]$selected.name
    }

    return $model.Trim()
}

function Get-SelectedGemmaPresetId {
    if ($null -eq $GemmaPresetComboBox) {
        return ''
    }

    $selected = $GemmaPresetComboBox.SelectedItem
    if ($null -eq $selected) {
        return ''
    }

    return [string]$selected.id
}

function Get-SelectedGemmaPresetDefinition {
    if ($null -eq $GemmaPresetComboBox) {
        return $null
    }

    return $GemmaPresetComboBox.SelectedItem
}

function Update-GemmaPresetTemplateUi {
    $preset = Get-SelectedGemmaPresetDefinition
    if ($null -eq $preset) {
        $GemmaPresetDetailsTextBox.Text = 'Select a prompt preset to inspect the exact template instructions that will be sent with the transcript.'
        return
    }

    $label = ([string]$preset.label).Trim()
    $description = ([string]$preset.description).Trim()
    $instructions = ([string]$preset.instructions).Trim()
    $parts = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($label)) {
        $parts.Add("Preset: $label")
    }

    if (-not [string]::IsNullOrWhiteSpace($description)) {
        $parts.Add('')
        $parts.Add("Description`n$description")
    }

    if (-not [string]::IsNullOrWhiteSpace($instructions)) {
        $parts.Add('')
        $parts.Add("Instructions`n$instructions")
    }

    $GemmaPresetDetailsTextBox.Text = ($parts -join [Environment]::NewLine).Trim()
    $GemmaPresetDetailsTextBox.ScrollToHome()
    $GemmaPresetDetailsExpander.IsExpanded = $true
}

function Read-GemmaResultFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Gemma result file was not found at '$Path'."
    }

    $content = Read-TextFile -Path $Path
    if ([string]::IsNullOrWhiteSpace($content)) {
        throw "Gemma result file is empty: '$Path'."
    }

    return ($content | ConvertFrom-Json -ErrorAction Stop)
}

function Get-CachedGemmaResult {
    param(
        [string]$VideoId,
        [string]$Model,
        [string]$PresetId,
        [string]$TranscriptHash
    )

    if ([string]::IsNullOrWhiteSpace($VideoId) -or [string]::IsNullOrWhiteSpace($Model) -or [string]::IsNullOrWhiteSpace($PresetId) -or [string]::IsNullOrWhiteSpace($TranscriptHash)) {
        return $null
    }

    $resultPath = Get-GuiGemmaResultPath -VideoId $VideoId -Model $Model -PresetId $PresetId -TranscriptHash $TranscriptHash
    if (-not (Test-Path -LiteralPath $resultPath -PathType Leaf)) {
        return $null
    }

    try {
        $payload = Read-GemmaResultFile -Path $resultPath
    }
    catch {
        Write-Log -Level 'WARN' -Message ("Failed to parse cached Gemma result '{0}':`n{1}" -f $resultPath, (Get-ErrorDetails -ErrorObject $_))
        return $null
    }

    if ([string]$payload.sourceTextHash -ne $TranscriptHash -or [string]$payload.presetId -ne $PresetId -or [string]$payload.model -ne $Model) {
        Write-Log -Level 'WARN' -Message ("Ignoring cached Gemma result because its metadata did not match the current request. Path: {0}" -f $resultPath)
        return $null
    }

    return [pscustomobject]@{
        ResultPath = $resultPath
        Payload    = $payload
    }
}

function Get-CachedWhisperTranscriptResult {
    param([string]$VideoId)

    if ([string]::IsNullOrWhiteSpace($VideoId)) {
        return $null
    }

    $transcriptPath = Get-GuiWhisperTranscriptPath -VideoId $VideoId
    $cachedTranscript = Read-TextFile -Path $transcriptPath
    if ([string]::IsNullOrWhiteSpace($cachedTranscript)) {
        return $null
    }

    Write-Log -Level 'INFO' -Message "Using cached Whisper transcript for video ID: $VideoId"
    $audioFile = Get-GuiWhisperExistingAudioFile -VideoId $VideoId
    $timingPath = Get-GuiWhisperTimingPath -VideoId $VideoId
    $timingJson = Read-TextFile -Path $timingPath
    return [pscustomobject]@{
        Text           = $cachedTranscript
        Status         = 'Loaded cached local Whisper transcript.'
        Source         = Split-Path -Leaf $transcriptPath
        Output         = ''
        AudioPath      = if ($audioFile) { $audioFile.FullName } else { '' }
        TranscriptPath = $transcriptPath
        TimingJson     = [string]$timingJson
        TimingPath     = if (Test-Path -LiteralPath $timingPath -PathType Leaf) { $timingPath } else { '' }
        UsedCache      = $true
    }
}

function Get-WhisperTranscriptResult {
    param(
        [string]$TranscriptPath,
        [string]$TimingPath,
        [string]$AudioPath,
        [string]$Output = '',
        [bool]$UsedCache = $false
    )

    $text = Read-TextFile -Path $TranscriptPath
    $timingJson = Read-TextFile -Path $TimingPath
    $status = if ([string]::IsNullOrWhiteSpace($text)) {
        'Whisper completed, but no readable text was produced.'
    }
    elseif ($UsedCache) {
        'Loaded cached local Whisper transcript.'
    }
    else {
        'Transcript generated locally with Whisper.'
    }

    return [pscustomobject]@{
        Text           = $text
        Status         = $status
        Source         = Split-Path -Leaf $TranscriptPath
        Output         = [string]$Output
        AudioPath      = [string]$AudioPath
        TranscriptPath = $TranscriptPath
        TimingJson     = [string]$timingJson
        TimingPath     = if (Test-Path -LiteralPath $TimingPath -PathType Leaf) { $TimingPath } else { '' }
        UsedCache      = $UsedCache
    }
}

function Prepare-WhisperTranscription {
    param(
        [string]$ExePath,
        [string]$Url,
        [string]$VideoId,
        [string[]]$CookieArgs
    )

    if ([string]::IsNullOrWhiteSpace($VideoId)) {
        throw 'Video ID is required for Whisper transcription.'
    }

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw 'Video URL is required for Whisper transcription.'
    }

    Ensure-GuiWhisperStorage

    $cachedResult = Get-CachedWhisperTranscriptResult -VideoId $VideoId
    if ($null -ne $cachedResult) {
        return [pscustomobject]@{
            Mode   = 'cached'
            Result = $cachedResult
        }
    }

    $audioFile = Get-GuiWhisperExistingAudioFile -VideoId $VideoId
    $audioWasExtracted = $false
    $audioExtractionOutput = ''
    $audioExtractionUsedCookieFallback = $false
    if ($null -eq $audioFile) {
        $ffmpegBin = Get-RepoFfmpegBinDirectory
        $outputTemplate = Join-Path $script:guiAudioDir '%(id)s.%(ext)s'
        $audioArgs = @(
            '--no-update',
            '--no-playlist',
            '--no-warnings',
            '--extract-audio',
            '--audio-format', 'wav',
            '--ffmpeg-location', $ffmpegBin,
            '--output', $outputTemplate,
            $Url
        )

        $audioRequest = Invoke-YtDlpCaptureWithCookieFallback -ExePath $ExePath -BaseArguments $audioArgs -CookieArgs $CookieArgs -OperationName 'Whisper audio extraction'
        $audioResult = $audioRequest.Result
        if ($audioResult.ExitCode -ne 0) {
            throw "Audio extraction failed (exit code $($audioResult.ExitCode)).`n$($audioResult.Output)"
        }

        $audioFile = Get-GuiWhisperExistingAudioFile -VideoId $VideoId
        $audioWasExtracted = $true
        $audioExtractionOutput = [string]$audioResult.Output
        $audioExtractionUsedCookieFallback = [bool]$audioRequest.UsedCookieFallback
        if ($audioFile) {
            Write-BundleAudioExtractionArtifacts -VideoId $VideoId -AudioPath $audioFile.FullName -ExtractionLog $audioExtractionOutput -UsedCookieFallback $audioExtractionUsedCookieFallback | Out-Null
        }
    }

    if ($null -eq $audioFile) {
        throw "Expected extracted audio file for video ID '$VideoId' was not found."
    }

    $transcriptPath = Get-GuiWhisperTranscriptPath -VideoId $VideoId
    $timingPath = Get-GuiWhisperTimingPath -VideoId $VideoId
    $ffmpegBin = Get-RepoFfmpegBinDirectory
    $whisperArgs = @(
        '--audio', $audioFile.FullName,
        '--output', $transcriptPath,
        '--timings-output', $timingPath,
        '--ffmpeg-dir', $ffmpegBin,
        '--model', $script:guiWhisperModel,
        '--language', $script:guiWhisperLanguage
    )

    return [pscustomobject]@{
        Mode                       = 'launch'
        VideoId                    = $VideoId
        AudioPath                  = $audioFile.FullName
        TranscriptPath             = $transcriptPath
        TimingPath                 = $timingPath
        Arguments                  = $whisperArgs
        AudioWasExtracted          = $audioWasExtracted
        AudioExtractionOutput      = $audioExtractionOutput
        AudioExtractionUsedCookieFallback = $audioExtractionUsedCookieFallback
    }
}

function Start-WhisperTranscription {
    param(
        [string]$ExePath,
        [string]$Url,
        [string]$VideoId,
        [string[]]$CookieArgs
    )

    $prep = Prepare-WhisperTranscription -ExePath $ExePath -Url $Url -VideoId $VideoId -CookieArgs $CookieArgs
    if ([string]$prep.Mode -eq 'cached') {
        return $prep
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $stdoutPath = Join-Path $script:logDir ("whisper-{0}-{1}.stdout.log" -f $VideoId, $timestamp)
    $stderrPath = Join-Path $script:logDir ("whisper-{0}-{1}.stderr.log" -f $VideoId, $timestamp)
    $processStart = Start-GuiPythonApp -Arguments $prep.Arguments -StdOutPath $stdoutPath -StdErrPath $stderrPath
    return [pscustomobject]@{
        Mode            = 'launch'
        VideoId         = $VideoId
        AudioPath       = [string]$prep.AudioPath
        TranscriptPath  = [string]$prep.TranscriptPath
        TimingPath      = [string]$prep.TimingPath
        Process         = $processStart.Process
        StdOutPath      = $processStart.StdOutPath
        StdErrPath      = $processStart.StdErrPath
        RunnerDescription = $processStart.RunnerDescription
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

        $fetchRequest = Invoke-YtDlpCaptureWithCookieFallback -ExePath $ExePath -BaseArguments $args -CookieArgs $CookieArgs -OperationName 'Subtitle fetch'
        $fetchResult = $fetchRequest.Result

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
                Text                  = ''
                Status                = $status
                Source                = ''
                Output                = $fetchResult.Output
                CaptionContent        = ''
                CaptionExtension      = ''
                UsedCookieFallback    = [bool]$fetchRequest.UsedCookieFallback
                SubtitleCandidateCount = [int]$subtitleCount
            }
        }

        $bestFile = Get-BestSubtitleFile -Files $subtitleFiles
        Write-Log -Level 'INFO' -Message "Selected subtitle file: $($bestFile.FullName)"
        $text = Convert-SubtitleFileToText -Path $bestFile.FullName
        $captionContent = Read-TextFile -Path $bestFile.FullName
        Write-Log -Level 'INFO' -Message ("Transcript character count: {0}" -f $text.Length)

        $statusMessage = if ([string]::IsNullOrWhiteSpace($text)) {
            "Subtitle file found but no readable lines extracted ($($bestFile.Name))."
        } else {
            "Transcript extracted from $($bestFile.Name)."
        }

        return [pscustomobject]@{
            Text                  = $text
            Status                = $statusMessage
            Source                = $bestFile.Name
            Output                = $fetchResult.Output
            CaptionContent        = [string]$captionContent
            CaptionExtension      = [string]$bestFile.Extension
            UsedCookieFallback    = [bool]$fetchRequest.UsedCookieFallback
            SubtitleCandidateCount = [int]$subtitleCount
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

    $cookieArgs = @(Get-CookieArgs -UseCookies $UseCookies -ProfilePath $ProfilePath)
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

    $metaRequest = Invoke-YtDlpCaptureWithCookieFallback -ExePath $ExePath -BaseArguments $metadataArgs -CookieArgs $cookieArgs -OperationName 'Metadata fetch'
    $metaResult = $metaRequest.Result
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

    $resolvedVideoUrl = if ([string]::IsNullOrWhiteSpace([string]$info.webpage_url)) { [string]$Url } else { [string]$info.webpage_url }
    $bundleArtifacts = Write-BundleFetchArtifacts `
        -VideoId ([string]$info.id) `
        -Title ([string]$info.title) `
        -SourceUrl $resolvedVideoUrl `
        -MetadataJson ([string]$rawJsonPretty) `
        -MetadataFetchLog ([string]$metaResult.Output) `
        -SubtitleFetchLog ([string]$transcript.Output) `
        -TranscriptText ([string]$transcript.Text) `
        -TranscriptStatus ([string]$transcript.Status) `
        -TranscriptSource ([string]$transcript.Source) `
        -CaptionContent ([string]$transcript.CaptionContent) `
        -CaptionExtension ([string]$transcript.CaptionExtension) `
        -UseCookies $UseCookies `
        -ProfilePath $ProfilePath `
        -MetadataUsedCookieFallback ([bool]$metaRequest.UsedCookieFallback) `
        -SubtitleUsedCookieFallback ([bool]$transcript.UsedCookieFallback) `
        -SubtitleCandidateCount ([int]$transcript.SubtitleCandidateCount)

    return [pscustomobject]@{
        Title            = [string]$info.title
        Description      = [string]$info.description
        MetaRows         = $metaRows
        Chapters         = $chapterRows
        Transcript       = [string]$transcript.Text
        TranscriptStatus = [string]$transcript.Status
        TranscriptSource = [string]$transcript.Source
        TranscriptMode   = if ([string]::IsNullOrWhiteSpace([string]$transcript.Text)) { 'none' } else { 'youtube' }
        VideoId          = [string]$info.id
        VideoUrl         = $resolvedVideoUrl
        AudioPath        = ''
        TranscriptPath   = ''
        TranscriptTimingJson = ''
        TranscriptTimingPath = ''
        BundleRoot       = [string]$bundleArtifacts.BundleRoot
        BundleManifestPath = [string]$bundleArtifacts.ManifestPath
        BundleEventsPath = [string]$bundleArtifacts.EventsPath
        BundleMetadataPath = [string]$bundleArtifacts.MetadataPath
        BundleCaptionPath = [string]$bundleArtifacts.CaptionPath
        BundleTranscriptPath = [string]$bundleArtifacts.TranscriptPath
        BundleAudioPath = ''
        BundleTimingPath = ''
        BundleLlmResultPath = ''
        BundlePromptPath = ''
        BundleDisplayPath = ''
        BundleSessionPath = ''
        GemmaDisplayText = ''
        GemmaActivityText = ''
        GemmaStatus      = ''
        GemmaModel       = ''
        GemmaPresetId    = ''
        GemmaSourceTextHash = ''
        GemmaResultPath  = ''
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

function Resolve-YoutubeInputToUrl {
    param([string]$InputText)

    $candidate = ([string]$InputText).Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        throw 'Please enter a YouTube video ID or URL.'
    }

    if ($candidate -match '^[A-Za-z0-9_-]{11}$') {
        return ("https://www.youtube.com/watch?v={0}" -f $candidate)
    }

    return $candidate
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
        Title="YT Research GUI" Height="920" Width="1480" MinHeight="760" MinWidth="1160"
        WindowStartupLocation="CenterScreen">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="420"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Video ID or URL:" VerticalAlignment="Center" Margin="0,0,8,0" FontWeight="SemiBold"/>
            <TextBox Grid.Column="1" x:Name="UrlBox" Height="30" Margin="0,0,8,0" VerticalContentAlignment="Center"/>
            <Button Grid.Column="2" x:Name="FetchButton" Width="120" Height="30" Content="Fetch"/>
            <WrapPanel Grid.Column="4" HorizontalAlignment="Right">
                <Button x:Name="CopyTitleButton" Content="Copy Title" Width="110" Height="28" Margin="0,0,8,0"/>
                <Button x:Name="CopyDescriptionButton" Content="Copy Description" Width="130" Height="28" Margin="0,0,8,0"/>
                <Button x:Name="CopyTranscriptButton" Content="Copy Transcript" Width="130" Height="28" Margin="0,0,8,0"/>
                <Button x:Name="CopyJsonButton" Content="Copy Raw JSON" Width="120" Height="28"/>
            </WrapPanel>
        </Grid>

        <Expander Grid.Row="1" x:Name="BrowserOptionsExpander" Header="Browser Options" IsExpanded="False" Margin="0,0,0,8">
            <Grid Margin="8,6,8,8">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <CheckBox Grid.Column="0" x:Name="UseCookiesCheck" VerticalAlignment="Center" Content="Use Firefox cookies" Margin="0,0,16,0"/>
                <TextBlock Grid.Column="1" Text="Profile folder:" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBox Grid.Column="2" x:Name="ProfilePathBox" Height="28" Margin="0,0,8,0" VerticalContentAlignment="Center"/>
                <Button Grid.Column="3" x:Name="BrowseProfileButton" Width="110" Height="28" Content="Browse..."/>
            </Grid>
        </Expander>

        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="360"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <TabControl Grid.Column="0" Margin="0,0,8,0">
                <TabItem Header="Overview">
                    <Grid Margin="8">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" x:Name="TitleTextBlock" Text="No video loaded." FontSize="18" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,4"/>
                        <TextBlock Grid.Row="1" x:Name="VideoIdTextBlock" Text="Fetch a video to load metadata." Foreground="DimGray" Margin="0,0,0,10" TextWrapping="Wrap"/>
                        <DataGrid Grid.Row="2" x:Name="MetaGrid" IsReadOnly="True" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Field" Binding="{Binding Field}" Width="140"/>
                                <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="*"/>
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

                <TabItem Header="Chapters">
                    <Grid Margin="8">
                        <DataGrid x:Name="ChaptersGrid" IsReadOnly="True" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Start" Binding="{Binding Start}" Width="90"/>
                                <DataGridTextColumn Header="End" Binding="{Binding End}" Width="90"/>
                                <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
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

            <GridSplitter Grid.Column="1" Width="8" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="#D9D9D9" ShowsPreview="True"/>

            <TabControl Grid.Column="2">
                <TabItem Header="Transcript Workspace">
                    <Grid Margin="8">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <Grid Grid.Row="0" Margin="0,0,0,6">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" x:Name="TranscriptStatusText" Margin="0,0,8,0" TextWrapping="Wrap"/>
                            <Button Grid.Column="1" x:Name="WhisperTranscriptButton" Width="190" Height="28" Margin="0,0,8,0" Content="Transcribe with Whisper" IsEnabled="False"/>
                            <Button Grid.Column="2" x:Name="OpenTranscriptStudioButton" Width="136" Height="28" Content="Open In Studio" IsEnabled="False"/>
                        </Grid>

                        <TextBlock Grid.Row="1" x:Name="WhisperProgressText" Margin="0,0,0,8" Foreground="DimGray" Text="Whisper idle." TextWrapping="Wrap"/>

                        <Grid Grid.Row="2" Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="220"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="240"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="Ollama Model:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,8,0"/>
                            <ComboBox Grid.Column="1" x:Name="OllamaModelComboBox" Height="28" Margin="0,0,12,0" IsEnabled="False"/>
                            <TextBlock Grid.Column="2" Text="Prompt Preset:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,8,0"/>
                            <ComboBox Grid.Column="3" x:Name="GemmaPresetComboBox" Height="28" Margin="0,0,12,0" IsEnabled="False"/>
                            <Button Grid.Column="4" x:Name="RunGemmaButton" Width="150" Height="28" Margin="0,0,8,0" Content="Run with Ollama" IsEnabled="False"/>
                            <Button Grid.Column="5" x:Name="StopGemmaButton" Width="90" Height="28" Content="Stop" IsEnabled="False"/>
                        </Grid>

                        <Expander Grid.Row="3" x:Name="GemmaPresetDetailsExpander" Header="Selected Prompt Template" IsExpanded="True" Margin="0,0,0,8">
                            <TextBox x:Name="GemmaPresetDetailsTextBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                                     FontFamily="Consolas" FontSize="12" MinHeight="96" MaxHeight="180" Padding="8"/>
                        </Expander>

                        <Grid Grid.Row="4">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="8"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <GroupBox Grid.Column="0" Header="Transcript Text">
                                <TextBox x:Name="TranscriptTextBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                         VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                                         FontFamily="Consolas" FontSize="13"/>
                            </GroupBox>
                            <GridSplitter Grid.Column="1" Width="8" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="#D9D9D9" ShowsPreview="True"/>
                            <GroupBox Grid.Column="2" Header="Ollama Result">
                                <DockPanel>
                                    <TextBlock DockPanel.Dock="Top" x:Name="GemmaStatusText" Margin="8,8,8,0" Foreground="DimGray" Text="Ollama idle." TextWrapping="Wrap"/>
                                    <TextBox x:Name="GemmaResultTextBox" Margin="8" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                                             FontFamily="Consolas" FontSize="13"/>
                                </DockPanel>
                            </GroupBox>
                        </Grid>

                        <Grid Grid.Row="5" Margin="0,8,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="12"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Expander Grid.Column="0" x:Name="WhisperActivityExpander" Header="Whisper Activity" IsExpanded="False">
                                <TextBox x:Name="WhisperActivityTextBox" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True"
                                         VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                         FontFamily="Consolas" FontSize="12" MinHeight="120" MaxHeight="220"/>
                            </Expander>
                            <Expander Grid.Column="2" x:Name="GemmaActivityExpander" Header="Ollama Activity" IsExpanded="False">
                                <TextBox x:Name="GemmaActivityTextBox" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True"
                                         VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                         FontFamily="Consolas" FontSize="12" MinHeight="120" MaxHeight="220"/>
                            </Expander>
                        </Grid>
                    </Grid>
                </TabItem>

                <TabItem Header="Transcript Timing">
                    <Grid Margin="8">
                        <TextBox x:Name="TranscriptTimingTextBox" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                 FontFamily="Consolas" FontSize="12"/>
                    </Grid>
                </TabItem>
            </TabControl>
        </Grid>

        <TextBlock Grid.Row="3" x:Name="StatusTextBlock" Margin="0,8,0,0" Text="Ready. Enter a YouTube video ID or paste a full URL, then click Fetch."/>
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
$BrowserOptionsExpander = $window.FindName('BrowserOptionsExpander')
$TitleTextBlock = $window.FindName('TitleTextBlock')
$VideoIdTextBlock = $window.FindName('VideoIdTextBlock')
$MetaGrid = $window.FindName('MetaGrid')
$ChaptersGrid = $window.FindName('ChaptersGrid')
$DescriptionTextBox = $window.FindName('DescriptionTextBox')
$TranscriptStatusText = $window.FindName('TranscriptStatusText')
$WhisperTranscriptButton = $window.FindName('WhisperTranscriptButton')
$OpenTranscriptStudioButton = $window.FindName('OpenTranscriptStudioButton')
$WhisperProgressText = $window.FindName('WhisperProgressText')
$WhisperActivityExpander = $window.FindName('WhisperActivityExpander')
$WhisperActivityTextBox = $window.FindName('WhisperActivityTextBox')
$OllamaModelComboBox = $window.FindName('OllamaModelComboBox')
$GemmaPresetComboBox = $window.FindName('GemmaPresetComboBox')
$RunGemmaButton = $window.FindName('RunGemmaButton')
$StopGemmaButton = $window.FindName('StopGemmaButton')
$GemmaPresetDetailsExpander = $window.FindName('GemmaPresetDetailsExpander')
$GemmaPresetDetailsTextBox = $window.FindName('GemmaPresetDetailsTextBox')
$GemmaStatusText = $window.FindName('GemmaStatusText')
$TranscriptTextBox = $window.FindName('TranscriptTextBox')
$TranscriptTimingTextBox = $window.FindName('TranscriptTimingTextBox')
$GemmaActivityExpander = $window.FindName('GemmaActivityExpander')
$GemmaActivityTextBox = $window.FindName('GemmaActivityTextBox')
$GemmaResultTextBox = $window.FindName('GemmaResultTextBox')
$RawJsonTextBox = $window.FindName('RawJsonTextBox')
$StatusTextBlock = $window.FindName('StatusTextBlock')

$script:lastResult = $null
$ProfilePathBox.Text = if ($defaultProfilePath) { $defaultProfilePath } else { '' }
$UseCookiesCheck.IsChecked = $false

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

function Set-WhisperActivityUi {
    param(
        [string]$ProgressText = 'Whisper idle.',
        [string]$ActivityText = ''
    )

    if ([string]::IsNullOrWhiteSpace($ProgressText)) {
        $ProgressText = 'Whisper idle.'
    }

    $WhisperProgressText.Text = $ProgressText
    $nextActivityText = [string]$ActivityText
    if ($script:isTranscribingWhisper -and -not [string]::IsNullOrWhiteSpace($nextActivityText)) {
        $WhisperActivityExpander.IsExpanded = $true
    }
    if ($WhisperActivityTextBox.Text -ne $nextActivityText) {
        $WhisperActivityTextBox.Text = $nextActivityText
        Scroll-TextBoxToBottom -TextBox $WhisperActivityTextBox
    }
}

function Get-GemmaStatusDisplayText {
    param($Result)

    if ($null -eq $Result) {
        return (Get-ReadyGemmaStatusText -Result $null)
    }

    $status = [string]$Result.GemmaStatus
    if ([string]::IsNullOrWhiteSpace($status)) {
        return (Get-ReadyGemmaStatusText -Result $Result)
    }

    return $status
}

function Set-GemmaUi {
    param(
        [string]$StatusText,
        [string]$ResultText,
        [string]$ActivityText = '',
        [switch]$ScrollResultToEnd
    )

    $nextStatus = if ([string]::IsNullOrWhiteSpace($StatusText)) { Get-ReadyGemmaStatusText -Result $script:lastResult } else { $StatusText }
    $GemmaStatusText.Text = $nextStatus

    $nextActivityText = [string]$ActivityText
    if ($script:isProcessingGemma -and -not [string]::IsNullOrWhiteSpace($nextActivityText)) {
        $GemmaActivityExpander.IsExpanded = $true
    }
    if ($GemmaActivityTextBox.Text -ne $nextActivityText) {
        $GemmaActivityTextBox.Text = $nextActivityText
        Scroll-TextBoxToBottom -TextBox $GemmaActivityTextBox
    }

    $nextResult = [string]$ResultText
    if ($GemmaResultTextBox.Text -ne $nextResult) {
        $GemmaResultTextBox.Text = $nextResult
        if ($ScrollResultToEnd) {
            Scroll-TextBoxToBottom -TextBox $GemmaResultTextBox
        }
        else {
            $GemmaResultTextBox.ScrollToHome()
        }
    }
}

function Scroll-TextBoxToBottom {
    param($TextBox)

    if ($null -eq $TextBox) {
        return
    }

    try {
        $TextBox.UpdateLayout()
        $TextBox.CaretIndex = $TextBox.Text.Length
        $TextBox.ScrollToEnd()
    }
    catch {
    }
}

function Normalize-ToolLogText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    return (([string]$Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Get-ToolLogLines {
    param(
        [string]$Path,
        [switch]$IncludeInternalMarkers
    )

    $normalized = Normalize-ToolLogText -Text (Read-TextFile -Path $Path)
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    $lines = @($normalized -split "`n")
    if (-not $IncludeInternalMarkers) {
        $lines = @(
            foreach ($line in $lines) {
                $text = [string]$line
                if (-not $text.StartsWith($script:ollamaChunkPrefix, [System.StringComparison]::Ordinal)) {
                    $text
                }
            }
        )
    }

    return @($lines)
}

function Get-ToolLogTailText {
    param(
        [string]$Path,
        [int]$MaxLines = 40
    )

    $lines = @(Get-ToolLogLines -Path $Path)
    if ($lines.Count -eq 0) {
        return ''
    }

    if ($lines.Count -gt $MaxLines) {
        $lines = $lines[($lines.Count - $MaxLines)..($lines.Count - 1)]
    }

    return (($lines | ForEach-Object { $_.TrimEnd() }) -join [Environment]::NewLine).Trim()
}

function Get-CombinedToolOutput {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath,
        [switch]$IncludeInternalMarkers
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $stdoutLines = @(Get-ToolLogLines -Path $StdOutPath -IncludeInternalMarkers:$IncludeInternalMarkers)
    $stderrLines = @(Get-ToolLogLines -Path $StdErrPath -IncludeInternalMarkers:$IncludeInternalMarkers)
    $stdoutText = (($stdoutLines | ForEach-Object { $_.TrimEnd() }) -join [Environment]::NewLine).Trim()
    $stderrText = (($stderrLines | ForEach-Object { $_.TrimEnd() }) -join [Environment]::NewLine).Trim()

    if (-not [string]::IsNullOrWhiteSpace($stdoutText)) {
        $parts.Add($stdoutText)
    }

    if (-not [string]::IsNullOrWhiteSpace($stderrText)) {
        $parts.Add($stderrText)
    }

    return ($parts -join [Environment]::NewLine).Trim()
}

function Get-ToolActivityDisplayText {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath,
        [int]$MaxLines = 25,
        [string]$StdOutLabel = 'stdout',
        [string]$StdErrLabel = 'stderr'
    )

    $stdoutTail = Get-ToolLogTailText -Path $StdOutPath -MaxLines $MaxLines
    $stderrTail = Get-ToolLogTailText -Path $StdErrPath -MaxLines $MaxLines
    $sections = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($stdoutTail)) {
        $sections.Add(('[{0}]' -f $StdOutLabel))
        $sections.Add($stdoutTail)
    }

    if (-not [string]::IsNullOrWhiteSpace($stderrTail)) {
        if ($sections.Count -gt 0) {
            $sections.Add('')
        }

        $sections.Add(('[{0}]' -f $StdErrLabel))
        $sections.Add($stderrTail)
    }

    return ($sections -join [Environment]::NewLine).Trim()
}

function Get-LastToolLogLine {
    param([string]$Path)

    $lines = @(Get-ToolLogLines -Path $Path)
    if ($lines.Count -eq 0) {
        return ''
    }

    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        $line = ([string]$lines[$i]).Trim()
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            return $line
        }
    }

    return ''
}

function Remove-ToolLogTimestampPrefix {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return ''
    }

    return (([string]$Line -replace '^\[\d{2}:\d{2}:\d{2}\]\s*', '').Trim())
}

function Get-ActivityExcerptText {
    param(
        [string]$Text,
        [int]$MaxLines = 18,
        [int]$MaxCharacters = 1800
    )

    $normalized = Normalize-ToolLogText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $lines = @($normalized -split "`n")
    if ($lines.Count -gt $MaxLines) {
        $lines = $lines[($lines.Count - $MaxLines)..($lines.Count - 1)]
    }

    $excerpt = (($lines | ForEach-Object { $_.TrimEnd() }) -join [Environment]::NewLine).Trim()
    if ($excerpt.Length -gt $MaxCharacters) {
        $excerpt = $excerpt.Substring($excerpt.Length - $MaxCharacters).Trim()
    }

    return $excerpt
}

function Get-WhisperCombinedOutput {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    return (Get-CombinedToolOutput -StdOutPath $StdOutPath -StdErrPath $StdErrPath)
}

function Get-WhisperActivityDisplayText {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    return (Get-ToolActivityDisplayText -StdOutPath $StdOutPath -StdErrPath $StdErrPath -MaxLines 25)
}

function Get-GemmaActivityDisplayText {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    return (Get-ToolActivityDisplayText -StdOutPath $StdOutPath -StdErrPath $StdErrPath -MaxLines 40)
}

function Get-GemmaStreamingPreviewText {
    param([string]$StdOutPath)

    $lines = @(Get-ToolLogLines -Path $StdOutPath -IncludeInternalMarkers)
    if ($lines.Count -eq 0) {
        return ''
    }

    $builder = New-Object System.Text.StringBuilder
    foreach ($line in $lines) {
        $text = [string]$line
        if (-not $text.StartsWith($script:ollamaChunkPrefix, [System.StringComparison]::Ordinal)) {
            continue
        }

        $encoded = $text.Substring($script:ollamaChunkPrefix.Length)
        if ([string]::IsNullOrWhiteSpace($encoded)) {
            continue
        }

        try {
            $bytes = [System.Convert]::FromBase64String($encoded)
            [void]$builder.Append([System.Text.Encoding]::UTF8.GetString($bytes))
        }
        catch {
        }
    }

    return $builder.ToString()
}

function Get-GemmaProgressDisplayText {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath,
        [string]$Model
    )

    $stderrLatest = Remove-ToolLogTimestampPrefix -Line (Get-LastToolLogLine -Path $StdErrPath)
    if (-not [string]::IsNullOrWhiteSpace($stderrLatest)) {
        return $stderrLatest
    }

    $stdoutLatest = Remove-ToolLogTimestampPrefix -Line (Get-LastToolLogLine -Path $StdOutPath)
    if (-not [string]::IsNullOrWhiteSpace($stdoutLatest)) {
        return $stdoutLatest
    }

    if ([string]::IsNullOrWhiteSpace($Model)) {
        return 'Running Ollama in the background...'
    }

    return ("Running Ollama model '{0}' in the background..." -f $Model)
}

function Get-WhisperProgressDisplayText {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    $stdoutText = Normalize-ToolLogText -Text (Read-TextFile -Path $StdOutPath)
    $stderrText = Normalize-ToolLogText -Text (Read-TextFile -Path $StdErrPath)

    $percentMatches = [System.Text.RegularExpressions.Regex]::Matches($stderrText, '(?<percent>\d{1,3})%')
    if ($percentMatches.Count -gt 0) {
        $percent = [int]$percentMatches[$percentMatches.Count - 1].Groups['percent'].Value
        if ($percent -lt 0) { $percent = 0 }
        if ($percent -gt 100) { $percent = 100 }
        return ("Whisper transcription in progress: {0}% complete." -f $percent)
    }

    $frameMatches = [System.Text.RegularExpressions.Regex]::Matches($stderrText, '(?<done>\d+)\s*/\s*(?<total>\d+)\s*frames')
    if ($frameMatches.Count -gt 0) {
        $match = $frameMatches[$frameMatches.Count - 1]
        $doneFrames = [int]$match.Groups['done'].Value
        $totalFrames = [int]$match.Groups['total'].Value
        if ($totalFrames -gt 0) {
            $framePercent = [math]::Round(($doneFrames / $totalFrames) * 100, 0)
            return ("Whisper transcription in progress: {0}% complete ({1}/{2} frames)." -f [int]$framePercent, $doneFrames, $totalFrames)
        }
    }

    if ($stdoutText -match 'Transcript written to') {
        return 'Whisper is finalizing transcript output...'
    }

    if ($stdoutText -match 'Transcribing ') {
        return 'Whisper is transcribing audio...'
    }

    if ($stdoutText -match 'Loading Whisper model') {
        return ("Whisper is loading model '{0}'..." -f $script:guiWhisperModel)
    }

    return 'Preparing Whisper transcription...'
}

function Apply-WhisperResultToUi {
    param($WhisperResult)

    if ($null -eq $script:lastResult -or $null -eq $WhisperResult) {
        return
    }

    $script:lastResult.Transcript = [string]$WhisperResult.Text
    $script:lastResult.TranscriptStatus = [string]$WhisperResult.Status
    $script:lastResult.TranscriptSource = [string]$WhisperResult.Source
    $script:lastResult.TranscriptMode = if ([string]::IsNullOrWhiteSpace([string]$WhisperResult.Text)) { 'none' } else { 'whisper' }
    $script:lastResult.AudioPath = [string]$WhisperResult.AudioPath
    $script:lastResult.TranscriptPath = [string]$WhisperResult.TranscriptPath
    $script:lastResult.TranscriptTimingJson = [string]$WhisperResult.TimingJson
    $script:lastResult.TranscriptTimingPath = [string]$WhisperResult.TimingPath
    if ($WhisperResult.PSObject.Properties.Name -contains 'BundleAudioPath') {
        $script:lastResult.BundleAudioPath = [string]$WhisperResult.BundleAudioPath
    }
    if ($WhisperResult.PSObject.Properties.Name -contains 'BundleTranscriptPath') {
        $script:lastResult.BundleTranscriptPath = [string]$WhisperResult.BundleTranscriptPath
    }
    if ($WhisperResult.PSObject.Properties.Name -contains 'BundleTimingPath') {
        $script:lastResult.BundleTimingPath = [string]$WhisperResult.BundleTimingPath
    }
    Reset-ResultGemmaState -Result $script:lastResult
    Apply-ResultToUi -Result $script:lastResult
}

function Clear-ActiveWhisperRunState {
    if ($script:whisperPollTimer) {
        $script:whisperPollTimer.Stop()
    }

    if ($script:activeWhisperProcess) {
        try {
            $script:activeWhisperProcess.Dispose()
        }
        catch {
        }
    }

    $script:activeWhisperProcess = $null
    $script:activeWhisperStdOutPath = ''
    $script:activeWhisperStdErrPath = ''
    $script:activeWhisperTranscriptPath = ''
    $script:activeWhisperTimingPath = ''
    $script:activeWhisperAudioPath = ''
    $script:activeWhisperVideoId = ''
    $script:activeWhisperOutputSummary = ''
    $script:activeWhisperBundleRunId = ''
}

function Complete-ActiveWhisperRun {
    if ($null -eq $script:activeWhisperProcess) {
        return
    }

    try {
        $script:activeWhisperProcess.Refresh()
    }
    catch {
    }

    if (-not $script:activeWhisperProcess.HasExited) {
        return
    }

    $videoId = [string]$script:activeWhisperVideoId
    $exitCode = $script:activeWhisperProcess.ExitCode
    $activityText = Get-WhisperActivityDisplayText -StdOutPath $script:activeWhisperStdOutPath -StdErrPath $script:activeWhisperStdErrPath
    $outputText = Get-WhisperCombinedOutput -StdOutPath $script:activeWhisperStdOutPath -StdErrPath $script:activeWhisperStdErrPath
    $exitCodeDisplay = if ([string]::IsNullOrWhiteSpace([string]$exitCode)) { 'unavailable' } else { [string]$exitCode }
    $didSucceed = Test-WhisperRunSucceeded -ExitCode $exitCode -TranscriptPath $script:activeWhisperTranscriptPath -OutputText $outputText
    $whisperRunId = if ([string]::IsNullOrWhiteSpace([string]$script:activeWhisperBundleRunId)) { (Get-BundleRunId -Prefix 'whisper') } else { [string]$script:activeWhisperBundleRunId }
    $script:activeWhisperOutputSummary = $outputText
    $script:isTranscribingWhisper = $false
    $FetchButton.IsEnabled = $true
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow

    if ($didSucceed) {
        $whisperResult = Get-WhisperTranscriptResult -TranscriptPath $script:activeWhisperTranscriptPath -TimingPath $script:activeWhisperTimingPath -AudioPath $script:activeWhisperAudioPath -Output $outputText
        $bundleWhisper = Write-BundleWhisperRunArtifacts `
            -VideoId $videoId `
            -RunId $whisperRunId `
            -Mode 'background' `
            -Status 'completed' `
            -AudioPath $script:activeWhisperAudioPath `
            -TranscriptPath $script:activeWhisperTranscriptPath `
            -TimingPath $script:activeWhisperTimingPath `
            -StdOutPath $script:activeWhisperStdOutPath `
            -StdErrPath $script:activeWhisperStdErrPath `
            -ExitCode $exitCode `
            -OutputText $outputText
        $whisperResult | Add-Member -NotePropertyName BundleAudioPath -NotePropertyValue ([string]$bundleWhisper.AudioPath) -Force
        $whisperResult | Add-Member -NotePropertyName BundleTranscriptPath -NotePropertyValue ([string]$bundleWhisper.TranscriptPath) -Force
        $whisperResult | Add-Member -NotePropertyName BundleTimingPath -NotePropertyValue ([string]$bundleWhisper.TimingPath) -Force
        Apply-WhisperResultToUi -WhisperResult $whisperResult
        Set-WhisperActivityUi -ProgressText $(if ([string]::IsNullOrWhiteSpace([string]$whisperResult.Text)) { 'Whisper finished, but no transcript text was produced.' } else { 'Whisper finished. Local transcript is now loaded.' }) -ActivityText $activityText

        if ([string]$script:lastResult.TranscriptMode -eq 'whisper') {
            Set-Status -Message "Loaded local Whisper transcript for video ID: $videoId" -Color 'Green'
        }
        else {
            Set-Status -Message "Whisper completed for video ID: $videoId, but no text was produced." -Color 'DarkOrange'
        }

        $transcriptLength = ([string]$script:lastResult.Transcript).Length
        Write-Log -Level 'INFO' -Message ("Whisper transcription completed. VideoId: {0} | ExitCode: {1} | TranscriptLength: {2}" -f $videoId, $exitCodeDisplay, $transcriptLength)
    }
    else {
        Write-BundleWhisperRunArtifacts `
            -VideoId $videoId `
            -RunId $whisperRunId `
            -Mode 'background' `
            -Status 'failed' `
            -AudioPath $script:activeWhisperAudioPath `
            -TranscriptPath $script:activeWhisperTranscriptPath `
            -TimingPath $script:activeWhisperTimingPath `
            -StdOutPath $script:activeWhisperStdOutPath `
            -StdErrPath $script:activeWhisperStdErrPath `
            -ExitCode $exitCode `
            -OutputText $outputText | Out-Null
        $errorMessage = "Whisper transcription failed (exit code $exitCodeDisplay). See the Whisper Activity panel for details."
        Set-WhisperActivityUi -ProgressText 'Whisper failed.' -ActivityText $activityText
        if ($script:lastResult) {
            $TranscriptStatusText.Text = Get-TranscriptStatusDisplayText -Result $script:lastResult
        }

        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($errorMessage + $logHint) -Color 'Red'
        Write-Log -Level 'ERROR' -Message ("Whisper transcription failed. VideoId: {0} | ExitCode: {1}`n{2}" -f $videoId, $exitCodeDisplay, $outputText)
        [System.Windows.MessageBox]::Show(($errorMessage + $logHint), 'Whisper Failed', 'OK', 'Error') | Out-Null
    }

    Clear-ActiveWhisperRunState
    Update-WhisperButtonState
    Update-GemmaButtonState
}

function Update-ActiveWhisperRunUi {
    if ($null -eq $script:activeWhisperProcess) {
        return
    }

    $activityText = Get-WhisperActivityDisplayText -StdOutPath $script:activeWhisperStdOutPath -StdErrPath $script:activeWhisperStdErrPath
    $progressText = Get-WhisperProgressDisplayText -StdOutPath $script:activeWhisperStdOutPath -StdErrPath $script:activeWhisperStdErrPath
    Set-WhisperActivityUi -ProgressText $progressText -ActivityText $activityText

    try {
        $script:activeWhisperProcess.Refresh()
    }
    catch {
    }

    if ($script:activeWhisperProcess.HasExited) {
        Complete-ActiveWhisperRun
    }
}

function Clear-ActiveGemmaRunState {
    if ($script:gemmaPollTimer) {
        $script:gemmaPollTimer.Stop()
    }

    if ($script:activeGemmaProcess) {
        try {
            $script:activeGemmaProcess.Dispose()
        }
        catch {
        }
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$script:activeGemmaInputPath) -and (Test-Path -LiteralPath $script:activeGemmaInputPath -PathType Leaf)) {
        Remove-Item -LiteralPath $script:activeGemmaInputPath -Force -ErrorAction SilentlyContinue
    }

    $script:activeGemmaProcess = $null
    $script:activeGemmaStdOutPath = ''
    $script:activeGemmaStdErrPath = ''
    $script:activeGemmaResultPath = ''
    $script:activeGemmaInputPath = ''
    $script:activeGemmaVideoId = ''
    $script:activeGemmaPresetId = ''
    $script:activeGemmaModel = ''
    $script:activeGemmaTranscriptHash = ''
    $script:activeGemmaBundleRunId = ''
}

function Apply-GemmaResultToUi {
    param(
        $GemmaPayload,
        [string]$Model,
        [string]$PresetId,
        [string]$Status,
        [string]$ResultPath,
        [string]$ActivityText = ''
    )

    if ($null -eq $script:lastResult -or $null -eq $GemmaPayload) {
        return
    }

    $script:lastResult.GemmaDisplayText = [string]$GemmaPayload.displayText
    $script:lastResult.GemmaActivityText = [string]$ActivityText
    $script:lastResult.GemmaStatus = [string]$Status
    $script:lastResult.GemmaModel = [string]$Model
    $script:lastResult.GemmaPresetId = [string]$PresetId
    $script:lastResult.GemmaSourceTextHash = [string]$GemmaPayload.sourceTextHash
    $script:lastResult.GemmaResultPath = [string]$ResultPath
    Apply-ResultToUi -Result $script:lastResult
}

function Test-GemmaRunSucceeded {
    param(
        [object]$ExitCode,
        [string]$ResultPath,
        [string]$OutputText
    )

    $parsedExitCode = -1
    if ($null -ne $ExitCode -and [int]::TryParse([string]$ExitCode, [ref]$parsedExitCode)) {
        return ($parsedExitCode -eq 0 -and (Test-Path -LiteralPath $ResultPath -PathType Leaf))
    }

    $hasResultFile = Test-Path -LiteralPath $ResultPath -PathType Leaf
    $hasCompletionMarker = ([string]$OutputText) -match 'Result written to'
    return ($hasResultFile -and $hasCompletionMarker)
}

function Complete-ActiveGemmaRun {
    if ($null -eq $script:activeGemmaProcess) {
        return
    }

    try {
        $script:activeGemmaProcess.Refresh()
    }
    catch {
    }

    if (-not $script:activeGemmaProcess.HasExited) {
        return
    }

    $videoId = [string]$script:activeGemmaVideoId
    $presetId = [string]$script:activeGemmaPresetId
    $model = [string]$script:activeGemmaModel
    $exitCode = $script:activeGemmaProcess.ExitCode
    $activityText = Get-GemmaActivityDisplayText -StdOutPath $script:activeGemmaStdOutPath -StdErrPath $script:activeGemmaStdErrPath
    $outputText = Get-CombinedToolOutput -StdOutPath $script:activeGemmaStdOutPath -StdErrPath $script:activeGemmaStdErrPath
    $activityExcerpt = Get-ActivityExcerptText -Text $outputText
    $didSucceed = Test-GemmaRunSucceeded -ExitCode $exitCode -ResultPath $script:activeGemmaResultPath -OutputText $outputText
    $exitCodeDisplay = if ([string]::IsNullOrWhiteSpace([string]$exitCode)) { 'unavailable' } else { [string]$exitCode }
    $llmRunId = if ([string]::IsNullOrWhiteSpace([string]$script:activeGemmaBundleRunId)) { (Get-BundleRunId -Prefix 'llm') } else { [string]$script:activeGemmaBundleRunId }
    $script:isProcessingGemma = $false
    $FetchButton.IsEnabled = $true
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow

    if ($didSucceed) {
        try {
            $gemmaPayload = Read-GemmaResultFile -Path $script:activeGemmaResultPath
            $bundleLlm = Write-BundleLlmRunArtifacts `
                -VideoId $videoId `
                -RunId $llmRunId `
                -Mode 'background' `
                -Status 'completed' `
                -Model $model `
                -PresetId $presetId `
                -TranscriptHash $script:activeGemmaTranscriptHash `
                -InputFilePath $script:activeGemmaInputPath `
                -ResultFilePath $script:activeGemmaResultPath `
                -StdOutPath $script:activeGemmaStdOutPath `
                -StdErrPath $script:activeGemmaStdErrPath `
                -ActivityText $activityText `
                -ExitCode $exitCode
            $gemmaStatus = "Loaded Ollama result from '$model' for preset '$presetId'."
            Apply-GemmaResultToUi -GemmaPayload $gemmaPayload -Model $model -PresetId $presetId -Status $gemmaStatus -ResultPath $script:activeGemmaResultPath -ActivityText $activityText
            if ($script:lastResult) {
                $script:lastResult.BundleLlmResultPath = [string]$bundleLlm.ResultPath
                $script:lastResult.BundlePromptPath = [string]$bundleLlm.PromptPath
                $script:lastResult.BundleDisplayPath = [string]$bundleLlm.DisplayPath
            }
            Set-Status -Message "Loaded Ollama result for video ID: $videoId" -Color 'Green'
            Write-Log -Level 'INFO' -Message ("Ollama processing completed. VideoId: {0} | Preset: {1} | Model: {2} | ExitCode: {3}" -f $videoId, $presetId, $model, $exitCodeDisplay)
        }
        catch {
            Write-BundleLlmRunArtifacts `
                -VideoId $videoId `
                -RunId $llmRunId `
                -Mode 'background' `
                -Status 'failed' `
                -Model $model `
                -PresetId $presetId `
                -TranscriptHash $script:activeGemmaTranscriptHash `
                -InputFilePath $script:activeGemmaInputPath `
                -ResultFilePath $script:activeGemmaResultPath `
                -StdOutPath $script:activeGemmaStdOutPath `
                -StdErrPath $script:activeGemmaStdErrPath `
                -ActivityText $activityText `
                -ExitCode $exitCode | Out-Null
            $errorMessage = "Ollama completed, but the result file could not be loaded. See the Ollama Activity panel for details."
            if ($script:lastResult) {
                $script:lastResult.GemmaActivityText = [string]$activityText
                $script:lastResult.GemmaStatus = $errorMessage
                $script:lastResult.GemmaModel = $model
                Apply-ResultToUi -Result $script:lastResult
            }

            $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
            Set-Status -Message ($errorMessage + $logHint) -Color 'Red'
            Write-Log -Level 'ERROR' -Message ("Ollama result loading failed. VideoId: {0} | Preset: {1} | Model: {2}`n{3}`n{4}" -f $videoId, $presetId, $model, $activityText, (Get-ErrorDetails -ErrorObject $_))
            $dialogMessage = $errorMessage
            if (-not [string]::IsNullOrWhiteSpace($activityExcerpt)) {
                $dialogMessage = $dialogMessage + "`n`nRecent helper output:`n" + $activityExcerpt
            }
            [System.Windows.MessageBox]::Show(($dialogMessage + $logHint), 'Ollama Failed', 'OK', 'Error') | Out-Null
        }
    }
    else {
        Write-BundleLlmRunArtifacts `
            -VideoId $videoId `
            -RunId $llmRunId `
            -Mode 'background' `
            -Status 'failed' `
            -Model $model `
            -PresetId $presetId `
            -TranscriptHash $script:activeGemmaTranscriptHash `
            -InputFilePath $script:activeGemmaInputPath `
            -ResultFilePath $script:activeGemmaResultPath `
            -StdOutPath $script:activeGemmaStdOutPath `
            -StdErrPath $script:activeGemmaStdErrPath `
            -ActivityText $activityText `
            -ExitCode $exitCode | Out-Null
        $errorMessage = "Ollama processing failed (exit code $exitCodeDisplay)."
        $statusDetail = Remove-ToolLogTimestampPrefix -Line (Get-GemmaProgressDisplayText -StdOutPath $script:activeGemmaStdOutPath -StdErrPath $script:activeGemmaStdErrPath -Model $model)
        if ($script:lastResult) {
            $script:lastResult.GemmaActivityText = [string]$activityText
            $script:lastResult.GemmaStatus = if ([string]::IsNullOrWhiteSpace($statusDetail)) { 'Ollama failed.' } else { $statusDetail }
            $script:lastResult.GemmaModel = $model
            Apply-ResultToUi -Result $script:lastResult
        }

        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($errorMessage + $logHint) -Color 'Red'
        Write-Log -Level 'ERROR' -Message ("Ollama processing failed. VideoId: {0} | Preset: {1} | Model: {2} | ExitCode: {3}`n{4}" -f $videoId, $presetId, $model, $exitCodeDisplay, $outputText)
        $dialogMessage = $errorMessage
        if (-not [string]::IsNullOrWhiteSpace($activityExcerpt)) {
            $dialogMessage = $dialogMessage + "`n`nRecent helper output:`n" + $activityExcerpt
        }
        [System.Windows.MessageBox]::Show(($dialogMessage + $logHint), 'Ollama Failed', 'OK', 'Error') | Out-Null
    }

    Clear-ActiveGemmaRunState
    Update-WhisperButtonState
    Update-GemmaButtonState
}

function Update-ActiveGemmaRunUi {
    if ($null -eq $script:activeGemmaProcess) {
        return
    }

    $activityText = Get-GemmaActivityDisplayText -StdOutPath $script:activeGemmaStdOutPath -StdErrPath $script:activeGemmaStdErrPath
    $progressText = Get-GemmaProgressDisplayText -StdOutPath $script:activeGemmaStdOutPath -StdErrPath $script:activeGemmaStdErrPath -Model $script:activeGemmaModel
    $streamingPreview = Get-GemmaStreamingPreviewText -StdOutPath $script:activeGemmaStdOutPath
    if ($script:lastResult) {
        $script:lastResult.GemmaActivityText = [string]$activityText
        $script:lastResult.GemmaStatus = [string]$progressText
        $script:lastResult.GemmaModel = [string]$script:activeGemmaModel
        if (-not [string]::IsNullOrWhiteSpace($streamingPreview)) {
            $script:lastResult.GemmaDisplayText = [string]$streamingPreview
            Set-GemmaUi -StatusText $progressText -ResultText ([string]$streamingPreview) -ActivityText $activityText -ScrollResultToEnd
        }
        else {
            Set-GemmaUi -StatusText $progressText -ResultText ([string]$script:lastResult.GemmaDisplayText) -ActivityText $activityText
        }
    }

    try {
        $script:activeGemmaProcess.Refresh()
    }
    catch {
    }

    if ($script:activeGemmaProcess.HasExited) {
        Complete-ActiveGemmaRun
    }
}

function Stop-ActiveGemmaRun {
    param(
        [string]$Reason = 'Ollama run stopped by user.'
    )

    if ($null -eq $script:activeGemmaProcess) {
        return
    }

    $model = [string]$script:activeGemmaModel
    $presetId = [string]$script:activeGemmaPresetId
    $activityText = Get-GemmaActivityDisplayText -StdOutPath $script:activeGemmaStdOutPath -StdErrPath $script:activeGemmaStdErrPath
    $streamingPreview = Get-GemmaStreamingPreviewText -StdOutPath $script:activeGemmaStdOutPath
    $videoId = if ($script:lastResult) { [string]$script:lastResult.VideoId } else { [string]$script:activeGemmaVideoId }
    $runId = if ([string]::IsNullOrWhiteSpace([string]$script:activeGemmaBundleRunId)) { (Get-BundleRunId -Prefix 'llm') } else { [string]$script:activeGemmaBundleRunId }
    $transcriptHash = [string]$script:activeGemmaTranscriptHash
    $inputPath = [string]$script:activeGemmaInputPath
    $resultPath = [string]$script:activeGemmaResultPath
    $stdoutPath = [string]$script:activeGemmaStdOutPath
    $stderrPath = [string]$script:activeGemmaStdErrPath

    try {
        $script:activeGemmaProcess.Refresh()
        if ($script:activeGemmaProcess.HasExited) {
            Complete-ActiveGemmaRun
            return
        }

        Write-Log -Level 'WARN' -Message ("Stopping active Ollama helper. Preset: {0} | Model: {1} | PID: {2}" -f $presetId, $model, $script:activeGemmaProcess.Id)
        $script:activeGemmaProcess.Kill()
    }
    catch {
        Write-Log -Level 'WARN' -Message ("Failed to stop the active Ollama helper:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
    }
    finally {
        Clear-ActiveGemmaRunState
    }

    $script:isProcessingGemma = $false
    Write-BundleLlmRunArtifacts `
        -VideoId $videoId `
        -RunId $runId `
        -Mode 'background' `
        -Status 'failed' `
        -Model $model `
        -PresetId $presetId `
        -TranscriptHash $transcriptHash `
        -InputFilePath $inputPath `
        -TranscriptText $(if ($script:lastResult) { [string]$script:lastResult.Transcript } else { '' }) `
        -ResultFilePath $resultPath `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath `
        -ActivityText $Reason | Out-Null
    $FetchButton.IsEnabled = $true
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    if ($script:lastResult) {
        $script:lastResult.GemmaActivityText = [string]$activityText
        if (-not [string]::IsNullOrWhiteSpace($streamingPreview)) {
            $script:lastResult.GemmaDisplayText = [string]$streamingPreview
        }
        $script:lastResult.GemmaStatus = [string]$Reason
        $script:lastResult.GemmaModel = [string]$model
        $script:lastResult.GemmaPresetId = [string]$presetId
        Apply-ResultToUi -Result $script:lastResult
    }

    Set-Status -Message $Reason -Color 'DarkOrange'
    Update-WhisperButtonState
    Update-GemmaButtonState
}

function Test-WhisperRunSucceeded {
    param(
        [object]$ExitCode,
        [string]$TranscriptPath,
        [string]$OutputText
    )

    $parsedExitCode = -1
    if ($null -ne $ExitCode -and [int]::TryParse([string]$ExitCode, [ref]$parsedExitCode)) {
        return ($parsedExitCode -eq 0)
    }

    $transcriptExists = Test-Path -LiteralPath $TranscriptPath -PathType Leaf
    $hasCompletionMarker = ([string]$OutputText) -match 'Transcript written to'
    if ($transcriptExists -and $hasCompletionMarker) {
        Write-Log -Level 'WARN' -Message ("Whisper helper exit code was unavailable; treating the run as successful because transcript output exists at '{0}'." -f $TranscriptPath)
        return $true
    }

    return $false
}

function Get-TranscriptStatusDisplayText {
    param($Result)

    if ($null -eq $Result) {
        return ''
    }

    $status = [string]$Result.TranscriptStatus
    if ([string]::IsNullOrWhiteSpace($status)) {
        $status = 'No transcript loaded.'
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Result.TranscriptSource)) {
        $status = "$status Source: $($Result.TranscriptSource)"
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Result.TranscriptTimingPath)) {
        $status = "$status Whisper timing metadata is available."
    }

    if ([string]$Result.TranscriptMode -ne 'whisper' -and -not [string]::IsNullOrWhiteSpace([string]$Result.VideoId)) {
        if (Test-WhisperTranscriptCached -VideoId ([string]$Result.VideoId)) {
            $status = "$status A cached Whisper transcript is available on demand."
        }
        else {
            $status = "$status Use the Whisper button to generate a local transcript."
        }
    }

    return $status
}

function Update-WhisperButtonState {
    $isEnabled = $false
    $content = 'Transcribe with Whisper'

    if ($script:isTranscribingWhisper) {
        $content = 'Whisper Running...'
    }
    elseif (-not $script:isFetching -and -not $script:isProcessingGemma -and $null -ne $script:lastResult -and -not [string]::IsNullOrWhiteSpace([string]$script:lastResult.VideoId)) {
        if (Test-WhisperTranscriptCached -VideoId ([string]$script:lastResult.VideoId)) {
            $content = 'Load Whisper Transcript'
        }
        $isEnabled = $true
    }

    $WhisperTranscriptButton.Content = $content
    $WhisperTranscriptButton.IsEnabled = $isEnabled
    Update-TranscriptStudioButtonState
}

function Update-TranscriptStudioButtonState {
    $OpenTranscriptStudioButton.IsEnabled = $false

    if ($script:isFetching -or $script:isTranscribingWhisper -or $script:isProcessingGemma) {
        return
    }

    if ($null -eq $script:lastResult) {
        return
    }

    if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.Transcript)) {
        return
    }

    $audioPath = [string]$script:lastResult.AudioPath
    if ([string]::IsNullOrWhiteSpace($audioPath)) {
        return
    }

    if (-not (Test-Path -LiteralPath $audioPath -PathType Leaf)) {
        return
    }

    $OpenTranscriptStudioButton.IsEnabled = $true
}

function Update-GemmaButtonState {
    $runEnabled = $false
    $runContent = 'Run with Ollama'
    $stopEnabled = $false
    $modelEnabled = $false
    $presetEnabled = $false

    if ($script:gemmaHelperAvailable) {
        if ($script:isProcessingGemma) {
            $runContent = 'Ollama Running...'
            $stopEnabled = $true
        }
        elseif (-not $script:isFetching -and -not $script:isTranscribingWhisper) {
            $modelEnabled = ($script:ollamaModels.Count -gt 0)

            if ($null -ne $script:lastResult -and -not [string]::IsNullOrWhiteSpace([string]$script:lastResult.Transcript)) {
                $selectedModel = Get-SelectedOllamaModel
                if (-not [string]::IsNullOrWhiteSpace($selectedModel)) {
                    $presetEnabled = ($script:gemmaPresets.Count -gt 0)
                    $selectedPresetId = Get-SelectedGemmaPresetId
                    if (-not [string]::IsNullOrWhiteSpace($selectedPresetId)) {
                        $runEnabled = $true
                    }
                }
            }
        }
    }

    $RunGemmaButton.IsEnabled = $runEnabled
    $RunGemmaButton.Content = $runContent
    $StopGemmaButton.IsEnabled = $stopEnabled
    $OllamaModelComboBox.IsEnabled = $modelEnabled
    $GemmaPresetComboBox.IsEnabled = $presetEnabled
    Update-TranscriptStudioButtonState
}

function Find-OllamaModelSelectionIndex {
    param(
        [object[]]$Models,
        [string]$PreferredModel
    )

    if ($null -eq $Models -or $Models.Count -eq 0) {
        return -1
    }

    $target = ([string]$PreferredModel).Trim()
    if ([string]::IsNullOrWhiteSpace($target)) {
        return 0
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add($target)
    if ($target -notmatch ':') {
        $candidates.Add(($target + ':latest'))
    }
    elseif ($target.EndsWith(':latest', [System.StringComparison]::OrdinalIgnoreCase)) {
        $candidates.Add(($target -replace ':latest$', ''))
    }

    foreach ($candidate in $candidates) {
        for ($i = 0; $i -lt $Models.Count; $i++) {
            if ([string]$Models[$i].model -eq $candidate) {
                return $i
            }
        }
    }

    return -1
}

function Initialize-OllamaModels {
    $script:ollamaModels = @()
    $OllamaModelComboBox.ItemsSource = $null

    $statusOverride = ''
    try {
        $loadedModels = New-Object System.Collections.Generic.List[object]
        foreach ($rawModel in @(Get-OllamaModelDefinitions)) {
            $entry = New-OllamaModelEntry `
                -Model ([string]$rawModel.model) `
                -Display ([string]$rawModel.display) `
                -Family ([string]$rawModel.family) `
                -ParameterSize ([string]$rawModel.parameterSize) `
                -QuantizationLevel ([string]$rawModel.quantizationLevel) `
                -Source 'ollama'
            if ($null -ne $entry) {
                $loadedModels.Add($entry)
            }
        }

        if ($loadedModels.Count -eq 0) {
            throw 'Ollama returned no installed models.'
        }

        $script:ollamaModels = @($loadedModels.ToArray())
        Write-Log -Level 'INFO' -Message ("Loaded {0} installed Ollama models." -f $script:ollamaModels.Count)
    }
    catch {
        $fallbackModel = New-OllamaModelEntry -Model $script:ollamaModel -Display ("{0} (configured default)" -f $script:ollamaModel) -Source 'config'
        if ($null -ne $fallbackModel) {
            $script:ollamaModels = @($fallbackModel)
        }
        $statusOverride = ("Could not query installed Ollama models. Using configured default '{0}'." -f $script:ollamaModel)
        Write-Log -Level 'WARN' -Message ("Ollama model initialization failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
    }

    $OllamaModelComboBox.DisplayMemberPath = 'display'
    $OllamaModelComboBox.SelectedValuePath = 'model'
    $OllamaModelComboBox.ItemsSource = $script:ollamaModels
    if ($script:ollamaModels.Count -gt 0) {
        $selectedIndex = Find-OllamaModelSelectionIndex -Models $script:ollamaModels -PreferredModel $script:ollamaModel
        if ($selectedIndex -lt 0) {
            $selectedIndex = 0
        }
        $OllamaModelComboBox.SelectedIndex = $selectedIndex
    }

    if (-not [string]::IsNullOrWhiteSpace($statusOverride)) {
        $existingResultText = ''
        $existingActivityText = ''
        if ($null -ne $script:lastResult) {
            $script:lastResult.GemmaStatus = $statusOverride
            $existingResultText = [string]$script:lastResult.GemmaDisplayText
            $existingActivityText = [string]$script:lastResult.GemmaActivityText
        }
        Set-GemmaUi -StatusText $statusOverride -ResultText $existingResultText -ActivityText $existingActivityText
    }
    elseif ($script:gemmaHelperAvailable) {
        $existingResultText = if ($script:lastResult) { [string]$script:lastResult.GemmaDisplayText } else { '' }
        $existingActivityText = if ($script:lastResult) { [string]$script:lastResult.GemmaActivityText } else { '' }
        Set-GemmaUi -StatusText (Get-ReadyGemmaStatusText -Result $script:lastResult) -ResultText $existingResultText -ActivityText $existingActivityText
    }

    Update-GemmaButtonState
}

function Initialize-GemmaPresets {
    $script:gemmaHelperAvailable = $false
    $script:gemmaPresets = @()
    $GemmaPresetComboBox.ItemsSource = $null
    $GemmaPresetDetailsTextBox.Text = 'Select a prompt preset to inspect the exact template instructions that will be sent with the transcript.'

    try {
        $presets = @(Get-GemmaPresetDefinitions)
        $script:gemmaPresets = $presets
        $script:gemmaHelperAvailable = ($presets.Count -gt 0)
        $GemmaPresetComboBox.DisplayMemberPath = 'label'
        $GemmaPresetComboBox.SelectedValuePath = 'id'
        $GemmaPresetComboBox.ItemsSource = $presets
        if ($presets.Count -gt 0) {
            $GemmaPresetComboBox.SelectedIndex = 0
            Update-GemmaPresetTemplateUi
            Write-Log -Level 'INFO' -Message ("Loaded {0} Gemma preset definitions." -f $presets.Count)
        }
        else {
            $GemmaStatusText.Text = 'No Ollama prompt presets were found.'
            Write-Log -Level 'WARN' -Message 'Gemma preset listing returned no presets.'
        }
    }
    catch {
        $script:gemmaHelperAvailable = $false
        $script:gemmaPresets = @()
        $GemmaPresetComboBox.ItemsSource = $null
        $GemmaStatusText.Text = 'The Ollama helper is unavailable. See the status bar or logs for details.'
        Write-Log -Level 'WARN' -Message ("Gemma preset initialization failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
    }

    if ($script:gemmaHelperAvailable) {
        Set-GemmaUi -StatusText (Get-ReadyGemmaStatusText -Result $script:lastResult) -ResultText '' -ActivityText ''
    }

    Update-GemmaButtonState
}

function Apply-ResultToUi {
    param($Result)

    if ($null -eq $Result) {
        $TitleTextBlock.Text = 'No video loaded.'
        $VideoIdTextBlock.Text = 'Fetch a video to load metadata.'
        $MetaGrid.ItemsSource = $null
        $ChaptersGrid.ItemsSource = $null
        $DescriptionTextBox.Text = ''
        $TranscriptTextBox.Text = ''
        $TranscriptTimingTextBox.Text = ''
        $TranscriptStatusText.Text = ''
        Set-GemmaUi -StatusText (Get-ReadyGemmaStatusText -Result $null) -ResultText '' -ActivityText ''
        Set-WhisperActivityUi -ProgressText 'Whisper idle.' -ActivityText ''
        $RawJsonTextBox.Text = ''
        Update-WhisperButtonState
        Update-GemmaButtonState
        return
    }

    $TitleTextBlock.Text = if ([string]::IsNullOrWhiteSpace([string]$Result.Title)) { 'Untitled video' } else { [string]$Result.Title }
    $VideoIdTextBlock.Text = if ([string]::IsNullOrWhiteSpace([string]$Result.VideoId)) {
        ([string]$Result.TranscriptStatus)
    }
    else {
        ("Video ID: {0}" -f [string]$Result.VideoId)
    }
    $MetaGrid.ItemsSource = $Result.MetaRows
    $ChaptersGrid.ItemsSource = $Result.Chapters
    $DescriptionTextBox.Text = [string]$Result.Description
    $TranscriptTextBox.Text = [string]$Result.Transcript
    $TranscriptTimingTextBox.Text = if ([string]::IsNullOrWhiteSpace([string]$Result.TranscriptTimingJson)) { 'No local Whisper timing metadata loaded.' } else { [string]$Result.TranscriptTimingJson }
    $TranscriptStatusText.Text = Get-TranscriptStatusDisplayText -Result $Result
    Set-GemmaUi -StatusText (Get-GemmaStatusDisplayText -Result $Result) -ResultText ([string]$Result.GemmaDisplayText) -ActivityText ([string]$Result.GemmaActivityText)
    $RawJsonTextBox.Text = [string]$Result.RawJson
    Update-WhisperButtonState
    Update-GemmaButtonState
}

$script:isFetching = $false
$script:isTranscribingWhisper = $false
$script:isProcessingGemma = $false
$script:whisperPollTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:whisperPollTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:whisperPollTimer.Add_Tick({
    Update-ActiveWhisperRunUi
})
$script:gemmaPollTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:gemmaPollTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:gemmaPollTimer.Add_Tick({
    Update-ActiveGemmaRunUi
})
Set-WhisperActivityUi -ProgressText 'Whisper idle.' -ActivityText ''
Set-GemmaUi -StatusText (Get-ReadyGemmaStatusText -Result $null) -ResultText '' -ActivityText ''
Update-WhisperButtonState
Update-GemmaButtonState
Initialize-GemmaPresets
Initialize-OllamaModels

$FetchButton.Add_Click({
    if ($script:isFetching) {
        return
    }

    try {
        $url = Resolve-YoutubeInputToUrl -InputText $UrlBox.Text
        Assert-ValidYoutubeUrl -Url $url

        $useCookies = [bool]$UseCookiesCheck.IsChecked
        $profilePath = $ProfilePathBox.Text.Trim()
        Write-Log -Level 'INFO' -Message ("Fetch clicked. URL: {0} | UseCookies: {1} | ProfilePath: {2}" -f $url, $useCookies, $(if ($profilePath) { $profilePath } else { '<empty>' }))

        $script:isFetching = $true
        $FetchButton.IsEnabled = $false
        $WhisperTranscriptButton.IsEnabled = $false
        $OpenTranscriptStudioButton.IsEnabled = $false
        $RunGemmaButton.IsEnabled = $false
        $OllamaModelComboBox.IsEnabled = $false
        $GemmaPresetComboBox.IsEnabled = $false
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Set-WhisperActivityUi -ProgressText 'Whisper idle.' -ActivityText ''
        Set-Status -Message 'Fetching metadata and transcript...' -Color 'DarkBlue'
        Write-Log -Level 'INFO' -Message "Fetch execution started on UI thread."

        $result = Get-VideoResearchData -ExePath $YtDlpPath -Url $url -UseCookies $useCookies -ProfilePath $profilePath
        $script:lastResult = $result
        Apply-ResultToUi -Result $script:lastResult

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
        Update-WhisperButtonState
        Update-GemmaButtonState
        Write-Log -Level 'DEBUG' -Message 'Fetch attempt finished.'
    }
})

$WhisperTranscriptButton.Add_Click({
    if ($script:isTranscribingWhisper -or $null -eq $script:lastResult) {
        return
    }

    if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.VideoId)) {
        return
    }

    try {
        $videoId = [string]$script:lastResult.VideoId
        $url = if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.VideoUrl)) {
            Resolve-YoutubeInputToUrl -InputText $UrlBox.Text
        }
        else {
            [string]$script:lastResult.VideoUrl
        }
        Assert-ValidYoutubeUrl -Url $url

        $useCookies = [bool]$UseCookiesCheck.IsChecked
        $profilePath = $ProfilePathBox.Text.Trim()
        $cookieArgs = @(Get-CookieArgs -UseCookies $useCookies -ProfilePath $profilePath)
        Write-Log -Level 'INFO' -Message ("Whisper transcription requested. VideoId: {0} | URL: {1} | UseCookies: {2}" -f $videoId, $url, $useCookies)

        $script:isTranscribingWhisper = $true
        $FetchButton.IsEnabled = $false
        $WhisperTranscriptButton.IsEnabled = $false
        $OpenTranscriptStudioButton.IsEnabled = $false
        $RunGemmaButton.IsEnabled = $false
        $OllamaModelComboBox.IsEnabled = $false
        $GemmaPresetComboBox.IsEnabled = $false
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $TranscriptStatusText.Text = 'Preparing local Whisper transcription. The transcript text will update when it finishes.'
        Set-WhisperActivityUi -ProgressText 'Preparing audio for Whisper...' -ActivityText 'Checking cached files and extracting audio if needed...'
        Set-Status -Message 'Preparing local Whisper transcription...' -Color 'DarkBlue'

        $whisperStart = Start-WhisperTranscription -ExePath $YtDlpPath -Url $url -VideoId $videoId -CookieArgs $cookieArgs
        if ([string]$whisperStart.Mode -eq 'cached') {
            $script:isTranscribingWhisper = $false
            $FetchButton.IsEnabled = $true
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $cachedWhisperRunId = Get-BundleRunId -Prefix 'whisper'
            $bundleWhisper = Write-BundleWhisperRunArtifacts `
                -VideoId $videoId `
                -RunId $cachedWhisperRunId `
                -Mode 'cached' `
                -Status 'completed' `
                -AudioPath ([string]$whisperStart.Result.AudioPath) `
                -TranscriptPath ([string]$whisperStart.Result.TranscriptPath) `
                -TimingPath ([string]$whisperStart.Result.TimingPath) `
                -OutputText 'Loaded cached Whisper transcript.'
            $whisperStart.Result | Add-Member -NotePropertyName BundleAudioPath -NotePropertyValue ([string]$bundleWhisper.AudioPath) -Force
            $whisperStart.Result | Add-Member -NotePropertyName BundleTranscriptPath -NotePropertyValue ([string]$bundleWhisper.TranscriptPath) -Force
            $whisperStart.Result | Add-Member -NotePropertyName BundleTimingPath -NotePropertyValue ([string]$bundleWhisper.TimingPath) -Force
            Apply-WhisperResultToUi -WhisperResult $whisperStart.Result
            Set-WhisperActivityUi -ProgressText 'Loaded cached Whisper transcript.' -ActivityText ("Loaded cached transcript from '{0}'." -f $whisperStart.Result.TranscriptPath)
            Set-Status -Message "Loaded cached local Whisper transcript for video ID: $videoId" -Color 'Green'
            Write-Log -Level 'INFO' -Message ("Loaded cached Whisper transcript without starting a new process. VideoId: {0}" -f $videoId)
            Update-WhisperButtonState
            Update-GemmaButtonState
            return
        }

        $script:activeWhisperProcess = $whisperStart.Process
        $script:activeWhisperStdOutPath = [string]$whisperStart.StdOutPath
        $script:activeWhisperStdErrPath = [string]$whisperStart.StdErrPath
        $script:activeWhisperTranscriptPath = [string]$whisperStart.TranscriptPath
        $script:activeWhisperTimingPath = [string]$whisperStart.TimingPath
        $script:activeWhisperAudioPath = [string]$whisperStart.AudioPath
        $script:activeWhisperVideoId = $videoId
        $script:activeWhisperBundleRunId = Get-BundleRunId -Prefix 'whisper'
        $bundleAudioForStart = Ensure-BundleAudioArtifact -VideoId $videoId -AudioPath ([string]$whisperStart.AudioPath)
        Add-BundleEvent -VideoId $videoId -Type 'whisper_started' -RunId $script:activeWhisperBundleRunId -Status 'running' -Inputs ([ordered]@{
            audio = if ([string]::IsNullOrWhiteSpace($bundleAudioForStart)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $videoId) -TargetPath $bundleAudioForStart) }
        }) -Details ([ordered]@{
            model    = $script:guiWhisperModel
            language = $script:guiWhisperLanguage
        })
        $TranscriptStatusText.Text = 'Running local Whisper transcription in background. The transcript text will update when it finishes.'
        Set-Status -Message 'Running local Whisper transcription in background...' -Color 'DarkBlue'
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        Update-WhisperButtonState
        Update-GemmaButtonState
        Update-ActiveWhisperRunUi
        $script:whisperPollTimer.Start()
        Write-Log -Level 'INFO' -Message ("Whisper process started. VideoId: {0} | PID: {1}" -f $videoId, $script:activeWhisperProcess.Id)
    }
    catch {
        Write-Log -Level 'ERROR' -Message ("Whisper transcription failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
        if ($script:activeWhisperProcess) {
            try {
                $script:activeWhisperProcess.Refresh()
                if (-not $script:activeWhisperProcess.HasExited) {
                    $script:activeWhisperProcess.Kill()
                }
            }
            catch {
            }
        }
        Clear-ActiveWhisperRunState
        $script:isTranscribingWhisper = $false
        $FetchButton.IsEnabled = $true
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($_.Exception.Message + $logHint) -Color 'Red'
        [System.Windows.MessageBox]::Show(($_.Exception.Message + $logHint), 'Whisper Failed', 'OK', 'Error') | Out-Null
        if ($script:lastResult) {
            $TranscriptStatusText.Text = Get-TranscriptStatusDisplayText -Result $script:lastResult
        }
        Update-WhisperButtonState
        Update-GemmaButtonState
    }
})

$OpenTranscriptStudioButton.Add_Click({
    if ($null -eq $script:lastResult) {
        return
    }

    $sessionPath = ''
    $launchResult = $null
    $videoId = [string]$script:lastResult.VideoId
    try {
        if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.Transcript)) {
            throw 'Transcript Studio requires transcript text.'
        }

        $audioPath = [string]$script:lastResult.AudioPath
        if ([string]::IsNullOrWhiteSpace($audioPath) -or -not (Test-Path -LiteralPath $audioPath -PathType Leaf)) {
            throw 'Transcript Studio requires a local audio file. Run Whisper first so the GUI has extracted audio to hand off.'
        }

        $sessionPath = Write-TranscriptStudioSessionFile -Result $script:lastResult
        $launchResult = Start-TranscriptStudioApp -SessionPath $sessionPath
        if (-not [string]::IsNullOrWhiteSpace($videoId)) {
            Add-BundleEvent -VideoId $videoId -Type 'transcript_studio_opened' -Outputs ([ordered]@{
                session = if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.BundleSessionPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $videoId) -TargetPath ([string]$script:lastResult.BundleSessionPath)) }
            }) -Details ([ordered]@{
                processId = $launchResult.Process.Id
                stdoutLog = if ([string]::IsNullOrWhiteSpace([string]$launchResult.StdOutPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $videoId) -TargetPath ([string]$launchResult.StdOutPath)) }
                stderrLog = if ([string]::IsNullOrWhiteSpace([string]$launchResult.StdErrPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $videoId) -TargetPath ([string]$launchResult.StdErrPath)) }
            })
        }
        $targetLabel = if ([string]::IsNullOrWhiteSpace($videoId)) { 'current session' } else { "video ID: $videoId" }
        Set-Status -Message "Opened Transcript Studio for $targetLabel" -Color 'DarkGreen'
        Write-Log -Level 'INFO' -Message ("Transcript Studio launched. VideoId: {0} | SessionPath: {1} | PID: {2} | stdout: {3} | stderr: {4}" -f $videoId, $sessionPath, $launchResult.Process.Id, [string]$launchResult.StdOutPath, [string]$launchResult.StdErrPath)
    }
    catch {
        Write-Log -Level 'ERROR' -Message ("Transcript Studio launch failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
        if (-not [string]::IsNullOrWhiteSpace($videoId)) {
            Add-BundleEvent -VideoId $videoId -Type 'transcript_studio_launch_failed' -Status 'failed' -Outputs ([ordered]@{
                session = if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.BundleSessionPath)) { '' } elseif (Test-Path -LiteralPath ([string]$script:lastResult.BundleSessionPath) -PathType Leaf) { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $videoId) -TargetPath ([string]$script:lastResult.BundleSessionPath)) } else { '' }
            }) -Details ([ordered]@{
                sessionPath = [string]$sessionPath
                stdoutLog   = if ($null -eq $launchResult) { '' } else { [string]$launchResult.StdOutPath }
                stderrLog   = if ($null -eq $launchResult) { '' } else { [string]$launchResult.StdErrPath }
                error       = [string]$_.Exception.Message
            })
        }
        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($_.Exception.Message + $logHint) -Color 'Red'
        [System.Windows.MessageBox]::Show(($_.Exception.Message + $logHint), 'Transcript Studio Failed', 'OK', 'Error') | Out-Null
    }
})

$GemmaPresetComboBox.Add_SelectionChanged({
    Update-GemmaPresetTemplateUi
    Update-GemmaButtonState

    if ($script:isProcessingGemma -or $null -eq $script:lastResult) {
        return
    }

    $selectedPresetId = Get-SelectedGemmaPresetId
    if ([string]::IsNullOrWhiteSpace($selectedPresetId)) {
        return
    }

    if ([string]$script:lastResult.GemmaPresetId -ne $selectedPresetId) {
        Reset-ResultGemmaState -Result $script:lastResult
        Apply-ResultToUi -Result $script:lastResult
    }
})

$OllamaModelComboBox.Add_SelectionChanged({
    Update-GemmaButtonState

    if ($script:isProcessingGemma -or $null -eq $script:lastResult) {
        return
    }

    $selectedModel = Get-SelectedOllamaModel
    if ([string]::IsNullOrWhiteSpace($selectedModel)) {
        return
    }

    if ([string]$script:lastResult.GemmaModel -ne $selectedModel) {
        Reset-ResultGemmaState -Result $script:lastResult
        Apply-ResultToUi -Result $script:lastResult
    }
})

$RunGemmaButton.Add_Click({
    if ($script:isProcessingGemma -or $null -eq $script:lastResult) {
        return
    }

    $transcriptText = [string]$script:lastResult.Transcript
    if ([string]::IsNullOrWhiteSpace($transcriptText)) {
        return
    }

    $presetId = Get-SelectedGemmaPresetId
    if ([string]::IsNullOrWhiteSpace($presetId)) {
        return
    }

    $selectedModel = Get-SelectedOllamaModel
    if ([string]::IsNullOrWhiteSpace($selectedModel)) {
        return
    }

    try {
        $videoId = [string]$script:lastResult.VideoId
        $transcriptHash = Get-TextSha256 -Text $transcriptText
        if ([string]::IsNullOrWhiteSpace($transcriptHash)) {
            throw 'The current transcript is empty after normalization.'
        }

        $cachedGemmaResult = Get-CachedGemmaResult -VideoId $videoId -Model $selectedModel -PresetId $presetId -TranscriptHash $transcriptHash
        if ($null -ne $cachedGemmaResult) {
            $cachedLlmRunId = Get-BundleRunId -Prefix 'llm'
            $bundleLlm = Write-BundleLlmRunArtifacts `
                -VideoId $videoId `
                -RunId $cachedLlmRunId `
                -Mode 'cached' `
                -Status 'completed' `
                -Model $selectedModel `
                -PresetId $presetId `
                -TranscriptHash $transcriptHash `
                -TranscriptText $transcriptText `
                -ResultFilePath $cachedGemmaResult.ResultPath `
                -ActivityText 'Loaded cached Ollama result.'
            $cachedStatus = "Loaded cached Ollama result from '$selectedModel' for preset '$presetId'."
            $cachedActivity = ("[cache]`nLoaded cached Ollama result from '{0}'.`nResult file: {1}" -f $selectedModel, $cachedGemmaResult.ResultPath)
            Apply-GemmaResultToUi -GemmaPayload $cachedGemmaResult.Payload -Model $selectedModel -PresetId $presetId -Status $cachedStatus -ResultPath $cachedGemmaResult.ResultPath -ActivityText $cachedActivity
            if ($script:lastResult) {
                $script:lastResult.BundleLlmResultPath = [string]$bundleLlm.ResultPath
                $script:lastResult.BundlePromptPath = [string]$bundleLlm.PromptPath
                $script:lastResult.BundleDisplayPath = [string]$bundleLlm.DisplayPath
            }
            Set-Status -Message "Loaded cached Ollama result for video ID: $videoId" -Color 'Green'
            Write-Log -Level 'INFO' -Message ("Loaded cached Ollama result without starting a new process. VideoId: {0} | Preset: {1} | Model: {2}" -f $videoId, $presetId, $selectedModel)
            Update-GemmaButtonState
            return
        }

        $resultPath = Get-GuiGemmaResultPath -VideoId $videoId -Model $selectedModel -PresetId $presetId -TranscriptHash $transcriptHash
        $inputPath = Get-GuiGemmaInputPath -VideoId $videoId -Model $selectedModel -PresetId $presetId -TranscriptHash $transcriptHash
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $stdoutPath = Join-Path $script:logDir ("gemma-{0}-{1}.stdout.log" -f $videoId, $timestamp)
        $stderrPath = Join-Path $script:logDir ("gemma-{0}-{1}.stderr.log" -f $videoId, $timestamp)

        Write-TextFileUtf8 -Path $inputPath -Text (Normalize-TranscriptSourceText -Text $transcriptText)
        $gemmaArgs = @(
            'run',
            '--model', $selectedModel,
            '--preset', $presetId,
            '--input-file', $inputPath,
            '--output-file', $resultPath,
            '--timeout-seconds', [string]$script:ollamaTimeoutSeconds
        )

        Write-Log -Level 'INFO' -Message ("Ollama processing requested. VideoId: {0} | Preset: {1} | Model: {2}" -f $videoId, $presetId, $selectedModel)
        $script:isProcessingGemma = $true
        $FetchButton.IsEnabled = $false
        $WhisperTranscriptButton.IsEnabled = $false
        $OpenTranscriptStudioButton.IsEnabled = $false
        $RunGemmaButton.IsEnabled = $false
        $OllamaModelComboBox.IsEnabled = $false
        $GemmaPresetComboBox.IsEnabled = $false
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        if ($script:lastResult) {
            $script:lastResult.GemmaDisplayText = ''
            $script:lastResult.GemmaActivityText = ("[request]`nModel: {0}`nPreset: {1}`nTranscript hash: {2}" -f $selectedModel, $presetId, $transcriptHash)
            $script:lastResult.GemmaStatus = "Preparing Ollama model '$selectedModel' for preset '$presetId'..."
            $script:lastResult.GemmaModel = $selectedModel
            $script:lastResult.GemmaPresetId = $presetId
            Apply-ResultToUi -Result $script:lastResult
        }
        Set-Status -Message 'Preparing Ollama transcript post-processing...' -Color 'DarkBlue'

        $processStart = Start-LocalLlmTool -Arguments $gemmaArgs -StdOutPath $stdoutPath -StdErrPath $stderrPath
        $script:activeGemmaProcess = $processStart.Process
        $script:activeGemmaStdOutPath = [string]$processStart.StdOutPath
        $script:activeGemmaStdErrPath = [string]$processStart.StdErrPath
        $script:activeGemmaResultPath = [string]$resultPath
        $script:activeGemmaInputPath = [string]$inputPath
        $script:activeGemmaVideoId = [string]$videoId
        $script:activeGemmaPresetId = [string]$presetId
        $script:activeGemmaModel = [string]$selectedModel
        $script:activeGemmaTranscriptHash = [string]$transcriptHash
        $script:activeGemmaBundleRunId = Get-BundleRunId -Prefix 'llm'
        Add-BundleEvent -VideoId $videoId -Type 'llm_started' -RunId $script:activeGemmaBundleRunId -Status 'running' -Details ([ordered]@{
            model          = [string]$selectedModel
            presetId       = [string]$presetId
            transcriptHash = [string]$transcriptHash
        })
        if ($script:lastResult) {
            $script:lastResult.GemmaStatus = "Waiting for Ollama model '$selectedModel' to respond..."
            Apply-ResultToUi -Result $script:lastResult
        }
        Set-Status -Message ("Running Ollama model '{0}' in the background..." -f $selectedModel) -Color 'DarkBlue'
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        Update-WhisperButtonState
        Update-GemmaButtonState
        Update-ActiveGemmaRunUi
        $script:gemmaPollTimer.Start()
        Write-Log -Level 'INFO' -Message ("Ollama helper process started. VideoId: {0} | Preset: {1} | Model: {2} | PID: {3}" -f $videoId, $presetId, $selectedModel, $script:activeGemmaProcess.Id)
    }
    catch {
        Write-Log -Level 'ERROR' -Message ("Ollama processing failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
        if ($script:activeGemmaProcess) {
            try {
                $script:activeGemmaProcess.Refresh()
                if (-not $script:activeGemmaProcess.HasExited) {
                    $script:activeGemmaProcess.Kill()
                }
            }
            catch {
            }
        }
        Clear-ActiveGemmaRunState
        $script:isProcessingGemma = $false
        $FetchButton.IsEnabled = $true
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        if ($script:lastResult) {
            $script:lastResult.GemmaStatus = 'Ollama failed.'
            $script:lastResult.GemmaModel = $selectedModel
            Apply-ResultToUi -Result $script:lastResult
        }
        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($_.Exception.Message + $logHint) -Color 'Red'
        [System.Windows.MessageBox]::Show(($_.Exception.Message + $logHint), 'Ollama Failed', 'OK', 'Error') | Out-Null
        Update-WhisperButtonState
        Update-GemmaButtonState
    }
})

$StopGemmaButton.Add_Click({
    if (-not $script:isProcessingGemma) {
        return
    }

    Stop-ActiveGemmaRun -Reason 'Ollama run stopped by user.'
})

$UseCookiesCheck.Add_Checked({
    $BrowserOptionsExpander.IsExpanded = $true
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
        $BrowserOptionsExpander.IsExpanded = $true
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

$window.Add_Closing({
    if ($script:whisperPollTimer) {
        $script:whisperPollTimer.Stop()
    }

    if ($script:gemmaPollTimer) {
        $script:gemmaPollTimer.Stop()
    }

    if ($script:activeWhisperProcess) {
        try {
            $script:activeWhisperProcess.Refresh()
            if (-not $script:activeWhisperProcess.HasExited) {
                Write-Log -Level 'WARN' -Message ("Stopping active Whisper helper because the window is closing. PID: {0}" -f $script:activeWhisperProcess.Id)
                $script:activeWhisperProcess.Kill()
            }
        }
        catch {
            Write-Log -Level 'WARN' -Message ("Failed to stop active Whisper helper during window close:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
        }
        finally {
            Clear-ActiveWhisperRunState
        }
    }

    if ($script:activeGemmaProcess) {
        try {
            $script:activeGemmaProcess.Refresh()
            if (-not $script:activeGemmaProcess.HasExited) {
                Write-Log -Level 'WARN' -Message ("Stopping active Ollama helper because the window is closing. PID: {0}" -f $script:activeGemmaProcess.Id)
                $script:activeGemmaProcess.Kill()
            }
        }
        catch {
            Write-Log -Level 'WARN' -Message ("Failed to stop active Ollama helper during window close:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
        }
        finally {
            Clear-ActiveGemmaRunState
        }
    }
})

Write-Log -Level 'INFO' -Message 'GUI window opening.'
$null = $window.ShowDialog()
Write-Log -Level 'INFO' -Message 'GUI window closed.'








