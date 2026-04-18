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
$script:currentLogLevel = 'INFO'
$script:logRetentionDays = 14
$script:guiDataRoot = Join-Path $PSScriptRoot 'data'
$script:guiAudioDir = Join-Path $script:guiDataRoot 'audio'
$script:guiTranscriptDir = Join-Path $script:guiDataRoot 'transcripts'
$script:guiLlmResultDir = Join-Path $script:guiDataRoot 'llm-results'
$script:guiPythonProjectRoot = $PSScriptRoot
$script:guiWhisperModule = 'yt_research_gui_whisper'
$script:guiWhisperApp = 'yt-research-gui-whisper'
$script:guiWhisperModel = 'turbo'
$script:guiWhisperLanguage = 'en'
$script:localLlmToolRoot = Join-Path $script:repoRoot 'tools\local_llm_text'
$script:localLlmToolScript = Join-Path $script:localLlmToolRoot 'cli.py'
$script:ollamaModel = 'gemma4'
$script:ollamaTimeoutSeconds = 180
$script:activeWhisperProcess = $null
$script:activeWhisperStdOutPath = ''
$script:activeWhisperStdErrPath = ''
$script:activeWhisperTranscriptPath = ''
$script:activeWhisperTimingPath = ''
$script:activeWhisperAudioPath = ''
$script:activeWhisperVideoId = ''
$script:activeWhisperOutputSummary = ''
$script:whisperPollTimer = $null
$script:activeGemmaProcess = $null
$script:activeGemmaStdOutPath = ''
$script:activeGemmaStdErrPath = ''
$script:activeGemmaResultPath = ''
$script:activeGemmaInputPath = ''
$script:activeGemmaVideoId = ''
$script:activeGemmaPresetId = ''
$script:activeGemmaTranscriptHash = ''
$script:gemmaPollTimer = $null
$script:gemmaPresets = @()
$script:gemmaHelperAvailable = $false

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
    $fileName = '{0}--{1}--{2}.json' -f (ConvertTo-SafeCacheSegment -Value $Model), (ConvertTo-SafeCacheSegment -Value $PresetId), $TranscriptHash
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
    $fileName = '{0}--{1}--{2}.input.txt' -f (ConvertTo-SafeCacheSegment -Value $Model), (ConvertTo-SafeCacheSegment -Value $PresetId), $TranscriptHash
    return (Join-Path $cacheDir $fileName)
}

function Write-TextFileUtf8 {
    param(
        [string]$Path,
        [string]$Text
    )

    Ensure-DirectoryExists -Path (Split-Path -Parent $Path)
    [System.IO.File]::WriteAllText($Path, [string]$Text, [System.Text.Encoding]::UTF8)
}

function Get-ReadyGemmaStatusText {
    param($Result)

    if (-not $script:gemmaHelperAvailable) {
        return 'Gemma 4 helper is unavailable in this environment.'
    }

    if ($null -eq $Result) {
        return 'Load a video to use Gemma 4.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$Result.Transcript)) {
        return 'No transcript available for Gemma 4 yet.'
    }

    return 'Ready to run Gemma 4 on the current transcript.'
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
    $Result.GemmaStatus = $nextStatus
    $Result.GemmaPresetId = ''
    $Result.GemmaSourceTextHash = ''
    $Result.GemmaResultPath = ''
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
        [string]$PresetId,
        [string]$TranscriptHash
    )

    if ([string]::IsNullOrWhiteSpace($VideoId) -or [string]::IsNullOrWhiteSpace($PresetId) -or [string]::IsNullOrWhiteSpace($TranscriptHash)) {
        return $null
    }

    $resultPath = Get-GuiGemmaResultPath -VideoId $VideoId -Model $script:ollamaModel -PresetId $PresetId -TranscriptHash $TranscriptHash
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

    if ([string]$payload.sourceTextHash -ne $TranscriptHash -or [string]$payload.presetId -ne $PresetId -or [string]$payload.model -ne $script:ollamaModel) {
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
        Mode           = 'launch'
        VideoId        = $VideoId
        AudioPath      = $audioFile.FullName
        TranscriptPath = $transcriptPath
        TimingPath     = $timingPath
        Arguments      = $whisperArgs
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
        VideoUrl         = if ([string]::IsNullOrWhiteSpace([string]$info.webpage_url)) { [string]$Url } else { [string]$info.webpage_url }
        TranscriptTimingJson = ''
        TranscriptTimingPath = ''
        GemmaDisplayText = ''
        GemmaStatus      = ''
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
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="180"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid Grid.Row="0" Margin="0,0,0,6">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" x:Name="TranscriptStatusText" Margin="0,0,8,0" TextWrapping="Wrap"/>
                        <Button Grid.Column="1" x:Name="WhisperTranscriptButton" Width="190" Height="28" Content="Transcribe with Whisper" IsEnabled="False"/>
                    </Grid>
                    <TextBlock Grid.Row="1" x:Name="WhisperProgressText" Margin="0,0,0,6" Foreground="DimGray" Text="Whisper idle." TextWrapping="Wrap"/>
                    <GroupBox Grid.Row="2" Header="Whisper Activity" Margin="0,0,0,8">
                        <TextBox x:Name="WhisperActivityTextBox" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                 FontFamily="Consolas" FontSize="12"/>
                    </GroupBox>
                    <Grid Grid.Row="3" Margin="0,0,0,6">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="280"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Gemma Prompt:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,8,0"/>
                        <ComboBox Grid.Column="1" x:Name="GemmaPresetComboBox" Height="28" Margin="0,0,8,0" IsEnabled="False"/>
                        <Button Grid.Column="2" x:Name="RunGemmaButton" Width="150" Height="28" Margin="0,0,8,0" Content="Run Gemma 4" IsEnabled="False"/>
                    </Grid>
                    <TextBlock Grid.Row="4" x:Name="GemmaStatusText" Margin="0,0,0,6" Foreground="DimGray" Text="Gemma 4 idle." TextWrapping="Wrap"/>
                    <Grid Grid.Row="5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="12"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <GroupBox Grid.Column="0" Header="Transcript Text">
                            <TextBox x:Name="TranscriptTextBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                                     FontFamily="Consolas" FontSize="13"/>
                        </GroupBox>
                        <GroupBox Grid.Column="2" Header="Gemma Result">
                            <TextBox x:Name="GemmaResultTextBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                                     FontFamily="Consolas" FontSize="13"/>
                        </GroupBox>
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
$WhisperTranscriptButton = $window.FindName('WhisperTranscriptButton')
$WhisperProgressText = $window.FindName('WhisperProgressText')
$WhisperActivityTextBox = $window.FindName('WhisperActivityTextBox')
$GemmaPresetComboBox = $window.FindName('GemmaPresetComboBox')
$RunGemmaButton = $window.FindName('RunGemmaButton')
$GemmaStatusText = $window.FindName('GemmaStatusText')
$TranscriptTextBox = $window.FindName('TranscriptTextBox')
$TranscriptTimingTextBox = $window.FindName('TranscriptTimingTextBox')
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
    if ($WhisperActivityTextBox.Text -ne $nextActivityText) {
        $WhisperActivityTextBox.Text = $nextActivityText
        $WhisperActivityTextBox.ScrollToEnd()
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
        [string]$ResultText
    )

    $nextStatus = if ([string]::IsNullOrWhiteSpace($StatusText)) { Get-ReadyGemmaStatusText -Result $script:lastResult } else { $StatusText }
    $GemmaStatusText.Text = $nextStatus

    $nextResult = [string]$ResultText
    if ($GemmaResultTextBox.Text -ne $nextResult) {
        $GemmaResultTextBox.Text = $nextResult
        $GemmaResultTextBox.ScrollToHome()
    }
}

function Normalize-WhisperLogText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    return (([string]$Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Get-WhisperLogTailText {
    param(
        [string]$Path,
        [int]$MaxLines = 40
    )

    $normalized = Normalize-WhisperLogText -Text (Read-TextFile -Path $Path)
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $lines = @($normalized -split "`n")
    if ($lines.Count -gt $MaxLines) {
        $lines = $lines[($lines.Count - $MaxLines)..($lines.Count - 1)]
    }

    return (($lines | ForEach-Object { $_.TrimEnd() }) -join [Environment]::NewLine).Trim()
}

function Get-WhisperCombinedOutput {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $stdoutText = Normalize-WhisperLogText -Text (Read-TextFile -Path $StdOutPath)
    $stderrText = Normalize-WhisperLogText -Text (Read-TextFile -Path $StdErrPath)

    if (-not [string]::IsNullOrWhiteSpace($stdoutText)) {
        $parts.Add($stdoutText)
    }

    if (-not [string]::IsNullOrWhiteSpace($stderrText)) {
        $parts.Add($stderrText)
    }

    return ($parts -join [Environment]::NewLine).Trim()
}

function Get-WhisperActivityDisplayText {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    $stdoutTail = Get-WhisperLogTailText -Path $StdOutPath -MaxLines 25
    $stderrTail = Get-WhisperLogTailText -Path $StdErrPath -MaxLines 25
    $sections = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($stdoutTail)) {
        $sections.Add('[stdout]')
        $sections.Add($stdoutTail)
    }

    if (-not [string]::IsNullOrWhiteSpace($stderrTail)) {
        if ($sections.Count -gt 0) {
            $sections.Add('')
        }

        $sections.Add('[stderr]')
        $sections.Add($stderrTail)
    }

    return ($sections -join [Environment]::NewLine).Trim()
}

function Get-WhisperProgressDisplayText {
    param(
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    $stdoutText = Normalize-WhisperLogText -Text (Read-TextFile -Path $StdOutPath)
    $stderrText = Normalize-WhisperLogText -Text (Read-TextFile -Path $StdErrPath)

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
    $script:lastResult.TranscriptTimingJson = [string]$WhisperResult.TimingJson
    $script:lastResult.TranscriptTimingPath = [string]$WhisperResult.TimingPath
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
    $script:activeWhisperOutputSummary = $outputText
    $script:isTranscribingWhisper = $false
    $FetchButton.IsEnabled = $true
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow

    if ($didSucceed) {
        $whisperResult = Get-WhisperTranscriptResult -TranscriptPath $script:activeWhisperTranscriptPath -TimingPath $script:activeWhisperTimingPath -AudioPath $script:activeWhisperAudioPath -Output $outputText
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
    $script:activeGemmaTranscriptHash = ''
}

function Apply-GemmaResultToUi {
    param(
        $GemmaPayload,
        [string]$PresetId,
        [string]$Status,
        [string]$ResultPath
    )

    if ($null -eq $script:lastResult -or $null -eq $GemmaPayload) {
        return
    }

    $script:lastResult.GemmaDisplayText = [string]$GemmaPayload.displayText
    $script:lastResult.GemmaStatus = [string]$Status
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
    $exitCode = $script:activeGemmaProcess.ExitCode
    $outputText = Get-WhisperCombinedOutput -StdOutPath $script:activeGemmaStdOutPath -StdErrPath $script:activeGemmaStdErrPath
    $didSucceed = Test-GemmaRunSucceeded -ExitCode $exitCode -ResultPath $script:activeGemmaResultPath -OutputText $outputText
    $exitCodeDisplay = if ([string]::IsNullOrWhiteSpace([string]$exitCode)) { 'unavailable' } else { [string]$exitCode }
    $script:isProcessingGemma = $false
    $FetchButton.IsEnabled = $true
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow

    if ($didSucceed) {
        try {
            $gemmaPayload = Read-GemmaResultFile -Path $script:activeGemmaResultPath
            $gemmaStatus = "Loaded Gemma 4 result for preset '$presetId'."
            Apply-GemmaResultToUi -GemmaPayload $gemmaPayload -PresetId $presetId -Status $gemmaStatus -ResultPath $script:activeGemmaResultPath
            Set-Status -Message "Loaded Gemma 4 result for video ID: $videoId" -Color 'Green'
            Write-Log -Level 'INFO' -Message ("Gemma 4 processing completed. VideoId: {0} | Preset: {1} | ExitCode: {2}" -f $videoId, $presetId, $exitCodeDisplay)
        }
        catch {
            $errorMessage = "Gemma 4 completed, but the result file could not be loaded. See the logs for details."
            if ($script:lastResult) {
                $script:lastResult.GemmaStatus = $errorMessage
                Apply-ResultToUi -Result $script:lastResult
            }

            $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
            Set-Status -Message ($errorMessage + $logHint) -Color 'Red'
            Write-Log -Level 'ERROR' -Message ("Gemma 4 result loading failed. VideoId: {0} | Preset: {1}`n{2}" -f $videoId, $presetId, (Get-ErrorDetails -ErrorObject $_))
            [System.Windows.MessageBox]::Show(($errorMessage + $logHint), 'Gemma 4 Failed', 'OK', 'Error') | Out-Null
        }
    }
    else {
        $errorMessage = "Gemma 4 processing failed (exit code $exitCodeDisplay)."
        if ($script:lastResult) {
            $script:lastResult.GemmaStatus = 'Gemma 4 failed.'
            Apply-ResultToUi -Result $script:lastResult
        }

        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($errorMessage + $logHint) -Color 'Red'
        Write-Log -Level 'ERROR' -Message ("Gemma 4 processing failed. VideoId: {0} | Preset: {1} | ExitCode: {2}`n{3}" -f $videoId, $presetId, $exitCodeDisplay, $outputText)
        [System.Windows.MessageBox]::Show(($errorMessage + $logHint), 'Gemma 4 Failed', 'OK', 'Error') | Out-Null
    }

    Clear-ActiveGemmaRunState
    Update-WhisperButtonState
    Update-GemmaButtonState
}

function Update-ActiveGemmaRunUi {
    if ($null -eq $script:activeGemmaProcess) {
        return
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
    $WhisperTranscriptButton.IsEnabled = $false
    $WhisperTranscriptButton.Content = 'Transcribe with Whisper'

    if ($script:isTranscribingWhisper) {
        $WhisperTranscriptButton.Content = 'Whisper Running...'
        return
    }

    if ($script:isFetching -or $script:isProcessingGemma) {
        return
    }

    if ($null -eq $script:lastResult) {
        return
    }

    if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.VideoId)) {
        return
    }

    if (Test-WhisperTranscriptCached -VideoId ([string]$script:lastResult.VideoId)) {
        $WhisperTranscriptButton.Content = 'Load Whisper Transcript'
    }

    $WhisperTranscriptButton.IsEnabled = $true
}

function Update-GemmaButtonState {
    $RunGemmaButton.IsEnabled = $false
    $RunGemmaButton.Content = 'Run Gemma 4'
    $GemmaPresetComboBox.IsEnabled = $false

    if (-not $script:gemmaHelperAvailable) {
        return
    }

    if ($script:isProcessingGemma) {
        $RunGemmaButton.Content = 'Gemma 4 Running...'
        return
    }

    if ($script:isFetching -or $script:isTranscribingWhisper) {
        return
    }

    if ($null -eq $script:lastResult) {
        return
    }

    if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.Transcript)) {
        return
    }

    $GemmaPresetComboBox.IsEnabled = ($script:gemmaPresets.Count -gt 0)
    $selectedPresetId = Get-SelectedGemmaPresetId
    if ([string]::IsNullOrWhiteSpace($selectedPresetId)) {
        return
    }

    $RunGemmaButton.IsEnabled = $true
}

function Initialize-GemmaPresets {
    $script:gemmaHelperAvailable = $false
    $script:gemmaPresets = @()
    $GemmaPresetComboBox.ItemsSource = $null

    try {
        $presets = @(Get-GemmaPresetDefinitions)
        $script:gemmaPresets = $presets
        $script:gemmaHelperAvailable = ($presets.Count -gt 0)
        $GemmaPresetComboBox.DisplayMemberPath = 'label'
        $GemmaPresetComboBox.SelectedValuePath = 'id'
        $GemmaPresetComboBox.ItemsSource = $presets
        if ($presets.Count -gt 0) {
            $GemmaPresetComboBox.SelectedIndex = 0
            Write-Log -Level 'INFO' -Message ("Loaded {0} Gemma preset definitions." -f $presets.Count)
        }
        else {
            $GemmaStatusText.Text = 'No Gemma 4 presets were found.'
            Write-Log -Level 'WARN' -Message 'Gemma preset listing returned no presets.'
        }
    }
    catch {
        $script:gemmaHelperAvailable = $false
        $script:gemmaPresets = @()
        $GemmaPresetComboBox.ItemsSource = $null
        $GemmaStatusText.Text = 'Gemma 4 helper is unavailable. See the status bar or logs for details.'
        Write-Log -Level 'WARN' -Message ("Gemma preset initialization failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
    }

    if ($script:gemmaHelperAvailable) {
        Set-GemmaUi -StatusText (Get-ReadyGemmaStatusText -Result $script:lastResult) -ResultText ''
    }

    Update-GemmaButtonState
}

function Apply-ResultToUi {
    param($Result)

    if ($null -eq $Result) {
        $MetaGrid.ItemsSource = $null
        $ChaptersGrid.ItemsSource = $null
        $DescriptionTextBox.Text = ''
        $TranscriptTextBox.Text = ''
        $TranscriptTimingTextBox.Text = ''
        $TranscriptStatusText.Text = ''
        Set-GemmaUi -StatusText (Get-ReadyGemmaStatusText -Result $null) -ResultText ''
        Set-WhisperActivityUi -ProgressText 'Whisper idle.' -ActivityText ''
        $RawJsonTextBox.Text = ''
        Update-WhisperButtonState
        Update-GemmaButtonState
        return
    }

    $MetaGrid.ItemsSource = $Result.MetaRows
    $ChaptersGrid.ItemsSource = $Result.Chapters
    $DescriptionTextBox.Text = [string]$Result.Description
    $TranscriptTextBox.Text = [string]$Result.Transcript
    $TranscriptTimingTextBox.Text = if ([string]::IsNullOrWhiteSpace([string]$Result.TranscriptTimingJson)) { 'No local Whisper timing metadata loaded.' } else { [string]$Result.TranscriptTimingJson }
    $TranscriptStatusText.Text = Get-TranscriptStatusDisplayText -Result $Result
    Set-GemmaUi -StatusText (Get-GemmaStatusDisplayText -Result $Result) -ResultText ([string]$Result.GemmaDisplayText)
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
Set-GemmaUi -StatusText (Get-ReadyGemmaStatusText -Result $null) -ResultText ''
Update-WhisperButtonState
Update-GemmaButtonState
Initialize-GemmaPresets

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
        $WhisperTranscriptButton.IsEnabled = $false
        $RunGemmaButton.IsEnabled = $false
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
        $url = if ([string]::IsNullOrWhiteSpace([string]$script:lastResult.VideoUrl)) { $UrlBox.Text.Trim() } else { [string]$script:lastResult.VideoUrl }
        Assert-ValidYoutubeUrl -Url $url

        $useCookies = [bool]$UseCookiesCheck.IsChecked
        $profilePath = $ProfilePathBox.Text.Trim()
        $cookieArgs = @(Get-CookieArgs -UseCookies $useCookies -ProfilePath $profilePath)
        Write-Log -Level 'INFO' -Message ("Whisper transcription requested. VideoId: {0} | URL: {1} | UseCookies: {2}" -f $videoId, $url, $useCookies)

        $script:isTranscribingWhisper = $true
        $FetchButton.IsEnabled = $false
        $WhisperTranscriptButton.IsEnabled = $false
        $RunGemmaButton.IsEnabled = $false
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

$GemmaPresetComboBox.Add_SelectionChanged({
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

    try {
        $videoId = [string]$script:lastResult.VideoId
        $transcriptHash = Get-TextSha256 -Text $transcriptText
        if ([string]::IsNullOrWhiteSpace($transcriptHash)) {
            throw 'The current transcript is empty after normalization.'
        }

        $cachedGemmaResult = Get-CachedGemmaResult -VideoId $videoId -PresetId $presetId -TranscriptHash $transcriptHash
        if ($null -ne $cachedGemmaResult) {
            $cachedStatus = "Loaded cached Gemma 4 result for preset '$presetId'."
            Apply-GemmaResultToUi -GemmaPayload $cachedGemmaResult.Payload -PresetId $presetId -Status $cachedStatus -ResultPath $cachedGemmaResult.ResultPath
            Set-Status -Message "Loaded cached Gemma 4 result for video ID: $videoId" -Color 'Green'
            Write-Log -Level 'INFO' -Message ("Loaded cached Gemma result without starting a new process. VideoId: {0} | Preset: {1}" -f $videoId, $presetId)
            Update-GemmaButtonState
            return
        }

        $resultPath = Get-GuiGemmaResultPath -VideoId $videoId -Model $script:ollamaModel -PresetId $presetId -TranscriptHash $transcriptHash
        $inputPath = Get-GuiGemmaInputPath -VideoId $videoId -Model $script:ollamaModel -PresetId $presetId -TranscriptHash $transcriptHash
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $stdoutPath = Join-Path $script:logDir ("gemma-{0}-{1}.stdout.log" -f $videoId, $timestamp)
        $stderrPath = Join-Path $script:logDir ("gemma-{0}-{1}.stderr.log" -f $videoId, $timestamp)

        Write-TextFileUtf8 -Path $inputPath -Text (Normalize-TranscriptSourceText -Text $transcriptText)
        $gemmaArgs = @(
            'run',
            '--model', $script:ollamaModel,
            '--preset', $presetId,
            '--input-file', $inputPath,
            '--output-file', $resultPath,
            '--timeout-seconds', [string]$script:ollamaTimeoutSeconds
        )

        Write-Log -Level 'INFO' -Message ("Gemma 4 processing requested. VideoId: {0} | Preset: {1} | Model: {2}" -f $videoId, $presetId, $script:ollamaModel)
        $script:isProcessingGemma = $true
        $FetchButton.IsEnabled = $false
        $WhisperTranscriptButton.IsEnabled = $false
        $RunGemmaButton.IsEnabled = $false
        $GemmaPresetComboBox.IsEnabled = $false
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        if ($script:lastResult) {
            $script:lastResult.GemmaStatus = "Preparing Gemma 4 preset '$presetId'..."
            Apply-ResultToUi -Result $script:lastResult
        }
        Set-Status -Message 'Preparing Gemma 4 transcript post-processing...' -Color 'DarkBlue'

        $processStart = Start-LocalLlmTool -Arguments $gemmaArgs -StdOutPath $stdoutPath -StdErrPath $stderrPath
        $script:activeGemmaProcess = $processStart.Process
        $script:activeGemmaStdOutPath = [string]$processStart.StdOutPath
        $script:activeGemmaStdErrPath = [string]$processStart.StdErrPath
        $script:activeGemmaResultPath = [string]$resultPath
        $script:activeGemmaInputPath = [string]$inputPath
        $script:activeGemmaVideoId = [string]$videoId
        $script:activeGemmaPresetId = [string]$presetId
        $script:activeGemmaTranscriptHash = [string]$transcriptHash
        if ($script:lastResult) {
            $script:lastResult.GemmaStatus = "Running Gemma 4 preset '$presetId' in the background."
            Apply-ResultToUi -Result $script:lastResult
        }
        Set-Status -Message 'Running Gemma 4 in the background...' -Color 'DarkBlue'
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        Update-WhisperButtonState
        Update-GemmaButtonState
        Update-ActiveGemmaRunUi
        $script:gemmaPollTimer.Start()
        Write-Log -Level 'INFO' -Message ("Gemma helper process started. VideoId: {0} | Preset: {1} | PID: {2}" -f $videoId, $presetId, $script:activeGemmaProcess.Id)
    }
    catch {
        Write-Log -Level 'ERROR' -Message ("Gemma 4 processing failed:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
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
            $script:lastResult.GemmaStatus = 'Gemma 4 failed.'
            Apply-ResultToUi -Result $script:lastResult
        }
        $logHint = if ($script:logPath) { " See log: $script:logPath" } else { '' }
        Set-Status -Message ($_.Exception.Message + $logHint) -Color 'Red'
        [System.Windows.MessageBox]::Show(($_.Exception.Message + $logHint), 'Gemma 4 Failed', 'OK', 'Error') | Out-Null
        Update-WhisperButtonState
        Update-GemmaButtonState
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
                Write-Log -Level 'WARN' -Message ("Stopping active Gemma helper because the window is closing. PID: {0}" -f $script:activeGemmaProcess.Id)
                $script:activeGemmaProcess.Kill()
            }
        }
        catch {
            Write-Log -Level 'WARN' -Message ("Failed to stop active Gemma helper during window close:`n{0}" -f (Get-ErrorDetails -ErrorObject $_))
        }
        finally {
            Clear-ActiveGemmaRunState
        }
    }
})

Write-Log -Level 'INFO' -Message 'GUI window opening.'
$null = $window.ShowDialog()
Write-Log -Level 'INFO' -Message 'GUI window closed.'








