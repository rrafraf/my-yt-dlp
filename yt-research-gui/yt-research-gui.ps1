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
$script:guiPythonProjectRoot = $PSScriptRoot
$script:guiWhisperModule = 'yt_research_gui_whisper'
$script:guiWhisperApp = 'yt-research-gui-whisper'
$script:guiWhisperModel = 'turbo'
$script:guiWhisperLanguage = 'en'
$script:activeWhisperProcess = $null
$script:activeWhisperStdOutPath = ''
$script:activeWhisperStdErrPath = ''
$script:activeWhisperTranscriptPath = ''
$script:activeWhisperAudioPath = ''
$script:activeWhisperVideoId = ''
$script:activeWhisperOutputSummary = ''
$script:whisperPollTimer = $null

function Load-LoggingSettings {
    $script:currentLogLevel = 'INFO'
    $script:logRetentionDays = 14

    if (-not (Test-Path -LiteralPath $script:loggingConfigPath -PathType Leaf)) {
        return
    }

    try {
        $config = Get-Content -LiteralPath $script:loggingConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        return
    }

    if ($null -eq $config -or -not ($config.PSObject.Properties.Name -contains 'logging')) {
        return
    }

    $logging = $config.logging
    if ($null -eq $logging) {
        return
    }

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
    return [pscustomobject]@{
        Text           = $cachedTranscript
        Status         = 'Loaded cached local Whisper transcript.'
        Source         = Split-Path -Leaf $transcriptPath
        Output         = ''
        AudioPath      = if ($audioFile) { $audioFile.FullName } else { '' }
        TranscriptPath = $transcriptPath
        UsedCache      = $true
    }
}

function Get-WhisperTranscriptResult {
    param(
        [string]$TranscriptPath,
        [string]$AudioPath,
        [string]$Output = '',
        [bool]$UsedCache = $false
    )

    $text = Read-TextFile -Path $TranscriptPath
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
    $ffmpegBin = Get-RepoFfmpegBinDirectory
    $whisperArgs = @(
        '--audio', $audioFile.FullName,
        '--output', $transcriptPath,
        '--ffmpeg-dir', $ffmpegBin,
        '--model', $script:guiWhisperModel,
        '--language', $script:guiWhisperLanguage
    )

    return [pscustomobject]@{
        Mode           = 'launch'
        VideoId        = $VideoId
        AudioPath      = $audioFile.FullName
        TranscriptPath = $transcriptPath
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
                    <TextBlock Grid.Row="3" Text="Transcript Text" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <TextBox Grid.Row="4" x:Name="TranscriptTextBox" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
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
$WhisperTranscriptButton = $window.FindName('WhisperTranscriptButton')
$WhisperProgressText = $window.FindName('WhisperProgressText')
$WhisperActivityTextBox = $window.FindName('WhisperActivityTextBox')
$TranscriptTextBox = $window.FindName('TranscriptTextBox')
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
        $whisperResult = Get-WhisperTranscriptResult -TranscriptPath $script:activeWhisperTranscriptPath -AudioPath $script:activeWhisperAudioPath -Output $outputText
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

    if ($script:isFetching) {
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

function Apply-ResultToUi {
    param($Result)

    if ($null -eq $Result) {
        $MetaGrid.ItemsSource = $null
        $ChaptersGrid.ItemsSource = $null
        $DescriptionTextBox.Text = ''
        $TranscriptTextBox.Text = ''
        $TranscriptStatusText.Text = ''
        Set-WhisperActivityUi -ProgressText 'Whisper idle.' -ActivityText ''
        $RawJsonTextBox.Text = ''
        Update-WhisperButtonState
        return
    }

    $MetaGrid.ItemsSource = $Result.MetaRows
    $ChaptersGrid.ItemsSource = $Result.Chapters
    $DescriptionTextBox.Text = [string]$Result.Description
    $TranscriptTextBox.Text = [string]$Result.Transcript
    $TranscriptStatusText.Text = Get-TranscriptStatusDisplayText -Result $Result
    $RawJsonTextBox.Text = [string]$Result.RawJson
    Update-WhisperButtonState
}

$script:isFetching = $false
$script:isTranscribingWhisper = $false
$script:whisperPollTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:whisperPollTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:whisperPollTimer.Add_Tick({
    Update-ActiveWhisperRunUi
})
Set-WhisperActivityUi -ProgressText 'Whisper idle.' -ActivityText ''
Update-WhisperButtonState

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
            return
        }

        $script:activeWhisperProcess = $whisperStart.Process
        $script:activeWhisperStdOutPath = [string]$whisperStart.StdOutPath
        $script:activeWhisperStdErrPath = [string]$whisperStart.StdErrPath
        $script:activeWhisperTranscriptPath = [string]$whisperStart.TranscriptPath
        $script:activeWhisperAudioPath = [string]$whisperStart.AudioPath
        $script:activeWhisperVideoId = $videoId
        $TranscriptStatusText.Text = 'Running local Whisper transcription in background. The transcript text will update when it finishes.'
        Set-Status -Message 'Running local Whisper transcription in background...' -Color 'DarkBlue'
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        Update-WhisperButtonState
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
})

Write-Log -Level 'INFO' -Message 'GUI window opening.'
$null = $window.ShowDialog()
Write-Log -Level 'INFO' -Message 'GUI window closed.'








