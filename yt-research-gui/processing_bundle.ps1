function Get-Iso8601Timestamp {
    return (Get-Date).ToString('o')
}

function Get-VideoBundleDirectory {
    param([string]$VideoId)

    $videoSegment = if ([string]::IsNullOrWhiteSpace($VideoId)) { '_no-video-id' } else { (ConvertTo-SafeCacheSegment -Value $VideoId) }
    return (Join-Path $script:guiBundleDir $videoSegment)
}

function Get-VideoBundleManifestPath {
    param([string]$VideoId)

    return (Join-Path (Get-VideoBundleDirectory -VideoId $VideoId) 'manifest.json')
}

function Get-VideoBundleEventsPath {
    param([string]$VideoId)

    return (Join-Path (Get-VideoBundleDirectory -VideoId $VideoId) 'events.ndjson')
}

function Ensure-VideoBundleStorage {
    param([string]$VideoId)

    Ensure-DirectoryExists -Path $script:guiDataRoot
    Ensure-DirectoryExists -Path $script:guiBundleDir

    $bundleRoot = Get-VideoBundleDirectory -VideoId $VideoId
    foreach ($path in @(
        $bundleRoot,
        (Join-Path $bundleRoot 'source'),
        (Join-Path $bundleRoot 'audio'),
        (Join-Path $bundleRoot 'audio\extractions'),
        (Join-Path $bundleRoot 'whisper'),
        (Join-Path $bundleRoot 'whisper\runs'),
        (Join-Path $bundleRoot 'llm'),
        (Join-Path $bundleRoot 'llm\runs'),
        (Join-Path $bundleRoot 'exports')
    )) {
        Ensure-DirectoryExists -Path $path
    }

    return $bundleRoot
}

function Get-BundleRunId {
    param([string]$Prefix)

    $safePrefix = if ([string]::IsNullOrWhiteSpace($Prefix)) { 'run' } else { (ConvertTo-SafeCacheSegment -Value $Prefix) }
    return ('{0}-{1}' -f $safePrefix, (Get-Date -Format 'yyyyMMdd-HHmmssfff'))
}

function Ensure-ObjectProperty {
    param(
        $Object,
        [string]$Name,
        $DefaultValue
    )

    if ($null -eq $Object) {
        return
    }

    if (-not ($Object.PSObject.Properties.Name -contains $Name)) {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $DefaultValue
    }
}

function New-VideoBundleManifest {
    param(
        [string]$VideoId,
        [string]$Title = '',
        [string]$SourceUrl = ''
    )

    $now = Get-Iso8601Timestamp
    return [pscustomobject]@{
        bundleVersion = 1
        videoId       = [string]$VideoId
        title         = [string]$Title
        sourceUrl     = [string]$SourceUrl
        createdAt     = $now
        updatedAt     = $now
        toolVersions  = [pscustomobject]@{
            whisperModel           = $script:guiWhisperModel
            whisperLanguage        = $script:guiWhisperLanguage
            ollamaDefaultModel     = $script:ollamaModel
            ollamaTimeoutSeconds   = $script:ollamaTimeoutSeconds
        }
        counters      = [pscustomobject]@{
            fetches                  = 0
            audioExtractions         = 0
            whisperRuns              = 0
            llmRuns                  = 0
            transcriptStudioLaunches = 0
        }
        latest        = [pscustomobject]@{
            fetchRunId              = ''
            metadataJson            = ''
            metadataFetchLog        = ''
            ytCaption               = ''
            ytTranscript            = ''
            subtitleFetchLog        = ''
            audio                   = ''
            audioExtractionRunId    = ''
            audioExtractionLog      = ''
            whisperRunId            = ''
            whisperTranscript       = ''
            whisperTimings          = ''
            whisperStdOutLog        = ''
            whisperStdErrLog        = ''
            llmRunId                = ''
            llmModel                = ''
            llmPresetId             = ''
            llmInputTranscript      = ''
            llmPrompt               = ''
            llmResult               = ''
            llmDisplay              = ''
            llmStdOutLog            = ''
            llmStdErrLog            = ''
            currentTranscript       = ''
            currentTranscriptMode   = ''
            currentTimings          = ''
            transcriptStudioSession = ''
        }
        artifacts     = @()
    }
}

function Ensure-VideoBundleManifestDefaults {
    param(
        $Manifest,
        [string]$VideoId = '',
        [string]$Title = '',
        [string]$SourceUrl = ''
    )

    if ($null -eq $Manifest) {
        $Manifest = New-VideoBundleManifest -VideoId $VideoId -Title $Title -SourceUrl $SourceUrl
    }

    Ensure-ObjectProperty -Object $Manifest -Name 'bundleVersion' -DefaultValue 1
    Ensure-ObjectProperty -Object $Manifest -Name 'videoId' -DefaultValue ([string]$VideoId)
    Ensure-ObjectProperty -Object $Manifest -Name 'title' -DefaultValue ([string]$Title)
    Ensure-ObjectProperty -Object $Manifest -Name 'sourceUrl' -DefaultValue ([string]$SourceUrl)
    Ensure-ObjectProperty -Object $Manifest -Name 'createdAt' -DefaultValue (Get-Iso8601Timestamp)
    Ensure-ObjectProperty -Object $Manifest -Name 'updatedAt' -DefaultValue (Get-Iso8601Timestamp)
    Ensure-ObjectProperty -Object $Manifest -Name 'toolVersions' -DefaultValue ([pscustomobject]@{})
    Ensure-ObjectProperty -Object $Manifest -Name 'counters' -DefaultValue ([pscustomobject]@{})
    Ensure-ObjectProperty -Object $Manifest -Name 'latest' -DefaultValue ([pscustomobject]@{})
    Ensure-ObjectProperty -Object $Manifest -Name 'artifacts' -DefaultValue @()

    Ensure-ObjectProperty -Object $Manifest.toolVersions -Name 'whisperModel' -DefaultValue $script:guiWhisperModel
    Ensure-ObjectProperty -Object $Manifest.toolVersions -Name 'whisperLanguage' -DefaultValue $script:guiWhisperLanguage
    Ensure-ObjectProperty -Object $Manifest.toolVersions -Name 'ollamaDefaultModel' -DefaultValue $script:ollamaModel
    Ensure-ObjectProperty -Object $Manifest.toolVersions -Name 'ollamaTimeoutSeconds' -DefaultValue $script:ollamaTimeoutSeconds

    foreach ($name in @('fetches', 'audioExtractions', 'whisperRuns', 'llmRuns', 'transcriptStudioLaunches')) {
        Ensure-ObjectProperty -Object $Manifest.counters -Name $name -DefaultValue 0
    }

    foreach ($name in @(
        'fetchRunId', 'metadataJson', 'metadataFetchLog', 'ytCaption', 'ytTranscript', 'subtitleFetchLog',
        'audio', 'audioExtractionRunId', 'audioExtractionLog',
        'whisperRunId', 'whisperTranscript', 'whisperTimings', 'whisperStdOutLog', 'whisperStdErrLog',
        'llmRunId', 'llmModel', 'llmPresetId', 'llmInputTranscript', 'llmPrompt', 'llmResult', 'llmDisplay', 'llmStdOutLog', 'llmStdErrLog',
        'currentTranscript', 'currentTranscriptMode', 'currentTimings', 'transcriptStudioSession'
    )) {
        Ensure-ObjectProperty -Object $Manifest.latest -Name $name -DefaultValue ''
    }

    if ($null -eq $Manifest.artifacts) {
        $Manifest.artifacts = @()
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$VideoId)) { $Manifest.videoId = [string]$VideoId }
    if (-not [string]::IsNullOrWhiteSpace([string]$Title)) { $Manifest.title = [string]$Title }
    if (-not [string]::IsNullOrWhiteSpace([string]$SourceUrl)) { $Manifest.sourceUrl = [string]$SourceUrl }

    return $Manifest
}

function Get-VideoBundleManifest {
    param(
        [string]$VideoId,
        [string]$Title = '',
        [string]$SourceUrl = ''
    )

    Ensure-VideoBundleStorage -VideoId $VideoId | Out-Null
    $manifestPath = Get-VideoBundleManifestPath -VideoId $VideoId
    $manifest = $null

    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
        try {
            $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            Write-Log -Level 'WARN' -Message ("Failed to read bundle manifest '{0}'. Recreating it.`n{1}" -f $manifestPath, (Get-ErrorDetails -ErrorObject $_))
            $manifest = $null
        }
    }

    return (Ensure-VideoBundleManifestDefaults -Manifest $manifest -VideoId $VideoId -Title $Title -SourceUrl $SourceUrl)
}

function Save-VideoBundleManifest {
    param(
        [string]$VideoId,
        $Manifest
    )

    if ($null -eq $Manifest) {
        return
    }

    $Manifest.updatedAt = Get-Iso8601Timestamp
    $manifestPath = Get-VideoBundleManifestPath -VideoId $VideoId
    $json = $Manifest | ConvertTo-Json -Depth 40
    Write-TextFileUtf8 -Path $manifestPath -Text $json
}

function Get-RelativePathFromBase {
    param(
        [string]$BasePath,
        [string]$TargetPath
    )

    if ([string]::IsNullOrWhiteSpace($BasePath) -or [string]::IsNullOrWhiteSpace($TargetPath)) {
        return ''
    }

    $resolvedBase = [System.IO.Path]::GetFullPath($BasePath)
    if (-not $resolvedBase.EndsWith('\')) {
        $resolvedBase = $resolvedBase + '\'
    }
    $resolvedTarget = [System.IO.Path]::GetFullPath($TargetPath)
    $baseUri = New-Object System.Uri($resolvedBase)
    $targetUri = New-Object System.Uri($resolvedTarget)
    $relativeUri = $baseUri.MakeRelativeUri($targetUri)
    return ([System.Uri]::UnescapeDataString($relativeUri.ToString())).Replace('/', '\')
}

function Get-FileSha256 {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return ''
    }

    try {
        $hash = Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop
        return ([string]$hash.Hash).ToLowerInvariant()
    }
    catch {
        return ''
    }
}

function New-BundleArtifactRecord {
    param(
        [string]$VideoId,
        [string]$Path,
        [string]$Kind,
        [string]$Role = ''
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }

    $bundleRoot = Get-VideoBundleDirectory -VideoId $VideoId
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    return [pscustomobject]@{
        path       = Get-RelativePathFromBase -BasePath $bundleRoot -TargetPath $item.FullName
        kind       = [string]$Kind
        role       = [string]$Role
        fileName   = $item.Name
        sizeBytes  = [int64]$item.Length
        sha256     = Get-FileSha256 -Path $item.FullName
        modifiedAt = $item.LastWriteTime.ToString('o')
    }
}

function Upsert-BundleArtifact {
    param(
        $Manifest,
        $Artifact
    )

    if ($null -eq $Manifest -or $null -eq $Artifact) {
        return
    }

    $artifacts = @($Manifest.artifacts)
    $updated = $false
    for ($index = 0; $index -lt $artifacts.Count; $index++) {
        if ([string]$artifacts[$index].path -eq [string]$Artifact.path) {
            $artifacts[$index] = $Artifact
            $updated = $true
            break
        }
    }

    if (-not $updated) {
        $artifacts += $Artifact
    }

    $Manifest.artifacts = $artifacts
}

function Write-BundleTextArtifact {
    param(
        [string]$VideoId,
        [string]$RelativePath,
        [string]$Text
    )

    $bundleRoot = Ensure-VideoBundleStorage -VideoId $VideoId
    $targetPath = Join-Path $bundleRoot $RelativePath
    Write-TextFileUtf8 -Path $targetPath -Text $Text
    return $targetPath
}

function Write-BundleJsonArtifact {
    param(
        [string]$VideoId,
        [string]$RelativePath,
        $Value
    )

    $json = $Value | ConvertTo-Json -Depth 40
    return (Write-BundleTextArtifact -VideoId $VideoId -RelativePath $RelativePath -Text $json)
}

function Copy-FileIntoBundle {
    param(
        [string]$VideoId,
        [string]$SourcePath,
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($SourcePath) -or -not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        return ''
    }

    $bundleRoot = Ensure-VideoBundleStorage -VideoId $VideoId
    $targetPath = Join-Path $bundleRoot $RelativePath
    Ensure-DirectoryExists -Path (Split-Path -Parent $targetPath)
    Copy-Item -LiteralPath $SourcePath -Destination $targetPath -Force
    return $targetPath
}

function Add-BundleEvent {
    param(
        [string]$VideoId,
        [string]$Type,
        [string]$RunId = '',
        $Inputs = $null,
        $Outputs = $null,
        $Details = $null,
        [string]$Status = 'completed'
    )

    if ([string]::IsNullOrWhiteSpace($VideoId)) {
        return
    }

    Ensure-VideoBundleStorage -VideoId $VideoId | Out-Null
    $eventPath = Get-VideoBundleEventsPath -VideoId $VideoId
    $payload = [ordered]@{
        timestamp = Get-Iso8601Timestamp
        type      = [string]$Type
        status    = [string]$Status
        videoId   = [string]$VideoId
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$RunId)) { $payload.runId = [string]$RunId }
    if ($null -ne $Inputs) { $payload.inputs = $Inputs }
    if ($null -ne $Outputs) { $payload.outputs = $Outputs }
    if ($null -ne $Details) { $payload.details = $Details }

    $json = $payload | ConvertTo-Json -Depth 40 -Compress
    [System.IO.File]::AppendAllText($eventPath, $json + [Environment]::NewLine, $script:utf8NoBomEncoding)
}

function Set-ManifestLatestPath {
    param(
        $Manifest,
        [string]$PropertyName,
        [string]$VideoId,
        [string]$AbsolutePath
    )

    if ($null -eq $Manifest -or [string]::IsNullOrWhiteSpace($PropertyName)) {
        return
    }

    Ensure-ObjectProperty -Object $Manifest.latest -Name $PropertyName -DefaultValue ''
    $Manifest.latest.$PropertyName = if ([string]::IsNullOrWhiteSpace($AbsolutePath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $AbsolutePath) }
}

function Update-ResultBundleMetadata {
    param(
        $Result,
        [string]$VideoId,
        [string]$BundleRoot,
        [string]$ManifestPath,
        [string]$EventsPath
    )

    if ($null -eq $Result) {
        return
    }

    $Result.BundleRoot = [string]$BundleRoot
    $Result.BundleManifestPath = [string]$ManifestPath
    $Result.BundleEventsPath = [string]$EventsPath
}

function Write-BundleFetchArtifacts {
    param(
        [string]$VideoId,
        [string]$Title,
        [string]$SourceUrl,
        [string]$MetadataJson,
        [string]$MetadataFetchLog,
        [string]$SubtitleFetchLog,
        [string]$TranscriptText,
        [string]$TranscriptStatus,
        [string]$TranscriptSource,
        [string]$CaptionContent,
        [string]$CaptionExtension = '.vtt',
        [bool]$UseCookies = $false,
        [string]$ProfilePath = '',
        [bool]$MetadataUsedCookieFallback = $false,
        [bool]$SubtitleUsedCookieFallback = $false,
        [int]$SubtitleCandidateCount = 0
    )

    $bundleRoot = Ensure-VideoBundleStorage -VideoId $VideoId
    $runId = Get-BundleRunId -Prefix 'fetch'
    $runRelativeDir = Join-Path 'source\fetches' $runId
    $manifest = Get-VideoBundleManifest -VideoId $VideoId -Title $Title -SourceUrl $SourceUrl

    $metadataPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'metadata.json') -Text $MetadataJson
    $metadataFetchLogPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'metadata-fetch.log') -Text $MetadataFetchLog
    $subtitleFetchLogPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'subtitle-fetch.log') -Text $SubtitleFetchLog
    $transcriptPath = ''
    $captionPath = ''

    if (-not [string]::IsNullOrWhiteSpace([string]$TranscriptText)) {
        $transcriptPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'yt-transcript.txt') -Text (Normalize-TranscriptSourceText -Text $TranscriptText)
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$CaptionContent)) {
        $extension = ([string]$CaptionExtension).Trim()
        if ([string]::IsNullOrWhiteSpace($extension)) { $extension = '.vtt' }
        if (-not $extension.StartsWith('.')) { $extension = '.' + $extension }
        $captionPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir ('yt-captions' + $extension)) -Text $CaptionContent
    }

    foreach ($artifact in @(
        (New-BundleArtifactRecord -VideoId $VideoId -Path $metadataPath -Kind 'video-metadata' -Role 'source'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $metadataFetchLogPath -Kind 'yt-dlp-log' -Role 'metadata-fetch'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $subtitleFetchLogPath -Kind 'yt-dlp-log' -Role 'subtitle-fetch'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $transcriptPath -Kind 'yt-transcript' -Role 'source'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $captionPath -Kind 'yt-caption-file' -Role 'source')
    )) {
        Upsert-BundleArtifact -Manifest $manifest -Artifact $artifact
    }

    $manifest.title = [string]$Title
    $manifest.sourceUrl = [string]$SourceUrl
    $manifest.counters.fetches = [int]$manifest.counters.fetches + 1
    $manifest.latest.fetchRunId = $runId
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'metadataJson' -VideoId $VideoId -AbsolutePath $metadataPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'metadataFetchLog' -VideoId $VideoId -AbsolutePath $metadataFetchLogPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'subtitleFetchLog' -VideoId $VideoId -AbsolutePath $subtitleFetchLogPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'ytCaption' -VideoId $VideoId -AbsolutePath $captionPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'ytTranscript' -VideoId $VideoId -AbsolutePath $transcriptPath
    if (-not [string]::IsNullOrWhiteSpace($transcriptPath)) {
        Set-ManifestLatestPath -Manifest $manifest -PropertyName 'currentTranscript' -VideoId $VideoId -AbsolutePath $transcriptPath
        $manifest.latest.currentTranscriptMode = 'youtube'
        $manifest.latest.currentTimings = ''
    }

    Save-VideoBundleManifest -VideoId $VideoId -Manifest $manifest
    Add-BundleEvent -VideoId $VideoId -Type 'video_fetched' -RunId $runId -Status 'completed' -Inputs ([ordered]@{
        sourceUrl  = [string]$SourceUrl
        useCookies = [bool]$UseCookies
        profile    = if ([string]::IsNullOrWhiteSpace($ProfilePath)) { '' } else { (Split-Path -Leaf $ProfilePath) }
    }) -Outputs ([ordered]@{
        metadataJson     = [string]$manifest.latest.metadataJson
        metadataFetchLog = [string]$manifest.latest.metadataFetchLog
        subtitleFetchLog = [string]$manifest.latest.subtitleFetchLog
        ytCaption        = [string]$manifest.latest.ytCaption
        ytTranscript     = [string]$manifest.latest.ytTranscript
    }) -Details ([ordered]@{
        title                      = [string]$Title
        transcriptStatus           = [string]$TranscriptStatus
        transcriptSource           = [string]$TranscriptSource
        transcriptLength           = ([string]$TranscriptText).Length
        subtitleCandidateCount     = [int]$SubtitleCandidateCount
        metadataUsedCookieFallback = [bool]$MetadataUsedCookieFallback
        subtitleUsedCookieFallback = [bool]$SubtitleUsedCookieFallback
    })

    return [pscustomobject]@{
        BundleRoot       = $bundleRoot
        ManifestPath     = Get-VideoBundleManifestPath -VideoId $VideoId
        EventsPath       = Get-VideoBundleEventsPath -VideoId $VideoId
        FetchRunId       = $runId
        MetadataPath     = [string]$metadataPath
        MetadataFetchLog = [string]$metadataFetchLogPath
        SubtitleFetchLog = [string]$subtitleFetchLogPath
        TranscriptPath   = [string]$transcriptPath
        CaptionPath      = [string]$captionPath
    }
}

function Ensure-BundleAudioArtifact {
    param(
        [string]$VideoId,
        [string]$AudioPath
    )

    if ([string]::IsNullOrWhiteSpace($VideoId) -or [string]::IsNullOrWhiteSpace($AudioPath) -or -not (Test-Path -LiteralPath $AudioPath -PathType Leaf)) {
        return ''
    }

    $leaf = Split-Path -Leaf $AudioPath
    $targetPath = Copy-FileIntoBundle -VideoId $VideoId -SourcePath $AudioPath -RelativePath (Join-Path 'audio' $leaf)
    $manifest = Get-VideoBundleManifest -VideoId $VideoId
    Upsert-BundleArtifact -Manifest $manifest -Artifact (New-BundleArtifactRecord -VideoId $VideoId -Path $targetPath -Kind 'audio-file' -Role 'audio')
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'audio' -VideoId $VideoId -AbsolutePath $targetPath
    Save-VideoBundleManifest -VideoId $VideoId -Manifest $manifest
    return $targetPath
}

function Write-BundleAudioExtractionArtifacts {
    param(
        [string]$VideoId,
        [string]$AudioPath,
        [string]$ExtractionLog,
        [bool]$UsedCookieFallback = $false
    )

    $bundleAudioPath = Ensure-BundleAudioArtifact -VideoId $VideoId -AudioPath $AudioPath
    $runId = Get-BundleRunId -Prefix 'audio'
    $runRelativeDir = Join-Path 'audio\extractions' $runId
    $logPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'extraction.log') -Text $ExtractionLog
    $runInfoPath = Write-BundleJsonArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'run.json') -Value ([ordered]@{
        runId              = $runId
        createdAt          = Get-Iso8601Timestamp
        audioPath          = if ([string]::IsNullOrWhiteSpace($bundleAudioPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleAudioPath) }
        usedCookieFallback = [bool]$UsedCookieFallback
    })

    $manifest = Get-VideoBundleManifest -VideoId $VideoId
    foreach ($artifact in @(
        (New-BundleArtifactRecord -VideoId $VideoId -Path $logPath -Kind 'yt-dlp-log' -Role 'audio-extraction'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $runInfoPath -Kind 'run-metadata' -Role 'audio-extraction')
    )) {
        Upsert-BundleArtifact -Manifest $manifest -Artifact $artifact
    }

    $manifest.counters.audioExtractions = [int]$manifest.counters.audioExtractions + 1
    $manifest.latest.audioExtractionRunId = $runId
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'audioExtractionLog' -VideoId $VideoId -AbsolutePath $logPath
    Save-VideoBundleManifest -VideoId $VideoId -Manifest $manifest
    Add-BundleEvent -VideoId $VideoId -Type 'audio_extracted' -RunId $runId -Outputs ([ordered]@{
        audio = if ([string]::IsNullOrWhiteSpace($bundleAudioPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleAudioPath) }
        log   = (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $logPath)
        run   = (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $runInfoPath)
    }) -Details ([ordered]@{
        usedCookieFallback = [bool]$UsedCookieFallback
    })

    return [pscustomobject]@{
        RunId      = $runId
        AudioPath  = [string]$bundleAudioPath
        LogPath    = [string]$logPath
        RunInfoPath = [string]$runInfoPath
    }
}

function Write-BundleWhisperRunArtifacts {
    param(
        [string]$VideoId,
        [string]$RunId,
        [string]$Mode,
        [string]$Status,
        [string]$AudioPath,
        [string]$TranscriptPath,
        [string]$TimingPath,
        [string]$StdOutPath = '',
        [string]$StdErrPath = '',
        [string]$RunnerDescription = '',
        [object]$ExitCode = $null,
        [string]$OutputText = ''
    )

    $bundleAudioPath = Ensure-BundleAudioArtifact -VideoId $VideoId -AudioPath $AudioPath
    $runRelativeDir = Join-Path 'whisper\runs' $RunId
    $bundleTranscriptPath = if ([string]::IsNullOrWhiteSpace($TranscriptPath)) { '' } else { Copy-FileIntoBundle -VideoId $VideoId -SourcePath $TranscriptPath -RelativePath (Join-Path $runRelativeDir 'transcript.txt') }
    $bundleTimingPath = if ([string]::IsNullOrWhiteSpace($TimingPath)) { '' } else { Copy-FileIntoBundle -VideoId $VideoId -SourcePath $TimingPath -RelativePath (Join-Path $runRelativeDir 'timings.json') }
    $bundleStdOutPath = if ([string]::IsNullOrWhiteSpace($StdOutPath)) { '' } else { Copy-FileIntoBundle -VideoId $VideoId -SourcePath $StdOutPath -RelativePath (Join-Path $runRelativeDir 'stdout.log') }
    $bundleStdErrPath = if ([string]::IsNullOrWhiteSpace($StdErrPath)) { '' } else { Copy-FileIntoBundle -VideoId $VideoId -SourcePath $StdErrPath -RelativePath (Join-Path $runRelativeDir 'stderr.log') }
    $runInfoPath = Write-BundleJsonArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'run.json') -Value ([ordered]@{
        runId             = [string]$RunId
        createdAt         = Get-Iso8601Timestamp
        mode              = [string]$Mode
        status            = [string]$Status
        model             = $script:guiWhisperModel
        language          = $script:guiWhisperLanguage
        audioPath         = if ([string]::IsNullOrWhiteSpace($bundleAudioPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleAudioPath) }
        transcriptPath    = if ([string]::IsNullOrWhiteSpace($bundleTranscriptPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleTranscriptPath) }
        timingPath        = if ([string]::IsNullOrWhiteSpace($bundleTimingPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleTimingPath) }
        stdoutLog         = if ([string]::IsNullOrWhiteSpace($bundleStdOutPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdOutPath) }
        stderrLog         = if ([string]::IsNullOrWhiteSpace($bundleStdErrPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdErrPath) }
        runnerDescription = [string]$RunnerDescription
        exitCode          = if ($null -eq $ExitCode) { '' } else { [string]$ExitCode }
        outputSummary     = [string]$OutputText
    })

    $manifest = Get-VideoBundleManifest -VideoId $VideoId
    foreach ($artifact in @(
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleTranscriptPath -Kind 'whisper-transcript' -Role 'whisper'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleTimingPath -Kind 'whisper-timings' -Role 'whisper'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleStdOutPath -Kind 'whisper-log' -Role 'stdout'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleStdErrPath -Kind 'whisper-log' -Role 'stderr'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $runInfoPath -Kind 'run-metadata' -Role 'whisper')
    )) {
        Upsert-BundleArtifact -Manifest $manifest -Artifact $artifact
    }

    $manifest.counters.whisperRuns = [int]$manifest.counters.whisperRuns + 1
    $manifest.latest.whisperRunId = [string]$RunId
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'whisperTranscript' -VideoId $VideoId -AbsolutePath $bundleTranscriptPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'whisperTimings' -VideoId $VideoId -AbsolutePath $bundleTimingPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'whisperStdOutLog' -VideoId $VideoId -AbsolutePath $bundleStdOutPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'whisperStdErrLog' -VideoId $VideoId -AbsolutePath $bundleStdErrPath
    if (-not [string]::IsNullOrWhiteSpace($bundleTranscriptPath)) {
        Set-ManifestLatestPath -Manifest $manifest -PropertyName 'currentTranscript' -VideoId $VideoId -AbsolutePath $bundleTranscriptPath
        $manifest.latest.currentTranscriptMode = 'whisper'
    }
    if (-not [string]::IsNullOrWhiteSpace($bundleTimingPath)) {
        Set-ManifestLatestPath -Manifest $manifest -PropertyName 'currentTimings' -VideoId $VideoId -AbsolutePath $bundleTimingPath
    }
    Save-VideoBundleManifest -VideoId $VideoId -Manifest $manifest

    $eventType = switch ([string]$Status) {
        'completed' { if ([string]$Mode -eq 'cached') { 'whisper_loaded_from_cache' } else { 'whisper_completed' } }
        'failed' { 'whisper_failed' }
        default { 'whisper_recorded' }
    }
    Add-BundleEvent -VideoId $VideoId -Type $eventType -RunId $RunId -Status $Status -Outputs ([ordered]@{
        audio      = if ([string]::IsNullOrWhiteSpace($bundleAudioPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleAudioPath) }
        transcript = if ([string]::IsNullOrWhiteSpace($bundleTranscriptPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleTranscriptPath) }
        timings    = if ([string]::IsNullOrWhiteSpace($bundleTimingPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleTimingPath) }
        stdoutLog  = if ([string]::IsNullOrWhiteSpace($bundleStdOutPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdOutPath) }
        stderrLog  = if ([string]::IsNullOrWhiteSpace($bundleStdErrPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdErrPath) }
        run        = (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $runInfoPath)
    }) -Details ([ordered]@{
        mode            = [string]$Mode
        model           = $script:guiWhisperModel
        language        = $script:guiWhisperLanguage
        exitCode        = if ($null -eq $ExitCode) { '' } else { [string]$ExitCode }
        transcriptChars = (Read-TextFile -Path $bundleTranscriptPath).Length
    })

    return [pscustomobject]@{
        AudioPath      = [string]$bundleAudioPath
        TranscriptPath = [string]$bundleTranscriptPath
        TimingPath     = [string]$bundleTimingPath
        StdOutPath     = [string]$bundleStdOutPath
        StdErrPath     = [string]$bundleStdErrPath
        RunInfoPath    = [string]$runInfoPath
    }
}

function Get-BundlePresetDefinition {
    param([string]$PresetId)

    if ([string]::IsNullOrWhiteSpace($PresetId) -or $null -eq $script:gemmaPresets) {
        return $null
    }

    return ($script:gemmaPresets | Where-Object { [string]$_.id -eq [string]$PresetId } | Select-Object -First 1)
}

function Write-BundleLlmRunArtifacts {
    param(
        [string]$VideoId,
        [string]$RunId,
        [string]$Mode,
        [string]$Status,
        [string]$Model,
        [string]$PresetId,
        [string]$TranscriptHash,
        [string]$InputFilePath = '',
        [string]$TranscriptText = '',
        [string]$ResultFilePath = '',
        [string]$StdOutPath = '',
        [string]$StdErrPath = '',
        [string]$ActivityText = '',
        [object]$ExitCode = $null
    )

    $runRelativeDir = Join-Path 'llm\runs' $RunId
    $preset = Get-BundlePresetDefinition -PresetId $PresetId
    $bundleInputTranscriptPath = if (-not [string]::IsNullOrWhiteSpace($InputFilePath) -and (Test-Path -LiteralPath $InputFilePath -PathType Leaf)) {
        Copy-FileIntoBundle -VideoId $VideoId -SourcePath $InputFilePath -RelativePath (Join-Path $runRelativeDir 'input-transcript.txt')
    }
    elseif (-not [string]::IsNullOrWhiteSpace($TranscriptText)) {
        Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'input-transcript.txt') -Text (Normalize-TranscriptSourceText -Text $TranscriptText)
    }
    else {
        ''
    }

    $bundleResultPath = if (-not [string]::IsNullOrWhiteSpace($ResultFilePath) -and (Test-Path -LiteralPath $ResultFilePath -PathType Leaf)) {
        Copy-FileIntoBundle -VideoId $VideoId -SourcePath $ResultFilePath -RelativePath (Join-Path $runRelativeDir 'output.json')
    }
    else {
        ''
    }
    $bundleStdOutPath = if ([string]::IsNullOrWhiteSpace($StdOutPath)) { '' } else { Copy-FileIntoBundle -VideoId $VideoId -SourcePath $StdOutPath -RelativePath (Join-Path $runRelativeDir 'stdout.log') }
    $bundleStdErrPath = if ([string]::IsNullOrWhiteSpace($StdErrPath)) { '' } else { Copy-FileIntoBundle -VideoId $VideoId -SourcePath $StdErrPath -RelativePath (Join-Path $runRelativeDir 'stderr.log') }
    $promptPath = ''
    $displayPath = ''
    $presetInfoPath = ''
    $payload = $null

    if (-not [string]::IsNullOrWhiteSpace($bundleResultPath)) {
        try {
            $payload = Get-Content -LiteralPath $bundleResultPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            $payload = $null
        }
    }

    if ($null -ne $preset) {
        $presetInfoPath = Write-BundleJsonArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'preset.json') -Value ([ordered]@{
            id           = [string]$preset.id
            label        = [string]$preset.label
            description  = [string]$preset.description
            instructions = [string]$preset.instructions
        })
    }

    if ($null -ne $payload -and -not [string]::IsNullOrWhiteSpace([string]$payload.promptText)) {
        $promptPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'prompt.txt') -Text ([string]$payload.promptText)
    }
    if ($null -ne $payload -and -not [string]::IsNullOrWhiteSpace([string]$payload.displayText)) {
        $displayPath = Write-BundleTextArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'display.txt') -Text ([string]$payload.displayText)
    }

    $runInfoPath = Write-BundleJsonArtifact -VideoId $VideoId -RelativePath (Join-Path $runRelativeDir 'run.json') -Value ([ordered]@{
        runId             = [string]$RunId
        createdAt         = Get-Iso8601Timestamp
        mode              = [string]$Mode
        status            = [string]$Status
        model             = [string]$Model
        presetId          = [string]$PresetId
        transcriptHash    = [string]$TranscriptHash
        inputTranscript   = if ([string]::IsNullOrWhiteSpace($bundleInputTranscriptPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleInputTranscriptPath) }
        presetPath        = if ([string]::IsNullOrWhiteSpace($presetInfoPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $presetInfoPath) }
        promptPath        = if ([string]::IsNullOrWhiteSpace($promptPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $promptPath) }
        resultPath        = if ([string]::IsNullOrWhiteSpace($bundleResultPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleResultPath) }
        displayPath       = if ([string]::IsNullOrWhiteSpace($displayPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $displayPath) }
        stdoutLog         = if ([string]::IsNullOrWhiteSpace($bundleStdOutPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdOutPath) }
        stderrLog         = if ([string]::IsNullOrWhiteSpace($bundleStdErrPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdErrPath) }
        activityText      = [string]$ActivityText
        exitCode          = if ($null -eq $ExitCode) { '' } else { [string]$ExitCode }
        ollama            = if ($null -eq $payload) { $null } else { $payload.ollama }
    })

    $manifest = Get-VideoBundleManifest -VideoId $VideoId
    foreach ($artifact in @(
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleInputTranscriptPath -Kind 'llm-input-transcript' -Role 'llm'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $presetInfoPath -Kind 'llm-preset' -Role 'llm'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $promptPath -Kind 'llm-prompt' -Role 'llm'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleResultPath -Kind 'llm-output-json' -Role 'llm'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $displayPath -Kind 'llm-display-text' -Role 'llm'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleStdOutPath -Kind 'llm-log' -Role 'stdout'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $bundleStdErrPath -Kind 'llm-log' -Role 'stderr'),
        (New-BundleArtifactRecord -VideoId $VideoId -Path $runInfoPath -Kind 'run-metadata' -Role 'llm')
    )) {
        Upsert-BundleArtifact -Manifest $manifest -Artifact $artifact
    }

    $manifest.counters.llmRuns = [int]$manifest.counters.llmRuns + 1
    $manifest.latest.llmRunId = [string]$RunId
    $manifest.latest.llmModel = [string]$Model
    $manifest.latest.llmPresetId = [string]$PresetId
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'llmInputTranscript' -VideoId $VideoId -AbsolutePath $bundleInputTranscriptPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'llmPrompt' -VideoId $VideoId -AbsolutePath $promptPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'llmResult' -VideoId $VideoId -AbsolutePath $bundleResultPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'llmDisplay' -VideoId $VideoId -AbsolutePath $displayPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'llmStdOutLog' -VideoId $VideoId -AbsolutePath $bundleStdOutPath
    Set-ManifestLatestPath -Manifest $manifest -PropertyName 'llmStdErrLog' -VideoId $VideoId -AbsolutePath $bundleStdErrPath
    Save-VideoBundleManifest -VideoId $VideoId -Manifest $manifest

    $eventType = switch ([string]$Status) {
        'completed' { if ([string]$Mode -eq 'cached') { 'llm_loaded_from_cache' } else { 'llm_completed' } }
        'failed' { 'llm_failed' }
        default { 'llm_recorded' }
    }
    Add-BundleEvent -VideoId $VideoId -Type $eventType -RunId $RunId -Status $Status -Outputs ([ordered]@{
        inputTranscript = if ([string]::IsNullOrWhiteSpace($bundleInputTranscriptPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleInputTranscriptPath) }
        preset          = if ([string]::IsNullOrWhiteSpace($presetInfoPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $presetInfoPath) }
        prompt          = if ([string]::IsNullOrWhiteSpace($promptPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $promptPath) }
        result          = if ([string]::IsNullOrWhiteSpace($bundleResultPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleResultPath) }
        display         = if ([string]::IsNullOrWhiteSpace($displayPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $displayPath) }
        stdoutLog       = if ([string]::IsNullOrWhiteSpace($bundleStdOutPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdOutPath) }
        stderrLog       = if ([string]::IsNullOrWhiteSpace($bundleStdErrPath)) { '' } else { (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $bundleStdErrPath) }
        run             = (Get-RelativePathFromBase -BasePath (Get-VideoBundleDirectory -VideoId $VideoId) -TargetPath $runInfoPath)
    }) -Details ([ordered]@{
        mode           = [string]$Mode
        model          = [string]$Model
        presetId       = [string]$PresetId
        transcriptHash = [string]$TranscriptHash
        exitCode       = if ($null -eq $ExitCode) { '' } else { [string]$ExitCode }
    })

    return [pscustomobject]@{
        InputTranscriptPath = [string]$bundleInputTranscriptPath
        PromptPath          = [string]$promptPath
        ResultPath          = [string]$bundleResultPath
        DisplayPath         = [string]$displayPath
        StdOutPath          = [string]$bundleStdOutPath
        StdErrPath          = [string]$bundleStdErrPath
        RunInfoPath         = [string]$runInfoPath
        PresetPath          = [string]$presetInfoPath
    }
}
