#Requires -Version 5.0
param(
    [switch]$TreatYtDlpErrorsAsWarnings = $false  # If true, yt-dlp non-zero exit codes will be warnings instead of errors
)

<#
.SYNOPSIS
    yt-dlp Helper Script for YouTube Downloads with improved error handling

.DESCRIPTION
    This script provides a menu-driven interface for downloading YouTube videos and playlists using yt-dlp.
    It includes automatic yt-dlp updates, FFmpeg setup, and robust error handling for unavailable videos.

.PARAMETER TreatYtDlpErrorsAsWarnings
    When specified, yt-dlp exit codes indicating failures (like unavailable videos) will be treated as warnings
    instead of terminating errors. This allows the script to continue even when some videos in a playlist
    are unavailable due to copyright claims, account termination, or regional restrictions.

.EXAMPLE
    .\yt-dlp-helper.ps1
    Run the script normally - yt-dlp errors will cause the script to terminate

.EXAMPLE
    .\yt-dlp-helper.ps1 -TreatYtDlpErrorsAsWarnings
    Run the script with graceful error handling - unavailable videos will generate warnings but won't stop the script

.NOTES
    The TreatYtDlpErrorsAsWarnings parameter is especially useful for large playlists where some videos
    may become unavailable over time. Without this flag, encountering a single unavailable video would
    stop the entire download process.
#>

# --- START Console Encoding ---
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {
    Write-Verbose "Could not set UTF-8 console encoding: $($_.Exception.Message)"
}
# --- END Console Encoding ---

# --- START Preference Handling ---
# Global preferences file (in script directory)
$globalPrefsFile = Join-Path -Path $PSScriptRoot -ChildPath "user_preferences.json"

# Script-level state
$script:currentYtDlpVersion = $null # Stores the version tag of the current yt-dlp.exe
$script:lastChoice = $null
$script:lastPlaylistIndex = $null
$script:lastPlaylistId = $null
$script:downloadRootPath = $null
$script:perRootPrefsFile = $null
$script:treatYtDlpErrorsAsWarningsPref = $false

# Function to save preferences (global only)
function Save-UserPreferences {
    [CmdletBinding()]
    param(
        [string]$menuChoice = $script:lastChoice,
        [string]$playlistIndex = $script:lastPlaylistIndex,
        [string]$playlistId = $script:lastPlaylistId,
        [string]$ytDlpVersion = $script:currentYtDlpVersion,
        [string]$downloadRoot = $script:downloadRootPath
    )

    # Update script variables
    $script:lastChoice = $menuChoice
    $script:lastPlaylistIndex = $playlistIndex
    $script:lastPlaylistId = $playlistId
    $script:currentYtDlpVersion = $ytDlpVersion
    if ($downloadRoot) { $script:downloadRootPath = $downloadRoot }

    # Ensure we have a download root (default to ./Downloads)
    if (-not $script:downloadRootPath) {
        $script:downloadRootPath = Join-Path $PSScriptRoot "Downloads"
    }

    # Build and save GLOBAL prefs only
    $globalPrefs = @{
        lastMenuChoice       = $script:lastChoice
        lastPlaylistIndex    = $script:lastPlaylistIndex
        lastPlaylistId       = $script:lastPlaylistId
        currentYtDlpVersion  = $script:currentYtDlpVersion
        lastDownloadRootPath = $script:downloadRootPath
        treatYtDlpErrorsAsWarningsPreferred = if ($TreatYtDlpErrorsAsWarnings) { $true } else { $script:treatYtDlpErrorsAsWarningsPref }
    }
    try {
        $globalPrefs | ConvertTo-Json -Depth 5 | Set-Content -Path $globalPrefsFile -Encoding UTF8 -Force
    } catch {
        Write-Warning "Could not save global preferences to '$globalPrefsFile': $($_.Exception.Message)"
    }
}

# Function to load preferences (global only)
function Load-UserPreferences {
    # Load global
    if (Test-Path $globalPrefsFile) {
        try {
            $global = Get-Content -Path $globalPrefsFile -Raw | ConvertFrom-Json -ErrorAction Stop
            if ($global.PSObject.Properties.Name -contains 'lastMenuChoice') { $script:lastChoice = $global.lastMenuChoice }
            if ($global.PSObject.Properties.Name -contains 'lastPlaylistIndex') { $script:lastPlaylistIndex = $global.lastPlaylistIndex }
            if ($global.PSObject.Properties.Name -contains 'lastPlaylistId') { $script:lastPlaylistId = $global.lastPlaylistId }
            if ($global.PSObject.Properties.Name -contains 'currentYtDlpVersion') { $script:currentYtDlpVersion = $global.currentYtDlpVersion }
            if ($global.PSObject.Properties.Name -contains 'lastDownloadRootPath') { $script:downloadRootPath = $global.lastDownloadRootPath }
            if ($global.PSObject.Properties.Name -contains 'treatYtDlpErrorsAsWarningsPreferred') { $script:treatYtDlpErrorsAsWarningsPref = [bool]$global.treatYtDlpErrorsAsWarningsPreferred }
        } catch {
            Write-Warning "Could not load or parse global preferences from '$globalPrefsFile': $($_.Exception.Message). Using defaults."
        }
    } else {
        Write-Host "Global preference file '$globalPrefsFile' not found. Using defaults." -ForegroundColor DarkGray
    }

    # Ensure download root default
    if (-not $script:downloadRootPath) {
        $script:downloadRootPath = Join-Path $PSScriptRoot "Downloads"
    }

    Write-Host "Loaded preferences: YT-DLP Ver: $($script:currentYtDlpVersion), Last Choice: $($script:lastChoice), Last Playlist: $($script:lastPlaylistIndex), Download Root: $($script:downloadRootPath)" -ForegroundColor DarkGray
}
# --- END Preference Handling ---


# --- START YT-DLP Nightly Update Check ---

# Function to get yt-dlp.exe version
function Get-LocalYtDlpVersion {
    param([string]$ExePath)
    if (-not (Test-Path $ExePath -PathType Leaf)) {
        return $null
    }
    try {
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $ExePath
        $processInfo.Arguments = '--version'
        $processInfo.UseShellExecute = $false
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.CreateNoWindow = $true
        
        $process = [System.Diagnostics.Process]::Start($processInfo)
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $exitCode = $process.ExitCode

        # Match standard YYYY.MM.DD or nightly YYYY.MM.DD.timestamp
        if ($exitCode -eq 0 -and $stdout -match '^\d{4}\.\d{2}\.\d{2}(?:\.\d{6})?') {
            return $matches[0]
        } else {
            Write-Warning "Could not get version from '$ExePath'. Exit Code: $exitCode Stdout: '$stdout' Stderr: '$stderr'"
            return $null
        }
    } catch {
        Write-Warning "Error running '$ExePath --version': $($_.Exception.Message)"
        return $null
    }
}

# Function to check for nightly updates and download if necessary
function Check-YtDlpNightlyUpdate {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExpectedExePath, # Path where yt-dlp.exe should be

        [Parameter(Mandatory=$true)]
        [ref]$ScriptVersionRef, # Reference to update $script:currentYtDlpVersion
        
        [Parameter(Mandatory=$true)]
        [ref]$ResolvedExePathRef # Outputs the final path if successful
    )
    
    Write-Host "--- Checking YT-DLP Nightly Build ---" -ForegroundColor Magenta
    $ResolvedExePathRef.Value = $null # Reset output path
    $localVersion = $null
    $exeExists = Test-Path $ExpectedExePath -PathType Leaf
    $nightlyApiUrl = "https://api.github.com/repos/yt-dlp/yt-dlp-nightly-builds/releases/latest"

    # 1. Determine Local Version
    if ($exeExists) {
        $localVersion = Get-LocalYtDlpVersion -ExePath $ExpectedExePath
        if (-not $localVersion -and $ScriptVersionRef.Value) {
             Write-Host "Could not run '$ExpectedExePath --version'. Using stored version: $($ScriptVersionRef.Value)" -ForegroundColor Yellow
             $localVersion = $ScriptVersionRef.Value
        } elseif ($localVersion -and $localVersion -ne $ScriptVersionRef.Value) {
             Write-Host "Executable version ($localVersion) differs from stored version ($($ScriptVersionRef.Value)). Updating stored version." -ForegroundColor Yellow
             $ScriptVersionRef.Value = $localVersion
             Save-UserPreferences -ytDlpVersion $localVersion
        }
        $ResolvedExePathRef.Value = $ExpectedExePath # Mark as found locally for now
    } else {
         Write-Host "yt-dlp.exe not found at '$ExpectedExePath'." -ForegroundColor Yellow
    }

    # 2. Check GitHub for Latest Nightly Version
    Write-Host "Checking GitHub for latest nightly release... ($nightlyApiUrl)" -ForegroundColor Cyan
    $latestRelease = $null
    try {
        $latestRelease = Invoke-RestMethod -Uri $nightlyApiUrl -UseBasicParsing -TimeoutSec 15
    } catch {
        Write-Warning "Failed to fetch latest nightly release info from GitHub: $($_.Exception.Message)"
        if ($exeExists) {
             Write-Host "Proceeding with existing local yt-dlp.exe (version: $(if ($localVersion) { $localVersion } else { 'unknown' }))" -ForegroundColor Yellow
             return $true # Indicate we can proceed (with the local version)
        } else {
            Write-Error "Cannot check for updates and local yt-dlp.exe is missing."
            return $false # Indicate we cannot proceed
        }
    }

    $latestVersion = $latestRelease.tag_name
    if ($latestVersion -like 'v*') { $latestVersion = $latestVersion.Substring(1) }
    Write-Host "Latest nightly version available: $latestVersion" -ForegroundColor Green

    # 3. Compare Versions and Decide if Update Needed
    $updateNeeded = $false
    if (-not $exeExists) {
        Write-Host "Local yt-dlp.exe missing. Update required." -ForegroundColor Yellow
        $updateNeeded = $true
    } elseif (-not $localVersion) {
         Write-Host "Could not determine local version. Assuming update is needed." -ForegroundColor Yellow
         $updateNeeded = $true
    } elseif ($latestVersion -gt $localVersion) { # String comparison works for YYYY.MM.DD[.timestamp]
        Write-Host "Newer nightly version available ($latestVersion > $localVersion). Update recommended." -ForegroundColor Yellow
        $updateNeeded = $true
    } else {
        Write-Host "You have the latest nightly version ($localVersion)." -ForegroundColor Green
        if ($localVersion -and $ScriptVersionRef.Value -ne $localVersion) {
             $ScriptVersionRef.Value = $localVersion
             Save-UserPreferences -ytDlpVersion $localVersion
        }
        return $true # Indicate we can proceed (with the up-to-date local version)
    }

    # 4. Download Update if Needed
    if ($updateNeeded) {
        $assetName = "yt-dlp.exe"
        $asset = $latestRelease.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1

        if (-not $asset) {
            Write-Error "No asset named '$assetName' found in the latest nightly release."
            if ($exeExists) {
                 Write-Host "Proceeding with existing local yt-dlp.exe (version: $(if ($localVersion) { $localVersion } else { 'unknown' }))" -ForegroundColor Yellow
                 return $true
            } else {
                return $false
            }
        }

        Write-Host "Downloading $($asset.name) (Version: $latestVersion)... -> $ExpectedExePath" -ForegroundColor Cyan
        $downloadUrl = $asset.browser_download_url
        try {
            if (Test-Path $ExpectedExePath) {
                try {
                    $fileStream = [System.IO.File]::Open($ExpectedExePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                    $fileStream.Close(); $fileStream.Dispose()
                } catch [System.IO.IOException] {
                    Write-Warning "Could not get write access to '$ExpectedExePath'. It might be in use. Skipping update."
                    if ($exeExists) { return $true }
                    else { return $false }
                } catch { Write-Warning "Error checking file lock: $($_.Exception.Message)" }
            }
            
            Invoke-WebRequest -Uri $downloadUrl -OutFile $ExpectedExePath -UseBasicParsing -ErrorAction Stop
            Write-Host "Download complete."-ForegroundColor Green

            $newVersion = Get-LocalYtDlpVersion -ExePath $ExpectedExePath
            if ($newVersion) {
                $ScriptVersionRef.Value = $newVersion # Should match latestVersion ideally
                $ResolvedExePathRef.Value = $ExpectedExePath
                Save-UserPreferences -ytDlpVersion $newVersion
                Write-Host "Successfully updated yt-dlp to nightly version $newVersion." -ForegroundColor Green
                return $true
            } else {
                 Write-Error "Downloaded file '$ExpectedExePath' but failed to execute or get version."
                 if (Test-Path $ExpectedExePath) { Remove-Item $ExpectedExePath -ErrorAction SilentlyContinue }
                 return $false
            }
        } catch {
            Write-Error "Failed to download '$($asset.name)' from '$downloadUrl': $($_.Exception.Message)"
            if ($exeExists) {
                Write-Host "Proceeding with existing local yt-dlp.exe." -ForegroundColor Yellow
                return $true
            } else {
                return $false
            }
        }
    }
    
    # Fallback
    if ($exeExists) { return $true } else { return $false }
}

# --- END YT-DLP Nightly Update Check ---

# --- Script Main Execution Start ---
Load-UserPreferences # Load preferences first

Write-Host "Starting yt-dlp Helper Script..."

# Prompt for or confirm Download Root (persist per-root and globally)
$defaultRoot = if ($script:downloadRootPath) { $script:downloadRootPath } else { Join-Path $PSScriptRoot "Downloads" }
$treatWarningsDefault = if ($TreatYtDlpErrorsAsWarnings) { $true } else { $script:treatYtDlpErrorsAsWarningsPref }
$enteredRoot = Read-Host "Enter download root directory (press Enter to use last/default): [$defaultRoot]"
if ([string]::IsNullOrWhiteSpace($enteredRoot)) { $enteredRoot = $defaultRoot }

# Validate and resolve the entered root
$selectedRoot = $enteredRoot
$fallbackUsed = $false
$fallbackReason = $null

# Check invalid characters
$invalidChars = [System.IO.Path]::GetInvalidPathChars()
if ($selectedRoot.IndexOfAny($invalidChars) -ge 0) {
    $fallbackUsed = $true
    $fallbackReason = "contains invalid path characters"
    $selectedRoot = $defaultRoot
}

# Resolve relative paths against script root
if (-not $fallbackUsed) {
    if (-not [System.IO.Path]::IsPathRooted($selectedRoot)) {
        $selectedRoot = Join-Path -Path $PSScriptRoot -ChildPath $selectedRoot
    }
}

# Ensure directory exists and is accessible; if creation fails, fallback
if (-not $fallbackUsed) {
    if (-not (Test-Path $selectedRoot -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $selectedRoot -Force -ErrorAction Stop | Out-Null
            Write-Host "Created download root: $selectedRoot"
        } catch {
            $fallbackUsed = $true
            $fallbackReason = "failed to create directory ($($_.Exception.Message))"
            $selectedRoot = $defaultRoot
        }
    }
}

# Final check: ensure the selectedRoot is a directory
if (-not $fallbackUsed) {
    if (-not (Test-Path $selectedRoot -PathType Container)) {
        $fallbackUsed = $true
        $fallbackReason = "path is not a directory or not accessible"
        $selectedRoot = $defaultRoot
    }
}

if ($fallbackUsed -and ($selectedRoot -ne $enteredRoot)) {
    Write-Warning "The entered download root '$enteredRoot' is invalid: $fallbackReason. Using default/previous root: '$selectedRoot'."
}

$script:downloadRootPath = $selectedRoot

# If we fell back to the default/previous root and it doesn't exist yet, ensure it is created
if ($fallbackUsed -and -not (Test-Path $script:downloadRootPath -PathType Container)) {
    try {
        New-Item -ItemType Directory -Path $script:downloadRootPath -Force -ErrorAction Stop | Out-Null
        Write-Host "Created download root: $($script:downloadRootPath)"
    } catch {
        Write-Error "Failed to create download root '$($script:downloadRootPath)': $($_.Exception.Message)"
        exit 1
    }
}

# Always confirm the final chosen download root
Write-Host "Using download root: $($script:downloadRootPath)" -ForegroundColor Cyan

# If user didn't explicitly pass the switch, use saved preference
if (-not $TreatYtDlpErrorsAsWarnings -and $treatWarningsDefault) {
    Write-Host "Using saved preference: Treat yt-dlp errors as warnings." -ForegroundColor DarkYellow
    $script:useWarningsForYtDlp = $true
} else {
    $script:useWarningsForYtDlp = [bool]$TreatYtDlpErrorsAsWarnings
}

# Save immediately so per-root prefs file is established
Save-UserPreferences -downloadRoot $script:downloadRootPath

$expectedYtDlpExePath = Join-Path $PSScriptRoot "yt-dlp.exe"
$ytDlpExePath = $null # Will be set by the update check

# Run the update check
$updateCheckSuccess = Check-YtDlpNightlyUpdate -ExpectedExePath $expectedYtDlpExePath -ScriptVersionRef ([ref]$script:currentYtDlpVersion) -ResolvedExePathRef ([ref]$ytDlpExePath)

if (-not $updateCheckSuccess -or -not $ytDlpExePath) {
    Write-Error "Halting script: Failed to find or update yt-dlp.exe."
    exit 1
}

Write-Host "==> Using YT-DLP Version: $($script:currentYtDlpVersion) | Path: $ytDlpExePath <==" -ForegroundColor White -BackgroundColor DarkBlue

# --- FFmpeg Setup Logic ---

# Define paths and constants
$ffmpegBuildSubDir = "ffmpeg_yt-dlp"
$ffmpegDownloadUrl = "https://github.com/yt-dlp/FFmpeg-Builds/releases/latest/download/ffmpeg-master-latest-win64-gpl.zip"
$ffmpegInstallPath = Join-Path -Path $PSScriptRoot -ChildPath $ffmpegBuildSubDir
$ffmpegLocationArgument = "" # Initialize
$absoluteFfmpegBinPath = "" # Will store the absolute path if found

Write-Host "Checking for local FFmpeg build in $ffmpegInstallPath..." -ForegroundColor Cyan

# Function to find the bin path (used if ffmpeg exists or after download)
# Returns the ABSOLUTE path to the bin dir if found, otherwise null
function Get-AbsoluteFfmpegBinPath { 
    param ([string]$baseInstallPath)
    $foundBinPath = $null
    # Look for a subdirectory containing a 'bin' folder
    $extractedFolder = Get-ChildItem -Path $baseInstallPath -Directory -ErrorAction SilentlyContinue | Where-Object { Test-Path (Join-Path -Path $_.FullName -ChildPath "bin") } | Select-Object -First 1

    if ($extractedFolder) {
        $foundBinPath = Join-Path -Path $extractedFolder.FullName -ChildPath "bin"
        Write-Host "  Found FFmpeg bin directory (Absolute): $foundBinPath" -ForegroundColor Green
    } else {
        # Check if the base path itself contains bin (less likely but possible)
        if (Test-Path (Join-Path -Path $baseInstallPath -ChildPath "bin") -ErrorAction SilentlyContinue) {
            $foundBinPath = Join-Path -Path $baseInstallPath -ChildPath "bin"
            Write-Host "  Found FFmpeg bin directory directly under (Absolute): $foundBinPath" -ForegroundColor Green
        } else {
            Write-Warning "  Could not locate the 'bin' directory within $baseInstallPath or its direct subdirectories."
        }
    }
    return $foundBinPath
}

# Try to find existing ffmpeg bin path (Absolute)
$absoluteFfmpegBinPath = Get-AbsoluteFfmpegBinPath -baseInstallPath $ffmpegInstallPath

# If not found, attempt download and extraction
if (-not $absoluteFfmpegBinPath) {
    Write-Host "Local FFmpeg not found. Attempting to download..." -ForegroundColor Yellow
    $zipPath = Join-Path -Path $ffmpegInstallPath -ChildPath "ffmpeg_download.zip"

    # Ensure install directory exists
    if (-not (Test-Path $ffmpegInstallPath -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $ffmpegInstallPath -Force -ErrorAction Stop | Out-Null
            Write-Host "  Created directory: $ffmpegInstallPath"
        } catch {
            Write-Error "  Failed to create directory $ffmpegInstallPath. Check permissions. Error: $($_.Exception.Message)"
            # Cannot proceed without install dir
            exit 1
        }
    }

    try {
        Write-Host "  Downloading from $ffmpegDownloadUrl..."
        Invoke-WebRequest -Uri $ffmpegDownloadUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        Write-Host "  Download complete." -ForegroundColor Green

        Write-Host "  Extracting FFmpeg to $ffmpegInstallPath..."
        Expand-Archive -Path $zipPath -DestinationPath $ffmpegInstallPath -Force -ErrorAction Stop

        Write-Host "  Removing temporary download file: $zipPath"
        Remove-Item $zipPath -Force

        # Try finding the bin path again after extraction (Absolute)
        $absoluteFfmpegBinPath = Get-AbsoluteFfmpegBinPath -baseInstallPath $ffmpegInstallPath

        if ($absoluteFfmpegBinPath) {
            Write-Host "  FFmpeg download and extraction successful." -ForegroundColor Green
        } else {
            Write-Error "  FFmpeg downloaded but failed to find bin directory after extraction."
        }

    } catch {
        Write-Error "  An error occurred during FFmpeg download/extraction: $($_.Exception.Message)"
        if (Test-Path $zipPath) {
            Write-Warning "  Cleaning up partially downloaded file: $zipPath"
            Remove-Item $zipPath -Force
        }
        Write-Warning "Proceeding without automatically configuring FFmpeg location for yt-dlp."
        $absoluteFfmpegBinPath = $null # Ensure it's null/empty if download failed
    }
}

# Set the argument for yt-dlp if an absolute path was found
if ($absoluteFfmpegBinPath) {
    # Use the ABSOLUTE path for the argument, enclosed in quotes for the command line
    $ffmpegLocationArgument = "--ffmpeg-location ""$absoluteFfmpegBinPath"""
} else {
    Write-Warning "FFmpeg location could not be determined. yt-dlp will try to find it in PATH or alongside yt-dlp.exe."
    $ffmpegLocationArgument = "" # Ensure it's empty
}

# --- End FFmpeg Setup Logic ---

# --- Authentication Setup --- 
# Hardcoding authentication using the full profile PATH
$firefoxProfilePath = "$env:USERPROFILE\AppData\Roaming\Mozilla\Firefox\Profiles\fw0m6kre.default-nightly" # !!! UPDATED WITH USER PATH !!!

# Determine whether authentication via Firefox cookies is available
$script:authAvailable = $false
$authType = $null
$authValue = $null

$authStatusMessage = $null
if ([string]::IsNullOrWhiteSpace($firefoxProfilePath)) {
    $authStatusMessage = 'profile path is empty or not set'
} elseif (-not (Test-Path $firefoxProfilePath -PathType Container)) {
    $authStatusMessage = "profile path does not exist: $firefoxProfilePath"
} else {
    try {
        # Check accessibility and that the directory is not empty
        $items = Get-ChildItem -Path $firefoxProfilePath -Force -ErrorAction Stop
        if ($items.Count -gt 0) {
            $profileFolderName = Split-Path -Leaf $firefoxProfilePath
            $authType = 'cookies_browser'
            $authValue = "firefox:$profileFolderName"
            $script:authAvailable = $true
            Write-Host "Using authentication: $authType with value $authValue" -ForegroundColor Yellow
        } else {
            $authStatusMessage = "profile path exists but is empty: $firefoxProfilePath"
        }
    } catch {
        $authStatusMessage = "profile path is not accessible: $($_.Exception.Message)"
    }
}

if (-not $script:authAvailable) {
    Write-Warning "No Firefox profile available for YouTube authentication ($authStatusMessage). The script will use public-only features."
}
# --- End Authentication Setup ---

# --- Download Functions ---
# Archive and outputs now live under the download root

# Core flags shared by all commands (archive is injected per-call for concurrency safety)
$commonFlagsCore = @(
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



# Function to download the best quality single video with metadata
function Download-BestVideo {
    param(
        [string]$ffmpegLocationArg,
        [string[]]$commonArgs,
        [string]$authTypeValue, # Pass type
        [string]$authPathValue,   # Pass value
        [string]$ytDlpPath,
        [switch]$TreatErrorsAsWarnings = $false
    )
    $videoUrl = Read-Host "Please enter the YouTube video URL"
    if (-not $videoUrl) { Write-Warning "No URL provided. Exiting."; return }
    
    $singlesRoot = Join-Path $script:downloadRootPath "Singles"
    if (-not (Test-Path $singlesRoot -PathType Container)) {
        try { New-Item -ItemType Directory -Path $singlesRoot -Force -ErrorAction Stop | Out-Null } catch { Write-Error "Failed to create directory $singlesRoot. Error: $($_.Exception.Message)"; return }
    }

    $archivePath = Join-Path $singlesRoot "download_archive.txt"

    # Construct the output template (Singles root)
    $outputTemplate = Join-Path $singlesRoot "%(title)s [%(id)s].%(ext)s"
    
    # Build Argument List
    $ArgumentList = @()
    if ($ffmpegLocationArg) { $ArgumentList += $ffmpegLocationArg.Split(' ', 2) }
    if ($commonArgs) { $ArgumentList += $commonArgs }
    $ArgumentList += '--download-archive', "`"$archivePath`""
    if ($authTypeValue -eq 'cookies_browser' -and $authPathValue) {
        $ArgumentList += '--cookies-from-browser'
        $ArgumentList += $authPathValue
    }
    $ArgumentList += '-o', $outputTemplate
    
    $ArgumentList += $videoUrl

    Write-Host ("Executing yt-dlp with arguments: {0}" -f ($ArgumentList -join ' ')) -ForegroundColor Yellow
    try {
        & $ytDlpPath $ArgumentList
        $singleExitCode = $LASTEXITCODE
        if ($singleExitCode -ne 0) { 
            $errorMessage = "yt-dlp finished with exit code $singleExitCode (video may have been unavailable or failed to download)"
            if ($TreatErrorsAsWarnings) {
                Write-Warning $errorMessage
                Write-Host "Single video download completed with warnings. Check the output above for details." -ForegroundColor Yellow
            } else {
                Write-Error $errorMessage
            }
        } else { 
            Write-Host "Download completed successfully." -ForegroundColor Green 
        }
    } catch { 
        $launchError = "An error occurred launching yt-dlp: $($_.Exception.Message)"
        if ($TreatErrorsAsWarnings) {
            Write-Warning $launchError
        } else {
            Write-Error $launchError
        }
    }
}

# Function to download a playlist with metadata into a named subfolder
function Download-Playlist {
    param(
        [string]$ffmpegLocationArg,
        [string[]]$commonArgs,
        [string]$authTypeValue, # Pass type
        [string]$authPathValue,   # Pass value
        [string]$ytDlpPath, 
        [string]$PlaylistUrlOverride = $null,
        [switch]$TreatErrorsAsWarnings = $false
    )
    $playlistUrl = $PlaylistUrlOverride
    if (-not $playlistUrl) { $playlistUrl = Read-Host "Please enter the YouTube playlist URL" }
    if (-not $playlistUrl) { Write-Warning "No URL provided. Exiting."; return }

    $baseOutputDir = $script:downloadRootPath

    # --- Pre-fetch Playlist Title --- 
    Write-Host "Fetching playlist title to determine output folder..." -ForegroundColor Cyan
    $playlistTitle = $null
    $playlistFolder = "Playlist_UnknownTitle" 
    $playlistIdForNaming = $null
    $playlistOutputDir = Join-Path -Path $baseOutputDir -ChildPath $playlistFolder

    # Build arg list for prefetch
    $prefetchArgList = @()
    if ($authTypeValue -eq 'cookies_browser' -and $authPathValue) {
        $prefetchArgList += '--cookies-from-browser'
        $prefetchArgList += $authPathValue # Now contains "firefox:profilename"
    }
    $prefetchArgList += '-J', '--flat-playlist', '--playlist-items', '0', $playlistUrl
    $jsonOutput = $null
    $prefetchExitCode = 1
    
    try {
        Write-Host "Running yt-dlp prefetch with arguments: $($prefetchArgList -join ' ' )" -ForegroundColor DarkGray
        # Execute directly, capture output
        $jsonOutput = (& $ytDlpPath $prefetchArgList) | Out-String
        $prefetchExitCode = $LASTEXITCODE

        if ($prefetchExitCode -ne 0) {
            Write-Warning "  yt-dlp prefetch command failed (Exit Code: $prefetchExitCode). Output: $jsonOutput"
            # Proceed with default folder name
        } else {
            # Try converting from JSON directly
            try {
                $playlistInfo = $jsonOutput | ConvertFrom-Json -ErrorAction Stop 
            } catch {
                Write-Warning "  Failed to parse direct JSON output. Error: $($_.Exception.Message)"
                Write-Warning "  Raw output was: `n$jsonOutput" 
                # Fallback: Try finding title in non-JSON output (less reliable)
                $jsonOutput -split '\r?\n' | Select-String -Pattern 'Downloading playlist:' | ForEach-Object { 
                    $playlistTitle = ($_.Line -replace '.*Downloading playlist:\s*','').Trim() 
                    Write-Host "  Found title via fallback regex: $playlistTitle" -ForegroundColor Yellow
                } 
            }

            # If JSON parsing succeeded, extract title and id
            if ($null -ne $playlistInfo) {
                if ($playlistInfo._type -eq 'playlist') {
                    if ($playlistInfo.title) { $playlistTitle = $playlistInfo.title }
                    if ($playlistInfo.id) { $playlistIdForNaming = $playlistInfo.id }
                    Write-Host "  Playlist title/id for naming: $playlistTitle / $playlistIdForNaming" -ForegroundColor Green
                } else {
                    Write-Warning "  JSON parsed, but missing '_type: playlist'. Using defaults." 
                }
            }
        }
        
        # Choose folder name with ID suffix for uniqueness
        if ($playlistTitle) { 
            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            $sanitizedTitle = $playlistTitle
            foreach ($char in $invalidChars) { $sanitizedTitle = $sanitizedTitle.Replace($char, '_') }
            $sanitizedTitle = $sanitizedTitle -replace '_+', '_'
            $sanitizedTitle = $sanitizedTitle.Trim('_')
            if (-not [string]::IsNullOrWhiteSpace($sanitizedTitle)) { $playlistFolder = $sanitizedTitle }
        }
        if ($playlistIdForNaming) { $playlistFolder = "$playlistFolder [$playlistIdForNaming]" }

        # Construct final output path
        $playlistOutputDir = Join-Path -Path $baseOutputDir -ChildPath $playlistFolder

    } catch {
        Write-Warning "  Error running/parsing yt-dlp prefetch: $($_.Exception.Message)"
        Write-Warning "  Unable to determine specific playlist folder name, using default: $playlistFolder"
    }

    # --- Setup playlist directories --- 
    if (-not (Test-Path $playlistOutputDir -PathType Container)) {
         try {
            Write-Host "Creating playlist directory: $playlistOutputDir" -ForegroundColor Cyan
            New-Item -ItemType Directory -Path $playlistOutputDir -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "Failed to create directory $playlistOutputDir. Check permissions. Error: $($_.Exception.Message)"
            return # Cannot proceed without output dir
        }
    }
    # No subfolders; download everything directly into the playlist folder

    $logFilePath = Join-Path -Path $playlistOutputDir -ChildPath "_download.log"
    Write-Host "Logging yt-dlp output to: $logFilePath" -ForegroundColor Cyan

    # Per-playlist archive inside the playlist folder
    $archivePath = Join-Path -Path $playlistOutputDir -ChildPath "download_archive.txt"

    # --- Construct Main Download Arguments --- 
    $templateDir = $playlistOutputDir.Replace('\\', '/')

    # Argument List for the main download
    $mainArgList = @()
    if ($ffmpegLocationArg) { $mainArgList += $ffmpegLocationArg.Split(' ', 2) }
    if ($commonArgs) { $mainArgList += $commonArgs }
    $mainArgList += '--download-archive', "`"$archivePath`""
    if ($authTypeValue -eq 'cookies_browser' -and $authPathValue) {
        $mainArgList += '--cookies-from-browser'
        $mainArgList += $authPathValue
    }
    
    # Add output template with proper quoting for Windows paths
    $mainArgList += '--output', "`"$templateDir/%(title)s [%(id)s].%(ext)s`""
    
    $mainArgList += $playlistUrl

    # LOG: print without expanding arrays that might be interpreted strangely
    Write-Host ("Executing main download with arguments: {0}" -f ($mainArgList -join ' ')) -ForegroundColor Yellow
    
    # --- Execute and Log --- 
    $downloadExitCode = 1
    try {
        & $ytDlpPath $mainArgList
        $downloadExitCode = $LASTEXITCODE
        if ($downloadExitCode -ne 0) { 
            $errorMessage = "yt-dlp finished with exit code $downloadExitCode (some videos may have been unavailable or failed to download)"
            if ($TreatErrorsAsWarnings) {
                Write-Warning $errorMessage
                Write-Host "Playlist download completed with some warnings. Check the output above for details." -ForegroundColor Yellow
            } else {
                Write-Error $errorMessage
            }
        } else { 
            Write-Host "Playlist download completed successfully." -ForegroundColor Green 
        }
    } catch { 
        $launchError = "An error occurred launching yt-dlp for playlist download: $($_.Exception.Message)"
        if ($TreatErrorsAsWarnings) {
            Write-Warning $launchError
        } else {
            Write-Error $launchError
        }
    }
}

# Function to list user's playlists and initiate download
function List-And-Download-My-Playlists {
    param(
        [string]$ffmpegLocationArg,
        [string]$commonArgs,
        [string]$authTypeValue, # Pass type
        [string]$authPathValue,   # Pass value
        [string]$ytDlpPath,
        [switch]$ForceRefreshCache = $false,
        [switch]$TreatErrorsAsWarnings = $false
    )
    if ($authTypeValue -ne 'cookies_browser' -or -not $authPathValue) { Write-Warning "Authentication required..."; return }
    
    # Get the last playlist index and ID from our script-level variable
    $defaultPlaylistIndex = $script:lastPlaylistIndex
    $defaultPlaylistId = $script:lastPlaylistId
    
    # Setup cache file path (still under project root)
    $cacheDir = Join-Path -Path $PSScriptRoot -ChildPath "cache"
    if (-not (Test-Path $cacheDir -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $cacheDir -Force -ErrorAction Stop | Out-Null
            Write-Host "Created cache directory: $cacheDir"
        } catch {
            Write-Warning "Failed to create cache directory: $($_.Exception.Message)"
        }
    }
    $playlistCacheFile = Join-Path -Path $cacheDir -ChildPath "playlists_cache.json"
    $maxCacheAgeHours = 24 # Cache expires after 24 hours
    
    $useCache = $false
    $cacheAge = $null
    
    # Check if we have a valid cache file
    if (-not $ForceRefreshCache -and (Test-Path $playlistCacheFile)) {
        try {
            $cacheFileInfo = Get-Item $playlistCacheFile
            $cacheAge = (Get-Date) - $cacheFileInfo.LastWriteTime
            
            # Check if cache is fresh enough (less than maxCacheAgeHours old)
            if ($cacheAge.TotalHours -lt $maxCacheAgeHours) {
                $useCache = $true
                Write-Host "Using playlist cache (last updated $([Math]::Floor($cacheAge.TotalHours)) hours and $($cacheAge.Minutes) minutes ago)" -ForegroundColor Cyan
            } else {
                Write-Host "Playlist cache is too old ($([Math]::Floor($cacheAge.TotalHours)) hours old), refreshing..." -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Error checking cache file: $($_.Exception.Message)"
        }
    } elseif ($ForceRefreshCache) {
        Write-Host "Force refresh requested, fetching fresh playlist data..." -ForegroundColor Yellow
    } else {
        Write-Host "No playlist cache found, creating new cache..." -ForegroundColor Yellow
    }
    
    $playlists = @()
    $feedUrl = "https://www.youtube.com/feed/playlists"
    
    if ($useCache) {
        # Read from cache
        try {
            $cachedData = Get-Content -Path $playlistCacheFile -Raw | ConvertFrom-Json -ErrorAction Stop
            $playlists = $cachedData.entries | Where-Object { $_.url -and $_.title }
            
            if ($playlists.Count -eq 0) {
                Write-Warning "Cache file contains no playlists, fetching fresh data..."
                $useCache = $false
            }
        } catch {
            Write-Warning "Failed to read cache file: $($_.Exception.Message)"
            $useCache = $false
        }
    }
    
    if (-not $useCache) {
        Write-Host "Fetching your YouTube playlists... (This requires authentication)" -ForegroundColor Cyan
        
        # Build Argument list for listing playlists
        $listArgList = @()
        $listArgList += '--cookies-from-browser'
        $listArgList += $authPathValue # Now contains "firefox:profilename"
        $listArgList += '-J', '--flat-playlist', $feedUrl
        
        $jsonOutput = $null
        $listExitCode = 1

        try {
            Write-Host "Running yt-dlp list command with arguments: $($listArgList -join ' ' )" -ForegroundColor DarkGray
            # Execute directly, capture output
            $jsonOutput = (& $ytDlpPath $listArgList) | Out-String
            $listExitCode = $LASTEXITCODE
            
            if ($listExitCode -ne 0) {
                Write-Error "Failed to list playlists (Exit Code: $listExitCode). Output: $jsonOutput"
                return
            }

            $feedInfo = $null
            try {
                $feedInfo = $jsonOutput | ConvertFrom-Json -ErrorAction Stop
                
                # Save to cache file
                try {
                    $feedInfo | ConvertTo-Json -Depth 10 | Set-Content -Path $playlistCacheFile -Encoding UTF8 -Force
                    Write-Host "Playlist data cached successfully." -ForegroundColor Green
                } catch {
                    Write-Warning "Failed to save playlist cache: $($_.Exception.Message)"
                }
                
            } catch {
                Write-Error "Failed to parse playlist feed JSON: $($_.Exception.Message)"
                Write-Error "Raw output: `n$jsonOutput"
                return
            }

            if ($null -ne $feedInfo -and $feedInfo.entries) {
                $playlists = $feedInfo.entries | Where-Object { $_.url -and $_.title } 
            } else {
                Write-Error "Could not find playlist entries in the feed JSON."
                Write-Error "Raw output: `n$jsonOutput"
                return
            }
        } catch {
             Write-Error "An error occurred while fetching or processing playlists: $($_.Exception.Message)"
             return
        }
    }
    
    if ($playlists.Count -eq 0) {
         Write-Warning "No playlists found."
         return
    }

    # Pagination variables
    $pageSize = 20
    $currentPage = 0
    $totalPages = [Math]::Ceiling($playlists.Count / $pageSize)
    $selection = ""
    
    # If we have a playlist ID saved, try to find its current index
    $foundDefaultIndex = $null
    if ($defaultPlaylistId) {
        for ($i = 0; $i -lt $playlists.Count; $i++) {
            if ($playlists[$i].id -eq $defaultPlaylistId -or 
                $playlists[$i].url -match [regex]::Escape($defaultPlaylistId)) {
                $foundDefaultIndex = $i + 1 # Convert to 1-based index
                $currentPage = [Math]::Floor($i / $pageSize) # Go to the page containing this playlist
                Write-Host "Found your last selected playlist at position #$foundDefaultIndex" -ForegroundColor Cyan
                break
            }
        }
        
        if ($null -eq $foundDefaultIndex) {
            Write-Host "Your previously selected playlist was not found in the current list." -ForegroundColor Yellow
            # In this case, we'll fall back to the index-based selection if on the first page
        }
    }

    # Display playlists with pagination
    while ($true) {
        $startIndex = $currentPage * $pageSize
        $endIndex = [Math]::Min($startIndex + $pageSize - 1, $playlists.Count - 1)
        
        Write-Host "`nYour Playlists (Page $($currentPage + 1) of $totalPages):" -ForegroundColor Green
        for ($i = $startIndex; $i -le $endIndex; $i++) {
            $indexDisplay = ($i + 1).ToString()
            # Highlight the previously selected playlist if it's on this page
            if ($foundDefaultIndex -and ($i + 1) -eq $foundDefaultIndex) {
                Write-Host ("  -> {0,3}. {1}" -f $indexDisplay, $playlists[$i].title) -ForegroundColor Cyan
            } else {
                Write-Host ("    {0,3}. {1}" -f $indexDisplay, $playlists[$i].title)
            }
        }

        if ($currentPage -lt $totalPages - 1) {
            $prompt = "`nEnter playlist number to download, press Enter for next page, or 'q' to quit"
            # Use the found default if available, otherwise fall back to the index
            if ($foundDefaultIndex) {
                $prompt += " [Last: $foundDefaultIndex]"
            } elseif ($defaultPlaylistIndex -and $currentPage -eq 0) {
                $prompt += " [Last: $defaultPlaylistIndex]"
            }
        } else {
            $prompt = "`nEnter playlist number to download or 'q' to quit (last page)"
            if ($foundDefaultIndex) {
                $prompt += " [Last: $foundDefaultIndex]"
            } elseif ($defaultPlaylistIndex -and $currentPage -eq 0) {
                $prompt += " [Last: $defaultPlaylistIndex]"
            }
        }
        
        $selection = Read-Host $prompt
        
        # Check selection
        if ([string]::IsNullOrWhiteSpace($selection)) {
            # If we found the default playlist by ID, use that index
            if ($foundDefaultIndex) {
                $selection = $foundDefaultIndex.ToString()
                Write-Host "Using last playlist choice: $selection" -ForegroundColor DarkGray
                break
            }
            # Otherwise, use index-based if on first page
            elseif ($defaultPlaylistIndex -and $currentPage -eq 0) {
                $selection = $defaultPlaylistIndex
                Write-Host "Using last playlist choice by position: $selection" -ForegroundColor DarkGray
                break
            }
            
            # If no default or not on the right page, go to next page
            if ($currentPage -lt $totalPages - 1) {
                $currentPage++
            } else {
                Write-Host "Last page reached." -ForegroundColor Yellow
            }
            continue
        } elseif ($selection -eq "q") {
            Write-Host "Download cancelled."
            return
        } elseif ($selection -match '^[1-9][0-9]*$') {
            # Valid number entered, break the loop to process
            break
        } else {
            Write-Host "Invalid input. Please enter a valid playlist number, press Enter for next page, or 'q' to quit." -ForegroundColor Yellow
        }
    }

    # Process playlist selection
    $index = [int]$selection - 1
    if ($index -ge 0 -and $index -lt $playlists.Count) {
        $selectedPlaylist = $playlists[$index]
        Write-Host "Selected: $($selectedPlaylist.title)" -ForegroundColor Yellow
        
        # Extract playlist ID from URL
        $playlistId = $null
        if ($selectedPlaylist.id) {
            $playlistId = $selectedPlaylist.id
        } elseif ($selectedPlaylist.url -match 'list=([^&]+)') {
            $playlistId = $matches[1]
        }
        
        # Save the playlist choice globally with ID when available, preserving root
        Save-UserPreferences -menuChoice "3" -playlistIndex $selection -playlistId $playlistId -downloadRoot $script:downloadRootPath
        
        Download-Playlist -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlagsCore -authTypeValue $authType -authPathValue $authValue -PlaylistUrlOverride $selectedPlaylist.url -ytDlpPath $ytDlpPath -TreatErrorsAsWarnings:$script:useWarningsForYtDlp
    } else {
        Write-Warning "Invalid selection number."
    }
}

# --- Script Main Execution ---

# Helper: read menu choice with Esc handling
function Read-MenuInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$Prompt,
        [string]$Default
    )
    # Show prompt
    Write-Host $Prompt -NoNewline
    $builder = [System.Text.StringBuilder]::new()
    $usedDefault = $false
    while ($true) {
        $keyInfo = [System.Console]::ReadKey($true)
        switch ($keyInfo.Key) {
            'Escape' {
                Write-Host ""
                return [pscustomobject]@{ Escaped = $true; Input = $null; UsedDefault = $false }
            }
            'Enter' {
                Write-Host ""
                $text = $builder.ToString()
                if ([string]::IsNullOrWhiteSpace($text) -and -not [string]::IsNullOrWhiteSpace($Default)) {
                    $text = $Default
                    $usedDefault = $true
                }
                return [pscustomobject]@{ Escaped = $false; Input = $text; UsedDefault = $usedDefault }
            }
            'Backspace' {
                if ($builder.Length -gt 0) {
                    $builder.Length--
                    Write-Host "`b `b" -NoNewline
                }
            }
            Default {
                if ($keyInfo.KeyChar) {
                    [void]$builder.Append($keyInfo.KeyChar)
                    Write-Host $keyInfo.KeyChar -NoNewline
                }
            }
        }
    }
}

Write-Host "`nPlease choose an action:" -ForegroundColor Green
Write-Host "1. Download Single Video (Best Quality + Metadata)"
Write-Host "2. Download Playlist by URL (Best Quality + Metadata)"
if ($script:authAvailable) {
    Write-Host "3. List & Download My Playlist (Requires Auth)"
    Write-Host "4. List & Download My Playlist (Refresh Cache)"
} else {
    Write-Host "3. List & Download My Playlist (Requires Auth) - Unavailable (missing Firefox profile)" -ForegroundColor DarkGray
    Write-Host "4. List & Download My Playlist (Refresh Cache) - Unavailable (missing Firefox profile)" -ForegroundColor DarkGray
}

$defaultChoice = if ($script:lastChoice) { $script:lastChoice } else { "" }
$menuPrompt = "Enter your choice (1, 2, 3, or 4)$(if ($defaultChoice) { " [Last: $defaultChoice]" }) (Esc to exit): "
$menuInput = Read-MenuInput -Prompt $menuPrompt -Default $defaultChoice
if ($menuInput.Escaped) {
    Write-Host "Exiting by user request." -ForegroundColor Cyan
    return
}
$choice = $menuInput.Input

# If empty input and we have a default, use the default
if ($menuInput.UsedDefault) {
    Write-Host "Using last choice: $choice" -ForegroundColor DarkGray
} elseif ([string]::IsNullOrWhiteSpace($choice) -and $defaultChoice) {
    $choice = $defaultChoice
    Write-Host "Using last choice: $choice" -ForegroundColor DarkGray
}

try {
    # Save user's menu choice (keep per-root and global in sync)
    if (-not [string]::IsNullOrWhiteSpace($choice)) {
        if ($TreatYtDlpErrorsAsWarnings) { $script:treatYtDlpErrorsAsWarningsPref = $true }
        Save-UserPreferences -menuChoice $choice -playlistIndex $script:lastPlaylistIndex -playlistId $script:lastPlaylistId -downloadRoot $script:downloadRootPath
    }

    # Pass the confirmed yt-dlp path and separated auth info
    switch ($choice) {
        "1" {
            Write-Host "Selected: Download Single Video" -ForegroundColor Yellow
            Download-BestVideo -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlagsCore -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath -TreatErrorsAsWarnings:$script:useWarningsForYtDlp
        }
        "2" {
            Write-Host "Selected: Download Playlist by URL" -ForegroundColor Yellow
            Download-Playlist -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlagsCore -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath -TreatErrorsAsWarnings:$script:useWarningsForYtDlp
        }
        "3" {
            if (-not $script:authAvailable) {
                Write-Warning "Option 3 is unavailable because a valid Firefox profile with YouTube cookies was not found. Running in public-only mode."
                break
            }
            Write-Host "Selected: List & Download My Playlist" -ForegroundColor Yellow
            List-And-Download-My-Playlists -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlagsCore -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath -TreatErrorsAsWarnings:$script:useWarningsForYtDlp
        }
        "4" {
            if (-not $script:authAvailable) {
                Write-Warning "Option 4 is unavailable because a valid Firefox profile with YouTube cookies was not found. Running in public-only mode."
                break
            }
            Write-Host "Selected: List & Download My Playlist (Force Refresh Cache)" -ForegroundColor Yellow
            List-And-Download-My-Playlists -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlagsCore -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath -ForceRefreshCache -TreatErrorsAsWarnings:$script:useWarningsForYtDlp
        }
        default {
            Write-Warning "Invalid choice. Exiting."
        }
    }
}
catch [System.Management.Automation.PipelineStoppedException] {
    Write-Warning "Script execution stopped by user (Ctrl+C)."
}
catch {
    Write-Error "An unexpected error occurred in the main script body: $($_.Exception.Message)"
}
finally {
    Write-Host "Script finished or exited." -ForegroundColor Cyan
}