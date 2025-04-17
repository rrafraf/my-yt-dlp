#Requires -Version 5.0
param()

# --- START Preference Handling ---
$userPrefsFile = Join-Path -Path $PSScriptRoot -ChildPath "user_preferences.json"
$script:currentYtDlpVersion = $null # Stores the version tag of the current yt-dlp.exe
$script:lastChoice = $null # Keep existing preference
$script:lastPlaylistIndex = $null # Keep existing preference
$script:lastPlaylistId = $null # Keep existing preference

# Function to save user preferences
function Save-UserPreferences {
    [CmdletBinding()]
    param(
        [string]$menuChoice = $script:lastChoice,
        [string]$playlistIndex = $script:lastPlaylistIndex,
        [string]$playlistId = $script:lastPlaylistId,
        [string]$ytDlpVersion = $script:currentYtDlpVersion # Use current script value if not passed
    )
    
    # Update script variables before saving
    $script:lastChoice = $menuChoice
    $script:lastPlaylistIndex = $playlistIndex
    $script:lastPlaylistId = $playlistId
    $script:currentYtDlpVersion = $ytDlpVersion

    $preferences = @{
        lastMenuChoice = $script:lastChoice
        lastPlaylistIndex = $script:lastPlaylistIndex
        lastPlaylistId = $script:lastPlaylistId
        currentYtDlpVersion = $script:currentYtDlpVersion # Save the version
    }
    
    try {
        $preferences | ConvertTo-Json -Depth 5 | Set-Content -Path $userPrefsFile -Encoding UTF8 -Force
    } catch {
        Write-Warning "Could not save preferences to '$userPrefsFile': $($_.Exception.Message)"
    }
}

# Function to load user preferences
function Load-UserPreferences {
    if (Test-Path $userPrefsFile) {
        try {
            $userPrefs = Get-Content -Path $userPrefsFile -Raw | ConvertFrom-Json -ErrorAction Stop
            
            if ($userPrefs.PSObject.Properties.Name -contains 'lastMenuChoice') { $script:lastChoice = $userPrefs.lastMenuChoice }
            if ($userPrefs.PSObject.Properties.Name -contains 'lastPlaylistIndex') { $script:lastPlaylistIndex = $userPrefs.lastPlaylistIndex }
            if ($userPrefs.PSObject.Properties.Name -contains 'lastPlaylistId') { $script:lastPlaylistId = $userPrefs.lastPlaylistId }
            if ($userPrefs.PSObject.Properties.Name -contains 'currentYtDlpVersion') { $script:currentYtDlpVersion = $userPrefs.currentYtDlpVersion }

            Write-Host "Loaded preferences: YT-DLP Ver: $($script:currentYtDlpVersion), Last Choice: $($script:lastChoice), Last Playlist: $($script:lastPlaylistIndex)" -ForegroundColor DarkGray
        } catch {
            Write-Warning "Could not load or parse user preferences from '$userPrefsFile': $($_.Exception.Message). Using defaults."
        }
    } else {
        Write-Host "Preference file '$userPrefsFile' not found. Using defaults." -ForegroundColor DarkGray
    }
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
$firefoxProfilePath = 'C:\Users\gentl\AppData\Roaming\Mozilla\Firefox\Profiles\fw0m6kre.default-nightly' # !!! UPDATED WITH USER PATH !!!

if (-not (Test-Path $firefoxProfilePath -PathType Container)) {
    Write-Error "FATAL: The specified Firefox profile path does not exist: $firefoxProfilePath"
    Write-Error "Please correct the path in the script (line starting with `$firefoxProfilePath =`). You can find it in Firefox Nightly at about:profiles."
    exit 1
}

# Extract just the profile folder name without the full path
$profileFolderName = Split-Path -Leaf $firefoxProfilePath
# Format as "firefox:profilename" for yt-dlp
$authType = 'cookies_browser' 
$authValue = "firefox:$profileFolderName" 
Write-Host "Using authentication: $authType with value $authValue" -ForegroundColor Yellow
# --- End Authentication Setup ---

# --- Download Functions ---
$downloadArchiveFile = Join-Path -Path $PSScriptRoot -ChildPath "download_archive.txt"
# Define common metadata/download flags
# Build complex arguments outside the array
$archiveArg = "--download-archive ""$downloadArchiveFile"""

# Simple array of flags - NO trailing backslashes!
$commonFlags = @(
    $archiveArg,
    '--write-description',
    '--write-info-json',
    '--write-subs',
    '--write-auto-subs',
    '--sub-langs "en.*,en"',
    '--embed-metadata',
    '--embed-thumbnail',
    '--embed-subs',
    '--progress-delta 2' # Keep progress updates throttled
) -join " "

# Function to download the best quality single video with metadata
function Download-BestVideo {
    param(
        [string]$ffmpegLocationArg,
        [string]$commonArgs,
        [string]$authTypeValue, # Pass type
        [string]$authPathValue,   # Pass value
        [string]$ytDlpPath 
    )
    $videoUrl = Read-Host "Please enter the YouTube video URL"
    if (-not $videoUrl) { Write-Warning "No URL provided. Exiting."; return }
    
    $outputDir = Join-Path $PSScriptRoot "Downloads"
    if (-not (Test-Path $outputDir -PathType Container)) { 
        try {
            New-Item -ItemType Directory -Path $outputDir -Force -ErrorAction Stop | Out-Null
            Write-Host "Created output directory: $outputDir"
        } catch {
             Write-Error "Failed to create directory $outputDir. Check permissions. Error: $($_.Exception.Message)"
             return # Cannot proceed
        }
    }

    # Construct the output template 
    $outputTemplate = Join-Path $outputDir "%(title)s [%(id)s].%(ext)s"
    
    # Build Argument List
    $ArgumentList = @()
    if ($ffmpegLocationArg) { $ArgumentList += $ffmpegLocationArg.Split(' ', 2) } 
    if ($commonArgs) { $ArgumentList += $commonArgs.Split(' ') } 
    # Add auth args if present
    if ($authTypeValue -eq 'cookies_browser' -and $authPathValue) {
        $ArgumentList += '--cookies-from-browser'
        $ArgumentList += $authPathValue  # Now contains "firefox:profilename"
    }
    $ArgumentList += '-o', $outputTemplate 
    $ArgumentList += $videoUrl

    Write-Host "Executing yt-dlp with arguments: $($ArgumentList -join ' ' )" -ForegroundColor Yellow
    try {
        & $ytDlpPath $ArgumentList
        if ($LASTEXITCODE -ne 0) {
            Write-Error "yt-dlp failed with exit code $LASTEXITCODE. Check console output above for details."
        } else {
            Write-Host "Download completed successfully (or yt-dlp finished)." -ForegroundColor Green
        }
    } catch {
        Write-Error "An error occurred launching yt-dlp: $($_.Exception.Message)"
    }
}

# Function to download a playlist with metadata into a named subfolder
function Download-Playlist {
    param(
        [string]$ffmpegLocationArg,
        [string]$commonArgs,
        [string]$authTypeValue, # Pass type
        [string]$authPathValue,   # Pass value
        [string]$ytDlpPath, 
        [string]$PlaylistUrlOverride = $null
    )
    $playlistUrl = $PlaylistUrlOverride
    if (-not $playlistUrl) { $playlistUrl = Read-Host "Please enter the YouTube playlist URL" }
    if (-not $playlistUrl) { Write-Warning "No URL provided. Exiting."; return }

    $baseOutputDir = Join-Path $PSScriptRoot "Downloads"

    # --- Pre-fetch Playlist Title --- 
    Write-Host "Fetching playlist title to determine output folder..." -ForegroundColor Cyan
    $playlistTitle = $null
    $playlistFolder = "Playlist_UnknownTitle" 
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

            # If JSON parsing succeeded, extract title
            if ($playlistInfo -ne $null) {
                if ($playlistInfo._type -eq 'playlist' -and $playlistInfo.title) {
                    $playlistTitle = $playlistInfo.title
                    Write-Host "  Playlist Title Found via JSON: $playlistTitle" -ForegroundColor Green
                } else {
                    Write-Warning "  JSON parsed, but missing '_type: playlist' or 'title' field. Using default folder name." 
                }
            }
        }
        
        # If we have a title, sanitize and set $playlistFolder
        if ($playlistTitle) { 
            # Sanitize title for folder name
            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            $sanitizedTitle = $playlistTitle
            
            # Replace invalid characters with underscore, but don't add underscores between valid characters
            foreach ($char in $invalidChars) {
                $sanitizedTitle = $sanitizedTitle.Replace($char, '_')
            }
            
            # Replace multiple sequential underscores with a single one
            $sanitizedTitle = $sanitizedTitle -replace '_+', '_'
            $sanitizedTitle = $sanitizedTitle.Trim('_') # Remove leading/trailing underscores
            
            if (-not [string]::IsNullOrWhiteSpace($sanitizedTitle)) {
                $playlistFolder = $sanitizedTitle
            } else {
                Write-Warning "  Playlist title '$playlistTitle' resulted in an empty folder name after sanitization. Using default: $playlistFolder"
            }
        } else {
            Write-Warning "  Could not determine playlist title. Using default folder: $playlistFolder"
        } 

        # Construct final output path based on fetched/sanitized/default title
        $playlistOutputDir = Join-Path -Path $baseOutputDir -ChildPath $playlistFolder

    } catch {
        Write-Warning "  Error running/parsing yt-dlp prefetch: $($_.Exception.Message)"
        Write-Warning "  Unable to determine specific playlist folder name, using default: $playlistFolder"
    }

    # --- Setup Log File Path --- 
    if (-not (Test-Path $playlistOutputDir -PathType Container)) {
         try {
            Write-Host "Creating playlist directory: $playlistOutputDir" -ForegroundColor Cyan
            New-Item -ItemType Directory -Path $playlistOutputDir -Force -ErrorAction Stop | Out-Null
        } catch {
             Write-Error "Failed to create directory $playlistOutputDir. Check permissions. Error: $($_.Exception.Message)"
             return # Cannot proceed without output dir
        }
    }
    $logFilePath = Join-Path -Path $playlistOutputDir -ChildPath "_download.log"
    Write-Host "Logging yt-dlp output to: $logFilePath" -ForegroundColor Cyan

    # --- Construct Main Download Arguments --- 
    $templateDir = $playlistOutputDir.Replace('\', '/') 
    $templateFilename = "%(title)s [%(id)s]"
    $fullTemplateBase = "$templateDir/$templateFilename"

    # Argument List for the main download
    $mainArgList = @()
    if ($ffmpegLocationArg) { $mainArgList += $ffmpegLocationArg.Split(' ', 2) }
    if ($commonArgs) { $mainArgList += $commonArgs.Split(' ') }
    if ($authTypeValue -eq 'cookies_browser' -and $authPathValue) {
        $mainArgList += '--cookies-from-browser'
        $mainArgList += $authPathValue # Now contains "firefox:profilename"
    }
    
    # Add output template with proper quoting for Windows paths
    # Use individual templates for each output type to ensure proper naming
    $mainArgList += '--output', "`"$templateDir/%(title)s [%(id)s].%(ext)s`""
    $mainArgList += '--output-na-placeholder', "`"`""
    $mainArgList += '--paths', "temp:$templateDir/temp"  # Temp directory for partial downloads
    $mainArgList += '--paths', "home:$templateDir"       # Set home path
    
    # Add the playlist URL as the final argument
    $mainArgList += $playlistUrl

    Write-Host "Executing main download with arguments: $($mainArgList -join ' ' )" -ForegroundColor Yellow
    
    # --- Execute and Log --- 
    $downloadOutput = $null
    $downloadExitCode = 1
    try {
        # Setup log files
        $stdoutLogPath = Join-Path -Path $playlistOutputDir -ChildPath "_download.log"
        $stderrLogPath = Join-Path -Path $playlistOutputDir -ChildPath "_download_error.log"
        
        # Start the process asynchronously with proper output redirection
        Write-Host "Starting download process (this may take a while)..." -ForegroundColor Cyan
        
        # Use the console window to show output directly
        $process = Start-Process -FilePath $ytDlpPath -ArgumentList $mainArgList -NoNewWindow -PassThru -Wait
        $downloadExitCode = $process.ExitCode
        
        if ($downloadExitCode -ne 0) {
             Write-Error "yt-dlp failed with exit code $downloadExitCode."
        } else {
            Write-Host "Playlist download completed successfully." -ForegroundColor Green
        }

    } catch {
        Write-Error "An error occurred launching yt-dlp for playlist download: $($_.Exception.Message)"
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
        [switch]$ForceRefreshCache = $false
    )
    if ($authTypeValue -ne 'cookies_browser' -or -not $authPathValue) { Write-Warning "Authentication required..."; return }
    
    # Get the last playlist index and ID from our script-level variable
    $defaultPlaylistIndex = $script:lastPlaylistIndex
    $defaultPlaylistId = $script:lastPlaylistId
    
    # Setup cache file path
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

            if ($feedInfo -ne $null -and $feedInfo.entries) {
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
                Write-Host ("  → {0,3}. {1}" -f $indexDisplay, $playlists[$i].title) -ForegroundColor Cyan
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
        
        # Save the playlist choice globally with ID when available
        Save-UserPreferences -menuChoice "3" -playlistIndex $selection -playlistId $playlistId
        
        Download-Playlist -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlags -authTypeValue $authType -authPathValue $authValue -PlaylistUrlOverride $selectedPlaylist.url -ytDlpPath $ytDlpPath
    } else {
        Write-Warning "Invalid selection number."
    }
}

# --- Script Main Execution ---

# Path for saved user preferences
$userPrefsFile = Join-Path -Path $PSScriptRoot -ChildPath "user_preferences.json"
$script:lastChoice = $null
$script:lastPlaylistIndex = $null
$script:lastPlaylistId = $null

# Function to save user preferences
function Save-UserPreferences {
    param(
        [string]$menuChoice,
        [string]$playlistIndex = $null,
        [string]$playlistId = $null
    )
    
    $preferences = @{
        lastMenuChoice = $menuChoice
        lastPlaylistIndex = $playlistIndex
        lastPlaylistId = $playlistId
    }
    
    try {
        $preferences | ConvertTo-Json | Set-Content -Path $userPrefsFile -Encoding UTF8 -Force
        Write-Host "Preferences saved." -ForegroundColor DarkGray
    } catch {
        Write-Warning "Could not save preferences: $($_.Exception.Message)"
    }
}

# Try to load user preferences if they exist
if (Test-Path $userPrefsFile) {
    try {
        $userPrefs = Get-Content -Path $userPrefsFile -Raw | ConvertFrom-Json -ErrorAction Stop
        $script:lastChoice = $userPrefs.lastMenuChoice
        $script:lastPlaylistIndex = $userPrefs.lastPlaylistIndex
        $script:lastPlaylistId = $userPrefs.lastPlaylistId
        Write-Host "Loaded previous preferences. Last menu choice: $($script:lastChoice), Last playlist: $($script:lastPlaylistIndex)" -ForegroundColor DarkGray
    } catch {
        Write-Warning "Could not load user preferences: $($_.Exception.Message)"
    }
}

Write-Host "Starting yt-dlp Helper Script..." -ForegroundColor Cyan

# Construct and resolve path for yt-dlp.exe IN THE SCRIPT DIRECTORY
$expectedYtDlpExePath = Join-Path $PSScriptRoot "yt-dlp.exe"
$ytDlpExePath = $null # Initialize
try {
    $ytDlpExePath = Resolve-Path -Path $expectedYtDlpExePath -ErrorAction Stop
} catch {
    Write-Error "FATAL: Could not resolve expected path for yt-dlp.exe: $expectedYtDlpExePath. Error: $($_.Exception.Message)"
    exit 1
}

# Check if the resolved path points to an existing file
if (-not (Test-Path $ytDlpExePath -PathType Leaf)) {
    Write-Error "FATAL: yt-dlp.exe not found at resolved location: $ytDlpExePath. Please ensure it exists in the same directory as this script ('$PSScriptRoot')."
    exit 1
}
Write-Host "Found yt-dlp.exe at: $ytDlpExePath" -ForegroundColor Green

# --- FFmpeg Setup Logic --- 
# ... (FFmpeg setup logic uses $PSScriptRoot for its paths)
# ... (Result is $ffmpegLocationArgument)

# --- Authentication Setup --- (Now done above)

# --- Common Flags Definition --- 
$downloadArchiveFile = Join-Path -Path $PSScriptRoot -ChildPath "download_archive.txt"
$archiveArg = "--download-archive ""$downloadArchiveFile"""
$commonFlags = @(
    $archiveArg,
    '--write-description',
    '--write-info-json',
    '--write-subs',
    '--write-auto-subs',
    '--sub-langs "en.*,en"',
    '--embed-metadata',
    '--embed-thumbnail',
    '--embed-subs',
    '--progress-delta 2' 
) -join " "

# --- Menu --- 
Write-Host "`nPlease choose an action:" -ForegroundColor Green
Write-Host "1. Download Single Video (Best Quality + Metadata)"
Write-Host "2. Download Playlist by URL (Best Quality + Metadata)"
Write-Host "3. List & Download My Playlist (Requires Auth)"
Write-Host "4. List & Download My Playlist (Refresh Cache)"

$defaultChoice = if ($script:lastChoice) { $script:lastChoice } else { "" }
$choice = Read-Host "Enter your choice (1, 2, 3, or 4)$(if ($defaultChoice) { " [Last: $defaultChoice]" })"

# If empty input and we have a default, use the default
if ([string]::IsNullOrWhiteSpace($choice) -and $defaultChoice) {
    $choice = $defaultChoice
    Write-Host "Using last choice: $choice" -ForegroundColor DarkGray
}

try {
    # Save user's menu choice
    if (-not [string]::IsNullOrWhiteSpace($choice)) {
        Save-UserPreferences -menuChoice $choice -playlistIndex $script:lastPlaylistIndex -playlistId $script:lastPlaylistId
    }

    # Pass the confirmed yt-dlp path and separated auth info
    switch ($choice) {
        "1" {
            Write-Host "Selected: Download Single Video" -ForegroundColor Yellow
            Download-BestVideo -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlags -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath
        }
        "2" {
            Write-Host "Selected: Download Playlist by URL" -ForegroundColor Yellow
            Download-Playlist -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlags -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath
        }
        "3" {
            Write-Host "Selected: List & Download My Playlist" -ForegroundColor Yellow
            List-And-Download-My-Playlists -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlags -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath
        }
        "4" {
            Write-Host "Selected: List & Download My Playlist (Force Refresh Cache)" -ForegroundColor Yellow
            List-And-Download-My-Playlists -ffmpegLocationArg $ffmpegLocationArgument -commonArgs $commonFlags -authTypeValue $authType -authPathValue $authValue -ytDlpPath $ytDlpExePath -ForceRefreshCache
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