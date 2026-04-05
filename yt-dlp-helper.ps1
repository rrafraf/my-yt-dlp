#Requires -Version 5.0
param(
    [switch]$TreatYtDlpErrorsAsWarnings = $false
)

$helperScript = Join-Path $PSScriptRoot 'yt-dlp-helper\yt-dlp-helper.ps1'
if (-not (Test-Path -LiteralPath $helperScript -PathType Leaf)) {
    throw "Could not find helper implementation at '$helperScript'."
}

& $helperScript @PSBoundParameters
exit $LASTEXITCODE
