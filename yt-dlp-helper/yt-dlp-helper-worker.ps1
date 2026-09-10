#Requires -Version 5.0
param(
    [Parameter(Mandatory = $true)]
    [string]$RequestPath,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath,
    [string]$GuiConfigPath = ''
)

. (Join-Path $PSScriptRoot 'Helper.Core.ps1')
Initialize-HelperCore -Mode worker -LoggingConfigPath $GuiConfigPath
Set-HelperLogHandler -Handler {
    param($Entry)

    if ($Entry.Level -eq 'ERROR') {
        [Console]::Error.WriteLine($Entry.Line)
    }
    else {
        [Console]::Out.WriteLine($Entry.Line)
    }
}

try {
    if (-not (Test-Path -LiteralPath $RequestPath -PathType Leaf)) {
        throw "Request file not found: $RequestPath"
    }

    $request = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    $result = Invoke-HelperOperation -Request $request
    $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ResultPath -Encoding UTF8 -Force
    if ($result.Succeeded) {
        exit 0
    }

    exit 1
}
catch {
    $failureResult = [pscustomobject]@{
        Kind      = 'worker-error'
        Succeeded = $false
        Status    = $_.Exception.Message
        Summary   = 'Worker execution failed.'
        Data      = [pscustomobject]@{
            RequestPath = $RequestPath
            ResultPath  = $ResultPath
        }
        Warnings  = @()
    }

    try {
        $failureResult | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ResultPath -Encoding UTF8 -Force
    }
    catch {
    }

    [Console]::Error.WriteLine(("[{0}] [ERROR] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $_.Exception.Message))
    exit 1
}
