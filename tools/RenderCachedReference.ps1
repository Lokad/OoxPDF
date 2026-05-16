param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path

$inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
$renderReference = Join-Path $PSScriptRoot "RenderReference.ps1"
$rasterizePdf = Join-Path $PSScriptRoot "RasterizePdf.ps1"
$toolHashSource = @(
    (Get-FileHash -LiteralPath $renderReference -Algorithm SHA256).Hash
    (Get-FileHash -LiteralPath $rasterizePdf -Algorithm SHA256).Hash
) -join "|"
$toolHashBytes = [System.Text.Encoding]::UTF8.GetBytes($toolHashSource)
$toolHash = [System.Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($toolHashBytes)).ToLowerInvariant()
$extension = [System.IO.Path]::GetExtension($inputFull).TrimStart(".").ToLowerInvariant()
$key = "{0}-{1}-{2}-dpi{3}" -f $extension, $inputHash.Substring(0, 24), $toolHash.Substring(0, 12), $Dpi
$cacheRoot = Join-Path $repoRoot "artifacts/reference-cache"
$cacheDir = Join-Path $cacheRoot $key
$completeMarker = Join-Path $cacheDir "complete.txt"

if (-not (Test-Path -LiteralPath $completeMarker)) {
    $tempDir = Join-Path $cacheRoot ("_tmp-" + [System.Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    try {
        & $renderReference -InputPath $inputFull -OutputDirectory $tempDir -Dpi $Dpi
        New-Item -ItemType Directory -Force -Path $cacheRoot | Out-Null
        if (Test-Path -LiteralPath $cacheDir) {
            Remove-Item -LiteralPath $cacheDir -Recurse -Force
        }

        Move-Item -LiteralPath $tempDir -Destination $cacheDir
        Set-Content -LiteralPath $completeMarker -Value ("input={0}`ndpi={1}`n" -f $inputFull, $Dpi)
    }
    finally {
        if (Test-Path -LiteralPath $tempDir) {
            Remove-Item -LiteralPath $tempDir -Recurse -Force
        }
    }
}

Copy-Item -Path (Join-Path $cacheDir "*") -Destination $outputFull -Recurse -Force
Write-Host "Reference cache: $cacheDir"
