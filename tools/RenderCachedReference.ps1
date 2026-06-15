param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [switch] $CacheOnly,

    [string] $CacheVariant
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path
if (-not $CacheOnly -and -not [string]::IsNullOrWhiteSpace($CacheVariant)) {
    throw "CacheVariant is only supported in cache-only mode until the Office reference renderer accepts variant-specific settings."
}

$inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
$renderReference = Join-Path $PSScriptRoot "RenderReference.ps1"
$rasterizePdf = Join-Path $PSScriptRoot "RasterizePdf.ps1"
$toolHashSource = @(
    (Get-FileHash -LiteralPath $renderReference -Algorithm SHA256).Hash
    (Get-FileHash -LiteralPath $rasterizePdf -Algorithm SHA256).Hash
) -join "|"
$toolHashBytes = [System.Text.Encoding]::UTF8.GetBytes($toolHashSource)
$sha256 = [System.Security.Cryptography.SHA256]::Create()
try {
    $toolHash = ([System.BitConverter]::ToString($sha256.ComputeHash($toolHashBytes)) -replace "-", "").ToLowerInvariant()
}
finally {
    $sha256.Dispose()
}
$extension = [System.IO.Path]::GetExtension($inputFull).TrimStart(".").ToLowerInvariant()
$variantKeyPart = ""
if (-not [string]::IsNullOrWhiteSpace($CacheVariant)) {
    $variantBytes = [System.Text.Encoding]::UTF8.GetBytes($CacheVariant.Trim().ToLowerInvariant())
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $variantHash = ([System.BitConverter]::ToString($sha256.ComputeHash($variantBytes)) -replace "-", "").ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }

    $variantKeyPart = "-variant" + $variantHash.Substring(0, 12)
}

$key = "{0}-{1}-{2}{3}-dpi{4}" -f $extension, $inputHash.Substring(0, 24), $toolHash.Substring(0, 12), $variantKeyPart, $Dpi
$cacheRoot = Join-Path $repoRoot "artifacts/reference-cache"
$cacheDir = Join-Path $cacheRoot $key
$completeMarker = Join-Path $cacheDir "complete.txt"

if (-not (Test-Path -LiteralPath $completeMarker)) {
    if ($CacheOnly) {
        $variantMessage = if ([string]::IsNullOrWhiteSpace($CacheVariant)) { "" } else { " for variant '$CacheVariant'" }
        throw "Reference cache miss for '$inputFull'$variantMessage at $Dpi DPI. Cache-only mode refuses to invoke Office/COM reference rendering. Expected cache directory: $cacheDir"
    }

    $tempDir = Join-Path $cacheRoot ("_tmp-" + [System.Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    try {
        & $renderReference -InputPath $inputFull -OutputDirectory $tempDir -Dpi $Dpi
        New-Item -ItemType Directory -Force -Path $cacheRoot | Out-Null
        if (Test-Path -LiteralPath $cacheDir) {
            Remove-Item -LiteralPath $cacheDir -Recurse -Force
        }

        Move-Item -LiteralPath $tempDir -Destination $cacheDir
        Set-Content -LiteralPath $completeMarker -Value ("input={0}`ndpi={1}`nvariant={2}`n" -f $inputFull, $Dpi, $CacheVariant)
    }
    finally {
        if (Test-Path -LiteralPath $tempDir) {
            Remove-Item -LiteralPath $tempDir -Recurse -Force
        }
    }
}

Copy-Item -Path (Join-Path $cacheDir "*") -Destination $outputFull -Recurse -Force
Write-Host "Reference cache: $cacheDir"
