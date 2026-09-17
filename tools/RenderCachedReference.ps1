# Renders or copies cached Office reference PDFs.
#
# Cache identity (protocol v2) is the reference PDF itself: input hash,
# extension, variant, and rendering protocol. Rasterizer/DPI page PNGs are
# derivatives of that identity, verified against recorded page hashes on every
# hit. Pre-v2 flat entries resolve as legacy hits and are adopted into the v2
# layout. Corrupt or incomplete entries never satisfy a hit: CacheOnly throws a
# corrupt-entry error, while rendering mode deletes and re-renders.

param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [switch] $CacheOnly,

    [string] $CacheVariant,

    [string] $CaseId,

    [switch] $SkipBuild,

    [int] $RenderTimeoutSeconds = 600
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "ReferenceCache.ps1")

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path

function New-CacheMissMessage($resolution) {
    $variantMessage = if ([string]::IsNullOrWhiteSpace($CacheVariant)) { "" } else { " for variant '$CacheVariant'" }
    $populate = "pwsh tools/RenderCachedReference.ps1 -InputPath '$inputFull' -OutputDirectory <output-dir> -Dpi $Dpi" +
        $(if ([string]::IsNullOrWhiteSpace($CacheVariant)) { "" } else { " -CacheVariant '$CacheVariant'" })
    if (-not [string]::IsNullOrWhiteSpace($CaseId)) {
        $populate += " -CaseId '$CaseId'"
    }

    return "Reference cache miss for '$inputFull'$variantMessage at $Dpi DPI ($($resolution.Reason)). Cache-only mode refuses to invoke Office/COM reference rendering. Expected identity directory: artifacts/reference-cache/$($resolution.IdentityKey). To populate it on an Office setup, run: $populate"
}

$resolution = Use-ReferenceCacheLock ("resolve-" + (Get-ReferenceIdentityKey $inputFull $CacheVariant)) {
    Resolve-ReferenceCacheEntry $inputFull $Dpi $CacheVariant
}

if ($resolution.Outcome -eq "LegacyHit") {
    Write-Host ("Reference cache legacy entry ($($resolution.Reason)): $($resolution.CacheDirectory)")
    $resolution = Use-ReferenceCacheLock $resolution.IdentityKey {
        Move-LegacyReferenceEntry $resolution.CacheDirectory (Join-Path (Get-ReferenceCacheRoot) $resolution.IdentityKey) $inputFull $Dpi $resolution.DerivativeName $CacheVariant $CaseId
        Resolve-ReferenceCacheEntry $inputFull $Dpi $CacheVariant
    }
}

if ($resolution.Outcome -eq "Corrupt") {
    if ($CacheOnly) {
        throw "Reference cache entry is corrupt and CacheOnly refuses to re-render: $($resolution.Reason) Entry: $($resolution.CacheDirectory)"
    }

    Write-Host ("Reference cache entry corrupt ($($resolution.Reason)); re-rendering: $($resolution.CacheDirectory)")
    Use-ReferenceCacheLock $resolution.IdentityKey {
        if (Test-Path -LiteralPath $resolution.CacheDirectory) {
            Remove-Item -LiteralPath $resolution.CacheDirectory -Recurse -Force
        }
    }
    $resolution = [pscustomobject]@{ Outcome = "Miss"; IdentityKey = $resolution.IdentityKey; CacheDirectory = $resolution.CacheDirectory; Reason = "corrupt entry removed" }
}

if ($resolution.Outcome -eq "Miss") {
    if ($CacheOnly) {
        throw (New-CacheMissMessage $resolution)
    }

    $renderReference = Join-Path $PSScriptRoot "RenderReference.ps1"
    Use-ReferenceCacheLock $resolution.IdentityKey {
        $staging = Join-Path (Get-ReferenceCacheRoot) ("_render-" + [System.Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $staging | Out-Null
        try {
            & $renderReference -InputPath $inputFull -OutputDirectory $staging -Dpi $Dpi -TimeoutSeconds $RenderTimeoutSeconds
            $workerStatus = $null
            $workerStatusPath = Join-Path $staging "reference-worker-status.json"
            if (Test-Path -LiteralPath $workerStatusPath) {
                $workerStatus = Get-Content -Raw -LiteralPath $workerStatusPath | ConvertFrom-Json
            }

            foreach ($scratch in @("reference-status.json", "reference-worker-status.json", "reference-supervisor.log", "reference-worker.log", "reference-progress.log")) {
                $scratchPath = Join-Path $staging $scratch
                if (Test-Path -LiteralPath $scratchPath) {
                    Remove-Item -LiteralPath $scratchPath -Force
                }
            }

            $null = Write-ReferenceIdentityMetadata $staging $resolution.IdentityKey $inputFull ((Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()) $CacheVariant $CaseId @{
                Producer = "office-com"
                OfficeApp = $(if ($null -ne $workerStatus) { [string]$workerStatus.OfficeApp } else { "" })
                OfficeVersion = $(if ($null -ne $workerStatus) { [string]$workerStatus.OfficeVersion } else { "" })
                ExportSettings = $(if ($null -ne $workerStatus) { [string]$workerStatus.ExportSettings } else { "" })
            } $null
            Publish-ReferenceStagedDirectory $staging $resolution.CacheDirectory
            $derivativeName = Get-ReferenceDerivativeName $Dpi
            $derivativeDir = Join-Path (Join-Path $resolution.CacheDirectory "raster") $derivativeName
            New-Item -ItemType Directory -Force -Path $derivativeDir | Out-Null
            foreach ($page in @(Get-ChildItem -LiteralPath $resolution.CacheDirectory -Filter "page-*.png" | Sort-Object Name)) {
                Move-Item -LiteralPath $page.FullName -Destination (Join-Path $derivativeDir $page.Name) -Force
            }
            Add-ReferenceDerivativeRecord $resolution.CacheDirectory $derivativeName $Dpi (Get-RasterizerId)
        }
        finally {
            if (Test-Path -LiteralPath $staging) {
                Remove-Item -LiteralPath $staging -Recurse -Force
            }
        }
    }
    $resolution = Resolve-ReferenceCacheEntry $inputFull $Dpi $CacheVariant
    if ($resolution.Outcome -ne "Hit") {
        throw "Reference render completed but the cache entry failed verification: $($resolution.Reason)"
    }
}

# At this point the identity entry is verified. Ensure the requested DPI
# derivative exists (rasterize offline from the cached PDF when missing).
$metadata = [string]$resolution.Metadata.ReferencePdfSha256
$derivative = $resolution.Derivative
if ($derivative.State -eq "Missing") {
    if ($CacheOnly) {
        # Rasterizing from the cached PDF needs only local PDFium, so a missing
        # derivative is fulfillable offline. Fall through to rasterize below.
        Write-Host ("Reference derivative '$($resolution.DerivativeName)' missing; rasterizing offline from the cached PDF.")
    }

    Use-ReferenceCacheLock $resolution.IdentityKey {
        $derivativeDir = Join-Path (Join-Path $resolution.CacheDirectory "raster") $resolution.DerivativeName
        New-Item -ItemType Directory -Force -Path $derivativeDir | Out-Null
        & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf (Join-Path $resolution.CacheDirectory "reference.pdf") -OutputDirectory $derivativeDir -Dpi $Dpi -SkipBuild:$SkipBuild
        Add-ReferenceDerivativeRecord $resolution.CacheDirectory $resolution.DerivativeName $Dpi (Get-RasterizerId)
    }
    $derivative = Test-ReferenceDerivative $resolution.CacheDirectory $resolution.DerivativeName $metadata
}

if ($derivative.State -ne "Complete") {
    throw "Reference cache derivative is unusable ($($derivative.Reason)): $($derivative.Directory)"
}

Copy-Item -LiteralPath (Join-Path $resolution.CacheDirectory "reference.pdf") -Destination (Join-Path $outputFull "reference.pdf") -Force
foreach ($page in @(Get-ChildItem -LiteralPath $derivative.Directory -Filter "page-*.png" | Sort-Object Name)) {
    Copy-Item -LiteralPath $page.FullName -Destination (Join-Path $outputFull $page.Name) -Force
}

Copy-Item -LiteralPath (Join-Path $resolution.CacheDirectory "reference-metadata.json") -Destination (Join-Path $outputFull "reference-metadata.json") -Force
Write-Host ("Reference cache: {0} derivative {1}" -f $resolution.CacheDirectory, $resolution.DerivativeName)
