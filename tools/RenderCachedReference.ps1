# Renders or copies cached Office reference PDFs.
#
# Cache identity (protocol v2) is the reference PDF itself: input hash,
# extension, variant, and rendering protocol. Rasterizer/DPI page PNGs are
# derivatives of that identity, verified against recorded page hashes on every
# hit. Flat historical caches require a trusted re-import. A v2 entry without
# verified source provenance stays on disk and requires a trusted re-import.

param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [switch] $CacheOnly,

    [string] $CacheVariant,

    [string] $CaseId,

    [switch] $SkipBuild,

    [int] $RenderTimeoutSeconds = 600,

    [string] $RasterizerPath
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "ReferenceCache.ps1")

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
$rasterArguments = @{ Framework = $Framework; Dpi = $Dpi; SkipBuild = [bool]$SkipBuild }
if ($RasterizerPath) {
    $RasterizerPath = (Resolve-Path -LiteralPath $RasterizerPath).Path
    $rasterArguments.RasterizerPath = $RasterizerPath
}
$rasterizerId = Get-RasterizerId $RasterizerPath
$derivativeName = Get-ReferenceDerivativeName $Dpi $rasterizerId
$isDocxMarkupVariant = [IO.Path]::GetExtension($inputFull) -eq '.docx' -and $CacheVariant -match '(^|;)docxMarkup='
if ((Test-Path -LiteralPath $OutputDirectory -PathType Container) -and
    @(Get-ChildItem -LiteralPath $OutputDirectory -Filter '*.png' -File).Count) {
    throw 'Reference output already contains PNG images; choose a fresh directory.'
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path

function New-CacheMissMessage($resolution) {
    $variantMessage = if ([string]::IsNullOrWhiteSpace($CacheVariant)) { "" } else { " for variant '$CacheVariant'" }
    $populate = "pwsh tools/RenderCachedReference.ps1 -InputPath '$inputFull' -OutputDirectory <output-dir> -Dpi $Dpi" +
        $(if ([string]::IsNullOrWhiteSpace($CacheVariant)) { "" } else { " -CacheVariant '$CacheVariant'" })
    if (-not [string]::IsNullOrWhiteSpace($CaseId)) {
        $populate += " -CaseId '$CaseId'"
    }
    if ($isDocxMarkupVariant) {
        $populate = 'export with tools/ExportDocxMarkupReference.ps1 using explicit markup view settings, then import with tools/ImportDocxMarkupReferenceCache.ps1 for the case manifest. CacheVariant labels identify references; they do not set Word view options.'
    }

    return "Reference cache miss for '$inputFull'$variantMessage at $Dpi DPI ($($resolution.Reason)). Cache-only mode refuses to invoke Office/COM reference rendering. Expected identity directory: artifacts/reference-cache/$($resolution.IdentityKey). To populate it on an Office setup, run: $populate"
}

$inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
$identityKey = Get-ReferenceIdentityKey $inputFull $CacheVariant $inputHash
$resolution = Use-ReferenceCacheLock $identityKey {
    # Decide and publish under the same lock. Waiting callers see the complete
    # entry, including a derivative from a publisher that just finished.
    $entry = Resolve-ReferenceCacheEntry $inputFull $Dpi $CacheVariant $derivativeName
    if ($entry.IdentityKey -cne $identityKey) {
        throw 'Reference input changed during cache lookup'
    }
    if ($entry.Outcome -eq 'Hit') { return $entry }
    if ($entry.Outcome -eq 'Unverified') {
        throw "Reference cache entry is unverified. $($entry.Reason) Entry: $($entry.CacheDirectory)"
    }

    if ($entry.Outcome -eq 'Corrupt') {
        if ($CacheOnly) {
            throw "Reference cache entry is corrupt and CacheOnly refuses to re-render: $($entry.Reason) Entry: $($entry.CacheDirectory)"
        }
        if ($isDocxMarkupVariant) {
            throw "DOCX markup reference is corrupt; preserve the entry for inspection and re-import a trusted reference with explicit markup view settings. CacheVariant does not configure Word. Entry: $($entry.CacheDirectory)"
        }
        Write-Host ("Reference cache entry corrupt ($($entry.Reason)); re-rendering: $($entry.CacheDirectory)")
    }
    else {
        if ($CacheOnly) { throw (New-CacheMissMessage $entry) }
        if ($isDocxMarkupVariant) {
            throw "DOCX markup reference cache miss: export and import a trusted reference with explicit markup view settings via ExportDocxMarkupReference.ps1 and ImportDocxMarkupReferenceCache.ps1. CacheVariant does not configure Word; generic rendering would cache the wrong view."
        }
    }

    $cacheRoot = [IO.Path]::GetFullPath((Get-ReferenceCacheRoot))
    $destination = [IO.Path]::GetFullPath($entry.CacheDirectory)
    $staging = Join-Path $cacheRoot ('_render-' + [Guid]::NewGuid().ToString('N'))
    if ($destination -cne (Join-Path $cacheRoot $identityKey) -or [IO.Path]::GetDirectoryName($staging) -cne $cacheRoot) {
        throw 'Unexpected reference publication paths'
    }

    New-Item -ItemType Directory -Path $staging | Out-Null
    try {
        $renderReference = Join-Path $PSScriptRoot 'RenderReference.ps1'
        & $renderReference -InputPath $inputFull -OutputDirectory $staging -TimeoutSeconds $RenderTimeoutSeconds @rasterArguments | Out-Host
        if ((Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash -ne $inputHash) {
            throw 'Reference input changed during rendering'
        }

        $workerStatus = $null
        $workerStatusPath = Join-Path $staging 'reference-worker-status.json'
        if (Test-Path -LiteralPath $workerStatusPath) {
            $workerStatus = Get-Content -Raw -LiteralPath $workerStatusPath | ConvertFrom-Json
        }
        foreach ($scratch in @('reference-status.json', 'reference-worker-status.json', 'reference-supervisor.log',
                'reference-worker-stdout.log', 'reference-worker-stderr.log', 'reference-progress.log')) {
            $scratchPath = Join-Path $staging $scratch
            if (Test-Path -LiteralPath $scratchPath) { Remove-Item -LiteralPath $scratchPath -Force }
        }

        $null = Write-ReferenceIdentityMetadata $staging $identityKey $inputFull $inputHash $CacheVariant $CaseId @{
            Producer = 'office-com'
            OfficeApp = [string]$workerStatus.OfficeApp
            OfficeVersion = [string]$workerStatus.OfficeVersion
            ExportSettings = [string]$workerStatus.ExportSettings
        } $null
        $derivativeDir = Join-Path $staging ('raster/' + $derivativeName)
        New-Item -ItemType Directory -Force -Path $derivativeDir | Out-Null
        foreach ($page in @(Get-ChildItem -LiteralPath $staging -Filter 'page-*.png' | Sort-Object Name)) {
            Move-Item -LiteralPath $page.FullName -Destination (Join-Path $derivativeDir $page.Name)
        }
        Add-ReferenceDerivativeRecord $staging $derivativeName $Dpi $rasterizerId
        Publish-ReferenceStagedDirectory $staging $destination
    }
    finally {
        if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
    }

    $published = Resolve-ReferenceCacheEntry $inputFull $Dpi $CacheVariant $derivativeName
    if ($published.Outcome -ne 'Hit' -or $published.IdentityKey -cne $identityKey) {
        throw "Reference render completed but the cache entry failed verification: $($published.Reason)"
    }
    return $published
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
        $current = Test-ReferenceDerivative $resolution.CacheDirectory $resolution.DerivativeName $metadata
        if ($current.State -eq 'Complete') { return }
        if ($current.State -ne 'Missing') { throw ('Reference derivative changed: ' + $current.Reason) }
        $rasterRoot = [IO.Path]::GetFullPath((Join-Path $resolution.CacheDirectory 'raster'))
        $derivativeDir = [IO.Path]::GetFullPath((Join-Path $rasterRoot $resolution.DerivativeName))
        $staging = Join-Path $rasterRoot ('_raster-' + [Guid]::NewGuid().ToString('N'))
        $previous = $derivativeDir + '.incomplete-' + [Guid]::NewGuid().ToString('N')
        foreach ($owned in @($derivativeDir, $staging, $previous)) {
            if ([IO.Path]::GetDirectoryName($owned) -ne $rasterRoot) { throw 'Unexpected raster publication target' }
        }
        New-Item -ItemType Directory -Force -Path $staging | Out-Null
        try {
            & (Join-Path $PSScriptRoot 'RasterizePdf.ps1') -InputPdf (Join-Path $resolution.CacheDirectory 'reference.pdf') -OutputDirectory $staging @rasterArguments
            if (!(Get-ReferenceRasterPages $staging).Count) { throw 'Rasterizer produced no pages' }
            # Retain interrupted derivatives as evidence, never mix their pages
            # into a fresh run. Only complete pages receive a completion record.
            if (Test-Path -LiteralPath $derivativeDir) { [IO.Directory]::Move($derivativeDir, $previous) }
            [IO.Directory]::Move($staging, $derivativeDir)
            Add-ReferenceDerivativeRecord $resolution.CacheDirectory $resolution.DerivativeName $Dpi $rasterizerId
        }
        finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force } }
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
