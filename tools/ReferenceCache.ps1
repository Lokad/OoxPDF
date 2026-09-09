# Shared Office-reference cache helpers. Dot-source this file instead of
# copying key logic: cache keys carry an explicit rendering-protocol version so
# cosmetic script edits no longer orphan every cached reference. Bump
# $ReferenceCacheProtocol only when the rendered reference bytes meaningfully
# change (Office export flags, raster handling, variant semantics).
#
# Layout (protocol v2): reference PDF identity is separated from rasterizer/DPI
# derivatives. One identity directory holds reference.pdf plus
# reference-metadata.json; each rasterized derivative lives under
# raster/dpi<Dpi>-<rasterizerId>/ with page hashes bound to the PDF hash.
# Pre-v2 entries (flat page-*.png next to reference.pdf, DPI baked into the
# directory key) resolve as legacy hits and are adopted into the v2 layout.

$ReferenceCacheProtocol = "v2"
$ReferenceMetadataSchema = "ooxpdf-reference-metadata/2"

function Get-ShortSha256([byte[]] $Bytes, [int] $Length = 12) {
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = ([System.BitConverter]::ToString($sha256.ComputeHash($Bytes)) -replace "-", "").ToLowerInvariant()
        return $hash.Substring(0, [Math]::Min($Length, $hash.Length))
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-ReferenceCacheKey(
    [string] $InputPath,
    [int] $DpiValue,
    [string] $CacheVariant)
{
    # Legacy v1 key (DPI baked in). Kept only to resolve pre-v2 entries.
    $inputHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $extension = [System.IO.Path]::GetExtension($InputPath).TrimStart(".").ToLowerInvariant()
    $variantKeyPart = ""
    if (-not [string]::IsNullOrWhiteSpace($CacheVariant)) {
        $variantHash = Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($CacheVariant.Trim().ToLowerInvariant())) 12
        $variantKeyPart = "-variant" + $variantHash
    }

    "{0}-{1}-v1{2}-dpi{3}" -f $extension, $inputHash.Substring(0, 24), $variantKeyPart, $DpiValue
}

function Get-ReferenceVariantKeyPart([string] $CacheVariant) {
    if ([string]::IsNullOrWhiteSpace($CacheVariant)) {
        return ""
    }

    return "-variant" + (Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($CacheVariant.Trim().ToLowerInvariant())) 12)
}

function Get-ReferenceIdentityKey(
    [string] $InputPath,
    [string] $CacheVariant)
{
    $inputHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $extension = [System.IO.Path]::GetExtension($InputPath).TrimStart(".").ToLowerInvariant()
    "{0}-{1}-{2}{3}" -f $extension, $inputHash.Substring(0, 24), $ReferenceCacheProtocol, (Get-ReferenceVariantKeyPart $CacheVariant)
}

function Get-ReferenceCacheRoot {
    return (Join-Path (Split-Path -Parent $PSScriptRoot) "artifacts/reference-cache")
}

function Get-ReferenceCacheLockName([string] $IdentityKey) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes("ooxpdf-refcache|" + $IdentityKey)
    return "Global\OoxPdfRefCache-" + (Get-ShortSha256 $bytes 16)
}

function Use-ReferenceCacheLock([string] $IdentityKey, [scriptblock] $Action) {
    $mutex = New-Object System.Threading.Mutex($false, (Get-ReferenceCacheLockName $IdentityKey))
    try {
        if (-not $mutex.WaitOne([TimeSpan]::FromMinutes(10))) {
            throw "Timed out waiting for the reference-cache lock for '$IdentityKey'. Another publisher may be stuck; resolve it before retrying."
        }

        try {
            return & $Action
        }
        finally {
            $mutex.ReleaseMutex()
        }
    }
    finally {
        $mutex.Dispose()
    }
}

function Get-RasterizerId {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $projectDir = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer"
    $sources = @(Get-ChildItem -LiteralPath $projectDir -Recurse -Include *.cs, *.csproj -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' } |
        Sort-Object FullName)
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        foreach ($file in $sources) {
            $pathBytes = [System.Text.Encoding]::UTF8.GetBytes($file.FullName.Substring($projectDir.Length))
            $hasher.TransformBlock($pathBytes, 0, $pathBytes.Length, $null, 0) | Out-Null
            $content = [System.IO.File]::ReadAllBytes($file.FullName)
            $hasher.TransformBlock($content, 0, $content.Length, $null, 0) | Out-Null
        }

        $pdfium = Join-Path $repoRoot "tools/vendor/pdfium/win-x64/bin/pdfium.dll"
        $pdfiumBytes = if (Test-Path -LiteralPath $pdfium) {
            [System.IO.File]::ReadAllBytes($pdfium)
        }
        else {
            [System.Text.Encoding]::UTF8.GetBytes("no-pdfium-dll")
        }

        $hasher.TransformFinalBlock($pdfiumBytes, 0, $pdfiumBytes.Length) | Out-Null
        return ([System.BitConverter]::ToString($hasher.Hash) -replace "-", "").ToLowerInvariant().Substring(0, 12)
    }
    finally {
        $hasher.Dispose()
    }
}

function Get-ReferenceDerivativeName([int] $DpiValue) {
    return "dpi{0}-{1}" -f $DpiValue, (Get-RasterizerId).Substring(0, 8)
}

function Read-ReferenceMetadata([string] $CacheDir) {
    $metadataPath = Join-Path $CacheDir "reference-metadata.json"
    if (-not (Test-Path -LiteralPath $metadataPath)) {
        return $null
    }

    try {
        return Get-Content -Raw -LiteralPath $metadataPath | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Read-CompleteMarker([string] $CacheDir) {
    $markerPath = Join-Path $CacheDir "complete.txt"
    if (-not (Test-Path -LiteralPath $markerPath)) {
        return $null
    }

    $fields = @{}
    foreach ($line in (Get-Content -LiteralPath $markerPath)) {
        $match = [regex]::Match($line, '^(?<key>[A-Za-z]+)=(?<value>.*)$')
        if ($match.Success) {
            $fields[$match.Groups["key"].Value] = $match.Groups["value"].Value
        }
    }

    return $fields
}

function Test-ReferencePdfBytes([string] $PdfPath) {
    if (-not (Test-Path -LiteralPath $PdfPath)) {
        return "missing reference.pdf"
    }

    $info = Get-Item -LiteralPath $PdfPath
    if ($info.Length -lt 8) {
        return "reference.pdf is truncated ($($info.Length) bytes)"
    }

    $stream = [System.IO.File]::OpenRead($PdfPath)
    try {
        $header = New-Object byte[] 5
        if ($stream.Read($header, 0, 5) -ne 5) {
            return "reference.pdf is unreadable"
        }

        if ([System.Text.Encoding]::ASCII.GetString($header) -ne "%PDF-") {
            return "reference.pdf lacks a %PDF- header (corrupt import or partial render)"
        }
    }
    finally {
        $stream.Dispose()
    }

    return $null
}

function Test-ReferenceIdentityEntry([string] $CacheDir, [string] $InputSha256, [string] $CacheVariant) {
    $marker = Read-CompleteMarker $CacheDir
    if ($null -eq $marker) {
        return [pscustomobject]@{ State = "Incomplete"; Reason = "missing complete.txt (interrupted publish?)"; Metadata = $null }
    }

    $metadata = Read-ReferenceMetadata $CacheDir
    $pdfPath = Join-Path $CacheDir "reference.pdf"
    if ($null -eq $metadata) {
        $problem = Test-ReferencePdfBytes $pdfPath
        if ($null -ne $problem) {
            return [pscustomobject]@{ State = "Corrupt"; Reason = $problem; Metadata = $null }
        }

        return [pscustomobject]@{ State = "Legacy"; Reason = "no reference-metadata.json (pre-v2 provenance gap)"; Metadata = $null }
    }

    if ([string]$metadata.Protocol -ne $ReferenceCacheProtocol) {
        return [pscustomobject]@{ State = "Corrupt"; Reason = "metadata protocol '$($metadata.Protocol)' does not match '$ReferenceCacheProtocol'"; Metadata = $metadata }
    }

    if ([string]$metadata.InputSha256 -ne $InputSha256) {
        return [pscustomobject]@{ State = "Corrupt"; Reason = "metadata input hash does not match current input (stale entry)"; Metadata = $metadata }
    }

    if ([string]$metadata.CacheVariant -ne [string]$CacheVariant) {
        return [pscustomobject]@{ State = "Corrupt"; Reason = "metadata variant does not match requested variant"; Metadata = $metadata }
    }

    $problem = Test-ReferencePdfBytes $pdfPath
    if ($null -ne $problem) {
        return [pscustomobject]@{ State = "Corrupt"; Reason = $problem; Metadata = $metadata }
    }

    $actualPdfHash = (Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ([string]$metadata.ReferencePdfSha256 -ne $actualPdfHash) {
        return [pscustomobject]@{ State = "Corrupt"; Reason = "reference.pdf hash mismatch (expected $($metadata.ReferencePdfSha256), got $actualPdfHash)"; Metadata = $metadata }
    }

    return [pscustomobject]@{ State = "Complete"; Reason = $null; Metadata = $metadata }
}

function Test-ReferenceDerivative([string] $CacheDir, [string] $DerivativeName, [string] $PdfSha256) {
    $derivativeDir = Join-Path (Join-Path $CacheDir "raster") $DerivativeName
    $pages = @(Get-ChildItem -LiteralPath $derivativeDir -Filter "page-*.png" -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($pages.Count -eq 0) {
        return [pscustomobject]@{ State = "Missing"; Reason = "no raster pages for derivative '$DerivativeName'"; Directory = $derivativeDir }
    }

    $metadata = Read-ReferenceMetadata $CacheDir
    $recorded = $null
    if ($null -ne $metadata -and $null -ne $metadata.Derivatives) {
        $recorded = $metadata.Derivatives.$DerivativeName
    }

    if ($null -ne $recorded) {
        if ([string]$recorded.PdfSha256 -ne $PdfSha256) {
            return [pscustomobject]@{ State = "Corrupt"; Reason = "derivative '$DerivativeName' was rasterized from a different reference.pdf"; Directory = $derivativeDir }
        }

        $expected = @($recorded.Pages | ForEach-Object { $_.Sha256 })
        $actual = @($pages | ForEach-Object { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() })
        if (($expected -join ",") -ne ($actual -join ",")) {
            return [pscustomobject]@{ State = "Corrupt"; Reason = "derivative '$DerivativeName' page hashes do not match recorded metadata"; Directory = $derivativeDir }
        }
    }

    return [pscustomobject]@{ State = "Complete"; Reason = $null; Directory = $derivativeDir }
}

function Find-PreV1LegacyEntry([string] $InputPath, [int] $DpiValue, [string] $CacheVariant) {
    # Pre-v1 keys baked an irreproducible tool hash into the directory name, so
    # fall back to scanning complete.txt markers. Only used when v2/v1 lookups miss.
    $cacheRoot = Get-ReferenceCacheRoot
    if (-not (Test-Path -LiteralPath $cacheRoot)) {
        return $null
    }

    $inputFull = (Resolve-Path -LiteralPath $InputPath).Path
    $extension = [System.IO.Path]::GetExtension($inputFull).TrimStart(".").ToLowerInvariant()
    $candidates = @(Get-ChildItem -LiteralPath $cacheRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "$extension-*-dpi$DpiValue" } |
        Sort-Object Name)
    foreach ($candidate in $candidates) {
        $marker = Read-CompleteMarker $candidate.FullName
        if ($null -eq $marker -or $marker.ContainsKey("protocol")) {
            continue
        }

        $matchesInput = $false
        if ($marker.ContainsKey("inputSha256")) {
            $inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
            $matchesInput = ([string]$marker["inputSha256"] -eq $inputHash)
        }
        elseif ($marker.ContainsKey("input")) {
            try {
                $matchesInput = ([System.IO.Path]::GetFullPath($marker["input"]) -eq $inputFull)
            }
            catch {
                $matchesInput = $false
            }
        }

        if (-not $matchesInput -or [string]$marker["variant"] -ne [string]$CacheVariant) {
            continue
        }

        if (Test-Path -LiteralPath (Join-Path $candidate.FullName "reference.pdf")) {
            return $candidate.FullName
        }
    }

    return $null
}

function Resolve-ReferenceCacheEntry([string] $InputPath, [int] $DpiValue, [string] $CacheVariant) {
    $inputFull = (Resolve-Path -LiteralPath $InputPath).Path
    $inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $cacheRoot = Get-ReferenceCacheRoot
    $identityKey = Get-ReferenceIdentityKey $inputFull $CacheVariant
    $cacheDir = Join-Path $cacheRoot $identityKey
    $derivativeName = Get-ReferenceDerivativeName $DpiValue

    if (Test-Path -LiteralPath $cacheDir) {
        $entry = Test-ReferenceIdentityEntry $cacheDir $inputHash $CacheVariant
        if ($entry.State -eq "Complete") {
            $pdfHash = [string]$entry.Metadata.ReferencePdfSha256
            return [pscustomobject]@{
                Outcome = "Hit"
                IdentityKey = $identityKey
                CacheDirectory = $cacheDir
                DerivativeName = $derivativeName
                Derivative = (Test-ReferenceDerivative $cacheDir $derivativeName $pdfHash)
                Metadata = $entry.Metadata
            }
        }

        if ($entry.State -eq "Corrupt") {
            return [pscustomobject]@{
                Outcome = "Corrupt"
                IdentityKey = $identityKey
                CacheDirectory = $cacheDir
                Reason = $entry.Reason
                Metadata = $entry.Metadata
            }
        }

        if ($entry.State -eq "Legacy") {
            return [pscustomobject]@{
                Outcome = "LegacyHit"
                IdentityKey = $identityKey
                CacheDirectory = $cacheDir
                DerivativeName = $derivativeName
                Reason = $entry.Reason
                Metadata = $null
            }
        }

        return [pscustomobject]@{ Outcome = "Miss"; IdentityKey = $identityKey; CacheDirectory = $cacheDir; Reason = $entry.Reason }
    }

    $legacyV1Dir = Join-Path $cacheRoot (Get-ReferenceCacheKey $inputFull $DpiValue $CacheVariant)
    if ((Test-Path -LiteralPath (Join-Path $legacyV1Dir "complete.txt")) -and (Test-Path -LiteralPath (Join-Path $legacyV1Dir "reference.pdf"))) {
        return [pscustomobject]@{
            Outcome = "LegacyHit"
            IdentityKey = $identityKey
            CacheDirectory = $legacyV1Dir
            DerivativeName = $derivativeName
            Reason = "v1 flat entry (DPI baked into key, limited provenance)"
            Metadata = (Read-ReferenceMetadata $legacyV1Dir)
        }
    }

    $preV1 = Find-PreV1LegacyEntry $inputFull $DpiValue $CacheVariant
    if ($null -ne $preV1) {
        return [pscustomobject]@{
            Outcome = "LegacyHit"
            IdentityKey = $identityKey
            CacheDirectory = $preV1
            DerivativeName = $derivativeName
            Reason = "pre-v1 flat entry (tool-hash key, provenance gap)"
            Metadata = $null
        }
    }

    return [pscustomobject]@{ Outcome = "Miss"; IdentityKey = $identityKey; CacheDirectory = (Join-Path $cacheRoot $identityKey); Reason = "no complete entry" }
}

function Write-ReferenceIdentityMetadata(
    [string] $StagedDirectory,
    [string] $IdentityKey,
    [string] $InputPath,
    [string] $InputSha256,
    [string] $CacheVariant,
    [string] $CaseId,
    [hashtable] $Producer,
    [string] $ProvenanceNote) {
    $pdfPath = Join-Path $StagedDirectory "reference.pdf"
    $problem = Test-ReferencePdfBytes $pdfPath
    if ($null -ne $problem) {
        throw "Refusing to publish reference cache entry: $problem."
    }

    $pdfHash = (Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $metadata = [ordered]@{
        Schema = $ReferenceMetadataSchema
        Protocol = $ReferenceCacheProtocol
        CaseId = $CaseId
        InputExtension = [System.IO.Path]::GetExtension($InputPath).TrimStart(".").ToLowerInvariant()
        InputSha256 = $InputSha256
        CacheVariant = $CacheVariant
        IdentityKey = $IdentityKey
        ReferencePdfLength = (Get-Item -LiteralPath $pdfPath).Length
        ReferencePdfSha256 = $pdfHash
        Producer = [string]$Producer.Producer
        OfficeApp = $Producer.OfficeApp
        OfficeVersion = $Producer.OfficeVersion
        ExportSettings = $Producer.ExportSettings
        ProvenanceNote = $ProvenanceNote
        CreatedAtUtc = [DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)
        Derivatives = [ordered]@{}
    }

    $metadata | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $StagedDirectory "reference-metadata.json") -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $StagedDirectory "complete.txt") -Value (
        "protocol={0}`ninputSha256={1}`nvariant={2}`nreferencePdfSha256={3}`n" -f $ReferenceCacheProtocol, $InputSha256, $CacheVariant, $pdfHash) -NoNewline
    return $metadata
}

function Add-ReferenceDerivativeRecord(
    [string] $CacheDir,
    [string] $DerivativeName,
    [int] $DpiValue,
    [string] $RasterizerId) {
    $derivativeDir = Join-Path (Join-Path $CacheDir "raster") $DerivativeName
    $pages = @(Get-ChildItem -LiteralPath $derivativeDir -Filter "page-*.png" -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($pages.Count -eq 0) {
        throw "Cannot record derivative '$DerivativeName': no page-*.png raster output."
    }

    $pdfHash = (Get-FileHash -LiteralPath (Join-Path $CacheDir "reference.pdf") -Algorithm SHA256).Hash.ToLowerInvariant()
    $metadata = Read-ReferenceMetadata $CacheDir
    if ($null -eq $metadata) {
        throw "Cannot record derivative '$DerivativeName': identity metadata is missing."
    }

    $pageRecords = @($pages | ForEach-Object {
        [ordered]@{
            Name = $_.Name
            Length = $_.Length
            Sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    })
    if ($null -eq $metadata.Derivatives) {
        $metadata | Add-Member -NotePropertyName "Derivatives" -NotePropertyValue ([ordered]@{}) -Force
    }

    $metadata.Derivatives | Add-Member -NotePropertyName $DerivativeName -NotePropertyValue ([ordered]@{
        Dpi = $DpiValue
        RasterizerId = $RasterizerId
        PdfSha256 = $pdfHash
        PageCount = $pages.Count
        Pages = $pageRecords
    }) -Force
    $metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $CacheDir "reference-metadata.json") -Encoding UTF8
}

function Publish-ReferenceStagedDirectory([string] $StagedDirectory, [string] $CacheDirectory) {
    $problem = Test-ReferencePdfBytes (Join-Path $StagedDirectory "reference.pdf")
    if ($null -ne $problem) {
        throw "Refusing to publish reference cache entry: $problem."
    }

    if (-not (Test-Path -LiteralPath (Join-Path $StagedDirectory "complete.txt"))) {
        throw "Refusing to publish reference cache entry: staged complete.txt is missing."
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $CacheDirectory) | Out-Null
    if (Test-Path -LiteralPath $CacheDirectory) {
        Remove-Item -LiteralPath $CacheDirectory -Recurse -Force
    }

    Move-Item -LiteralPath $StagedDirectory -Destination $CacheDirectory
}

function Move-LegacyReferenceEntry(
    [string] $LegacyDirectory,
    [string] $CacheDirectory,
    [string] $InputPath,
    [int] $DpiValue,
    [string] $DerivativeName,
    [string] $CacheVariant,
    [string] $CaseId) {
    # Adopt a flat legacy entry into the v2 layout. The PDF becomes the identity;
    # existing same-DPI rasters move under the derivative directory with an
    # honest legacy-provenance note (their exact rasterizer is unverifiable).
    $pdfSource = Join-Path $LegacyDirectory "reference.pdf"
    $problem = Test-ReferencePdfBytes $pdfSource
    if ($null -ne $problem) {
        throw "Cannot adopt legacy reference entry: $problem."
    }

    $inputHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $staging = Join-Path (Split-Path -Parent $CacheDirectory) ("_adopt-" + [System.Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $staging | Out-Null
    try {
        Copy-Item -LiteralPath $pdfSource -Destination (Join-Path $staging "reference.pdf") -Force
        $derivativeDir = Join-Path (Join-Path $staging "raster") $DerivativeName
        New-Item -ItemType Directory -Force -Path $derivativeDir | Out-Null
        foreach ($page in @(Get-ChildItem -LiteralPath $LegacyDirectory -Filter "page-*.png" -ErrorAction SilentlyContinue | Sort-Object Name)) {
            Copy-Item -LiteralPath $page.FullName -Destination (Join-Path $derivativeDir $page.Name) -Force
        }

        $legacyMetadata = Read-ReferenceMetadata $LegacyDirectory
        $adoptedCaseId = $CaseId
        if ([string]::IsNullOrWhiteSpace($adoptedCaseId) -and $null -ne $legacyMetadata -and -not [string]::IsNullOrWhiteSpace([string]$legacyMetadata.CaseId)) {
            $adoptedCaseId = [string]$legacyMetadata.CaseId
        }

        $identityKey = Split-Path -Leaf $CacheDirectory
        $null = Write-ReferenceIdentityMetadata $staging $identityKey $InputPath $inputHash $CacheVariant $adoptedCaseId @{ Producer = "legacy-adopted" } ("adopted from flat legacy entry " + (Split-Path -Leaf $LegacyDirectory))
        Publish-ReferenceStagedDirectory $staging $CacheDirectory
        Add-ReferenceDerivativeRecord $CacheDirectory $DerivativeName $DpiValue ("legacy-" + (Get-RasterizerId).Substring(0, 8))
        if ((Resolve-Path -LiteralPath $LegacyDirectory).Path -ne (Resolve-Path -LiteralPath $CacheDirectory).Path -and (Test-Path -LiteralPath $LegacyDirectory)) {
            Remove-Item -LiteralPath $LegacyDirectory -Recurse -Force
        }

        Write-Host ("Adopted legacy reference entry into v2 layout: {0}" -f $CacheDirectory)
    }
    finally {
        if (Test-Path -LiteralPath $staging) {
            Remove-Item -LiteralPath $staging -Recurse -Force
        }
    }
}

function Import-ReferenceCacheEntry(
    [string] $InputPath,
    [string] $ReferencePdf,
    [string] $ReferenceDirectory,
    [int] $DpiValue,
    [string] $CacheVariant,
    [string] $CaseId,
    [hashtable] $Producer,
    [switch] $Force) {
    # Shared trusted-import path for ordinary DOCX/PPTX and markup-variant
    # references alike. The caller vouches for the reference PDF; this function
    # verifies structure, binds raster derivatives to the PDF hash, publishes
    # atomically, and re-verifies the published entry.
    $inputFull = (Resolve-Path -LiteralPath $InputPath).Path
    $inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $identityKey = Get-ReferenceIdentityKey $inputFull $CacheVariant
    $derivativeName = Get-ReferenceDerivativeName $DpiValue
    $cacheRoot = Get-ReferenceCacheRoot
    $cacheDir = Join-Path $cacheRoot $identityKey
    return Use-ReferenceCacheLock $identityKey {
        if ((Test-Path -LiteralPath (Join-Path $cacheDir "complete.txt")) -and -not $Force) {
            throw "Reference cache already exists: $cacheDir. Pass -Force to replace it."
        }

        $staging = Join-Path $cacheRoot ("_import-" + [System.Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $staging | Out-Null
        try {
            if (-not [string]::IsNullOrWhiteSpace($ReferencePdf)) {
                Copy-Item -LiteralPath (Resolve-Path -LiteralPath $ReferencePdf).Path -Destination (Join-Path $staging "reference.pdf") -Force
            }
            else {
                $sourceFull = (Resolve-Path -LiteralPath $ReferenceDirectory).Path
                if (-not (Test-Path -LiteralPath (Join-Path $sourceFull "reference.pdf"))) {
                    throw "Reference directory does not contain reference.pdf: $sourceFull"
                }

                Copy-Item -Path (Join-Path $sourceFull "*") -Destination $staging -Recurse -Force
            }

            $stagedPdf = Join-Path $staging "reference.pdf"
            $problem = Test-ReferencePdfBytes $stagedPdf
            if ($null -ne $problem) {
                throw "Imported reference is unusable: $problem."
            }

            $derivativeDir = Join-Path (Join-Path $staging "raster") $derivativeName
            $stagedPages = @(Get-ChildItem -LiteralPath $staging -Filter "page-*.png" -ErrorAction SilentlyContinue)
            if ($stagedPages.Count -eq 0) {
                New-Item -ItemType Directory -Force -Path $derivativeDir | Out-Null
                & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $stagedPdf -OutputDirectory $derivativeDir -Dpi $DpiValue
            }
            else {
                New-Item -ItemType Directory -Force -Path $derivativeDir | Out-Null
                foreach ($page in ($stagedPages | Sort-Object Name)) {
                    Move-Item -LiteralPath $page.FullName -Destination (Join-Path $derivativeDir $page.Name) -Force
                }
            }

            foreach ($leftover in @(Get-ChildItem -LiteralPath $staging -Filter "page-*.png" -ErrorAction SilentlyContinue)) {
                Remove-Item -LiteralPath $leftover.FullName -Force
            }

            $null = Write-ReferenceIdentityMetadata $staging $identityKey $inputFull $inputHash $CacheVariant $CaseId $Producer $null
            Publish-ReferenceStagedDirectory $staging $cacheDir
            Add-ReferenceDerivativeRecord $cacheDir $derivativeName $DpiValue (Get-RasterizerId)
            $check = Test-ReferenceIdentityEntry $cacheDir $inputHash $CacheVariant
            if ($check.State -ne "Complete") {
                throw "Imported reference failed post-publish verification: $($check.Reason)"
            }

            Write-Host "Imported Office reference cache: $cacheDir"
            return $cacheDir
        }
        finally {
            if (Test-Path -LiteralPath $staging) {
                Remove-Item -LiteralPath $staging -Recurse -Force
            }
        }
    }
}
