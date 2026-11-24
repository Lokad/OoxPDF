# Shared Office-reference cache helpers. Dot-source this file instead of
# copying key logic: cache keys carry an explicit rendering-protocol version so
# cosmetic script edits no longer orphan every cached reference. Bump
# $ReferenceCacheProtocol only when the rendered reference bytes meaningfully
# change (Office export flags or variant semantics).
#
# Layout (protocol v2): reference PDF identity is separated from rasterizer/DPI
# derivatives. One identity directory holds reference.pdf plus
# reference-metadata.json; each rasterized derivative lives under
# raster/dpi<Dpi>-<rasterizerId>/ with page hashes bound to the PDF hash.
# Cache hits require source-bound v2 metadata. Use the trusted import commands
# to recover an independently verified historical reference.

$ReferenceCacheProtocol = "v2"
$ReferenceMetadataSchema = "ooxpdf-reference-metadata/2"
. (Join-Path $PSScriptRoot 'PathHelpers.ps1')

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

function Get-ReferenceVariantKeyPart([string] $CacheVariant) {
    if ([string]::IsNullOrWhiteSpace($CacheVariant)) {
        return ""
    }

    return "-variant" + (Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($CacheVariant.Trim().ToLowerInvariant())) 12)
}

function Get-ReferenceIdentityKey(
    [string] $InputPath,
    [string] $CacheVariant,
    [ValidatePattern('^[0-9a-fA-F]{64}$')][string] $InputSha256)
{
    if (!$InputSha256) { $InputSha256 = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash }
    $extension = [System.IO.Path]::GetExtension($InputPath).TrimStart(".").ToLowerInvariant()
    "{0}-{1}-{2}{3}" -f $extension, $InputSha256.Substring(0, 24).ToLowerInvariant(), $ReferenceCacheProtocol, (Get-ReferenceVariantKeyPart $CacheVariant)
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

function Get-RasterizerId([string] $RasterizerPath) {
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

        $pdfium = if ($RasterizerPath) {
            $RasterizerPath = (Resolve-Path -LiteralPath $RasterizerPath).Path
            Get-RasterizerNativePath $RasterizerPath
        } else { Join-Path $repoRoot "tools/vendor/pdfium/win-x64/bin/pdfium.dll" }
        $pdfiumBytes = if (Test-Path -LiteralPath $pdfium -PathType Leaf) {
            [System.IO.File]::ReadAllBytes($pdfium)
        }
        else {
            [System.Text.Encoding]::UTF8.GetBytes("no-pdfium-dll")
        }

        if ($RasterizerPath) {
            $assemblyHash = (Get-FileHash -LiteralPath $RasterizerPath -Algorithm SHA256).Hash
            $assemblyBytes = [System.Text.Encoding]::ASCII.GetBytes($assemblyHash)
            $hasher.TransformBlock($assemblyBytes, 0, $assemblyBytes.Length, $null, 0) | Out-Null
        }
        $hasher.TransformFinalBlock($pdfiumBytes, 0, $pdfiumBytes.Length) | Out-Null
        return ([System.BitConverter]::ToString($hasher.Hash) -replace "-", "").ToLowerInvariant().Substring(0, 12)
    }
    finally {
        $hasher.Dispose()
    }
}

function Get-ReferenceDerivativeName([int] $DpiValue, [string] $RasterizerId) {
    if (!$RasterizerId) { $RasterizerId = Get-RasterizerId }
    return "dpi{0}-{1}" -f $DpiValue, $RasterizerId.Substring(0, 8)
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
    if (!(Test-Path -LiteralPath (Join-Path $CacheDir "complete.txt") -PathType Leaf)) {
        return [pscustomobject]@{ State = "Incomplete"; Reason = "missing complete.txt (interrupted publish?)"; Metadata = $null }
    }

    $metadata = Read-ReferenceMetadata $CacheDir
    # The old adopter recorded the hash at migration time. That hash does not
    # establish which source produced the earlier reference.
    if ($null -eq $metadata -or $metadata.Producer -eq 'legacy-adopted') {
        return [pscustomobject]@{
            State = "Unverified"
            Reason = "Missing source provenance. Preserve this entry and re-import a trusted reference for the exact input and view."
            Metadata = $metadata
        }
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

    $pdfPath = Join-Path $CacheDir "reference.pdf"
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

function Copy-VerifiedCachedReferencePdf([string] $InputPath, [string] $OutputDirectory, [string] $CacheVariant) {
    $inputHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash
    $key = Get-ReferenceIdentityKey $InputPath $CacheVariant $inputHash
    $cache = Join-Path (Get-ReferenceCacheRoot) $key
    Use-ReferenceCacheLock $key {
        $entry = Test-ReferenceIdentityEntry $cache $inputHash $CacheVariant
        if ($entry.State -ne 'Complete') {
            throw "A verified cached reference PDF is required for '$key' ($($entry.State)). $($entry.Reason)"
        }

        $pdf = Join-Path $OutputDirectory 'reference.pdf'
        Copy-Item -LiteralPath (Join-Path $cache 'reference.pdf') -Destination $pdf -Force
        if ((Get-FileHash -LiteralPath $pdf -Algorithm SHA256).Hash -ne $entry.Metadata.ReferencePdfSha256) {
            throw 'Cached reference PDF changed during copy.'
        }
        Copy-Item -LiteralPath (Join-Path $cache 'reference-metadata.json') -Destination $OutputDirectory -Force
    }
}

function Get-ReferenceRasterPages([string] $Directory) {
    $pages = @(Get-ChildItem -LiteralPath $Directory -Filter 'page-*.png' -File -ErrorAction SilentlyContinue)
    if (@($pages | Where-Object Name -CNotMatch '^page-[0-9]+\.png$').Count) { throw 'Invalid raster page name' }
    $pages = @($pages | Sort-Object { [int]$_.BaseName.Substring(5) })
    for ($index = 0; $index -lt $pages.Count; $index++) {
        if ($pages[$index].Name -cne ('page-{0:000}.png' -f ($index + 1))) { throw 'Incomplete raster page sequence' }
    }
    return ,$pages
}

function Test-ReferenceDerivative([string] $CacheDir, [string] $DerivativeName, [string] $PdfSha256) {
    $derivativeDir = Join-Path (Join-Path $CacheDir "raster") $DerivativeName
    $metadata = Read-ReferenceMetadata $CacheDir
    $recorded = $null
    if ($null -ne $metadata -and $null -ne $metadata.Derivatives) {
        $recorded = $metadata.Derivatives.$DerivativeName
    }

    if ($null -eq $recorded) {
        return [pscustomobject]@{ State = 'Missing'; Reason = "derivative '$DerivativeName' has no completion record"; Directory = $derivativeDir }
    }
    if ([string]$recorded.PdfSha256 -ne $PdfSha256) {
        return [pscustomobject]@{ State = 'Corrupt'; Reason = "derivative '$DerivativeName' was rasterized from a different reference.pdf"; Directory = $derivativeDir }
    }
    $tool = ([string]$recorded.RasterizerId) -replace '^legacy-', ''
    if ($tool -cnotmatch '^[0-9a-f]{8,64}$' -or $recorded.Dpi -le 0 -or
        $DerivativeName -cne ('dpi{0}-{1}' -f $recorded.Dpi, $tool.Substring(0, 8))) {
        return [pscustomobject]@{ State = 'Corrupt'; Reason = "derivative '$DerivativeName' DPI/tool identity differs"; Directory = $derivativeDir }
    }
    try { $pages = Get-ReferenceRasterPages $derivativeDir }
    catch { return [pscustomobject]@{ State = 'Corrupt'; Reason = $_.Exception.Message; Directory = $derivativeDir } }
    if (!$pages.Count) {
        return [pscustomobject]@{ State = 'Missing'; Reason = "no raster pages for derivative '$DerivativeName'"; Directory = $derivativeDir }
    }
    if ($recorded.PageCount -ne $pages.Count -or @($recorded.Pages).Count -ne $pages.Count) {
        return [pscustomobject]@{ State = 'Corrupt'; Reason = "derivative '$DerivativeName' page count differs"; Directory = $derivativeDir }
    }
    for ($index = 0; $index -lt $pages.Count; $index++) {
        if ($recorded.Pages[$index].Name -cne $pages[$index].Name -or
            $recorded.Pages[$index].Sha256 -ne (Get-FileHash -LiteralPath $pages[$index].FullName).Hash) {
            return [pscustomobject]@{ State = 'Corrupt'; Reason = "derivative '$DerivativeName' page identity differs"; Directory = $derivativeDir }
        }
    }

    return [pscustomobject]@{ State = "Complete"; Reason = $null; Directory = $derivativeDir }
}

function Resolve-ReferenceCacheEntry([string] $InputPath, [int] $DpiValue, [string] $CacheVariant, [string] $DerivativeName) {
    $inputFull = (Resolve-Path -LiteralPath $InputPath).Path
    $inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $cacheRoot = Get-ReferenceCacheRoot
    $identityKey = Get-ReferenceIdentityKey $inputFull $CacheVariant $inputHash
    $cacheDir = Join-Path $cacheRoot $identityKey
    if (!$DerivativeName) { $DerivativeName = Get-ReferenceDerivativeName $DpiValue }

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

        if ($entry.State -in @("Corrupt", "Unverified")) {
            return [pscustomobject]@{
                Outcome = $entry.State
                IdentityKey = $identityKey
                CacheDirectory = $cacheDir
                Reason = $entry.Reason
                Metadata = $entry.Metadata
            }
        }

        return [pscustomobject]@{ Outcome = "Miss"; IdentityKey = $identityKey; CacheDirectory = $cacheDir; Reason = $entry.Reason }
    }

    return [pscustomobject]@{ Outcome = "Miss"; IdentityKey = $identityKey; CacheDirectory = $cacheDir; Reason = "no complete entry" }
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
    $pages = Get-ReferenceRasterPages $derivativeDir
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
    $metadataPath = Join-Path $CacheDir 'reference-metadata.json'
    $stagingPath = $metadataPath + '.tmp-' + [Guid]::NewGuid().ToString('N')
    try {
        $metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $stagingPath -Encoding utf8
        [IO.File]::Move($stagingPath, $metadataPath, $true)
    }
    finally { if (Test-Path -LiteralPath $stagingPath) { Remove-Item -LiteralPath $stagingPath -Force } }
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
    $identityKey = Get-ReferenceIdentityKey $inputFull $CacheVariant $inputHash
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
