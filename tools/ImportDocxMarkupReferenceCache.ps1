param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [string] $ReferencePdf,

    [string] $ReferenceDirectory,

    [ValidateSet("final", "original", "simple", "all", "simple-markup", "all-markup")]
    [string] $DocxMarkup,

    [ValidateSet("preserve", "preserve-layout", "preserve-document-layout", "reserve", "reserve-margin", "markup-margin", "reserve-markup-margin", "word", "word-compatible", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup")]
    [string] $DocxMarkupGeometry,

    [int] $Dpi = 144,

    [switch] $Force
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

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
    $inputHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $renderReference = Join-Path $PSScriptRoot "RenderReference.ps1"
    $rasterizePdf = Join-Path $PSScriptRoot "RasterizePdf.ps1"
    $toolHashSource = @(
        (Get-FileHash -LiteralPath $renderReference -Algorithm SHA256).Hash
        (Get-FileHash -LiteralPath $rasterizePdf -Algorithm SHA256).Hash
    ) -join "|"
    $toolHash = Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($toolHashSource)) 12
    $extension = [System.IO.Path]::GetExtension($InputPath).TrimStart(".").ToLowerInvariant()
    $variantKeyPart = ""
    if (-not [string]::IsNullOrWhiteSpace($CacheVariant)) {
        $variantHash = Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($CacheVariant.Trim().ToLowerInvariant())) 12
        $variantKeyPart = "-variant" + $variantHash
    }

    "{0}-{1}-{2}{3}-dpi{4}" -f $extension, $inputHash.Substring(0, 24), $toolHash, $variantKeyPart, $DpiValue
}

function Copy-ReferenceDirectory([string] $SourceDirectory, [string] $DestinationDirectory) {
    if (-not (Test-Path -LiteralPath (Join-Path $SourceDirectory "reference.pdf"))) {
        throw "Reference directory does not contain reference.pdf: $SourceDirectory"
    }

    Copy-Item -Path (Join-Path $SourceDirectory "*") -Destination $DestinationDirectory -Recurse -Force
}

function ConvertTo-CanonicalDocxMarkup([string] $Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    switch ($Value.ToLowerInvariant()) {
        "simple-markup" { return "simple" }
        "all-markup" { return "all" }
        default { return $Value.ToLowerInvariant() }
    }
}

function ConvertTo-CanonicalDocxMarkupGeometry([string] $Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    switch ($Value.ToLowerInvariant()) {
        { $_ -in @("preserve-layout", "preserve-document-layout") } { return "preserve" }
        { $_ -in @("reserve", "reserve-margin", "markup-margin", "reserve-markup-margin") } { return "reserve-margin" }
        { $_ -in @("word", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup") } { return "word-compatible" }
        default { return $Value.ToLowerInvariant() }
    }
}

if (-not [string]::IsNullOrWhiteSpace($ReferencePdf) -and -not [string]::IsNullOrWhiteSpace($ReferenceDirectory)) {
    throw "Use either -ReferencePdf or -ReferenceDirectory, not both."
}

if ([string]::IsNullOrWhiteSpace($ReferencePdf) -and [string]::IsNullOrWhiteSpace($ReferenceDirectory)) {
    throw "Provide either -ReferencePdf or -ReferenceDirectory."
}

$caseFull = (Resolve-Path -LiteralPath $Case).Path
$caseDirectory = Split-Path -Parent $caseFull
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
if ($manifest.kind -ne "docx") {
    throw "DOCX markup reference cache import only supports DOCX visual cases."
}

$inputFull = (Resolve-Path -LiteralPath (Join-Path $caseDirectory $manifest.input)).Path
$docxMarkup = ConvertTo-CanonicalDocxMarkup $(if (-not [string]::IsNullOrWhiteSpace($DocxMarkup)) {
        $DocxMarkup
    }
    elseif ($manifest.PSObject.Properties.Name -contains "docxMarkup") {
        [string]$manifest.docxMarkup
    }
    else {
        $null
    })
if ([string]::IsNullOrWhiteSpace($docxMarkup)) {
    throw "DOCX markup reference cache import requires docxMarkup in the visual case or -DocxMarkup."
}

$docxMarkupGeometry = ConvertTo-CanonicalDocxMarkupGeometry $(if (-not [string]::IsNullOrWhiteSpace($DocxMarkupGeometry)) {
        $DocxMarkupGeometry
    }
    elseif ($manifest.PSObject.Properties.Name -contains "docxMarkupGeometry") {
        [string]$manifest.docxMarkupGeometry
    }
    else {
        $null
    })
if ([string]::IsNullOrWhiteSpace($docxMarkupGeometry)) {
    $docxMarkupGeometry = "preserve"
}

$caseId = if ($manifest.PSObject.Properties.Name -contains "docxMarkup" -and -not [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
    [string]$manifest.id
}
else {
    "private-docx-markup-{0}" -f $docxMarkup
}
$cacheVariant = "docxMarkup={0};docxMarkupGeometry={1}" -f $docxMarkup, $docxMarkupGeometry
$cacheKey = Get-ReferenceCacheKey $inputFull $Dpi $cacheVariant
$cacheRoot = Join-Path $repoRoot "artifacts/reference-cache"
$cacheDir = Join-Path $cacheRoot $cacheKey
$completeMarker = Join-Path $cacheDir "complete.txt"

if ((Test-Path -LiteralPath $completeMarker) -and -not $Force) {
    throw "Reference cache already exists: $cacheDir. Pass -Force to replace it."
}

$tempDir = Join-Path $cacheRoot ("_import-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
try {
    if (-not [string]::IsNullOrWhiteSpace($ReferencePdf)) {
        Copy-Item -LiteralPath (Resolve-Path -LiteralPath $ReferencePdf).Path -Destination (Join-Path $tempDir "reference.pdf") -Force
    }
    else {
        $referenceDirectoryFull = (Resolve-Path -LiteralPath $ReferenceDirectory).Path
        Copy-ReferenceDirectory $referenceDirectoryFull $tempDir
    }

    $referencePdfPath = Join-Path $tempDir "reference.pdf"
    if (-not (Test-Path -LiteralPath $referencePdfPath)) {
        throw "Imported reference did not provide reference.pdf."
    }

    $rasterPages = @(Get-ChildItem -LiteralPath $tempDir -Filter "page-*.png" -ErrorAction SilentlyContinue)
    if ($rasterPages.Count -eq 0) {
        & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdfPath -OutputDirectory $tempDir -Dpi $Dpi
    }

    $referencePdfHash = (Get-FileHash -LiteralPath $referencePdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $pageHashes = @(
        Get-ChildItem -LiteralPath $tempDir -Filter "page-*.png" -ErrorAction SilentlyContinue |
            Sort-Object Name |
            ForEach-Object {
                [pscustomobject]@{
                    Name = $_.Name
                    Length = $_.Length
                    Sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
                }
            }
    )

    $metadata = [ordered]@{
        CaseId = $caseId
        InputExtension = [System.IO.Path]::GetExtension($inputFull).TrimStart(".").ToLowerInvariant()
        InputSha256 = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
        Dpi = $Dpi
        CacheVariant = $cacheVariant
        CacheKey = $cacheKey
        ReferencePdfLength = (Get-Item -LiteralPath $referencePdfPath).Length
        ReferencePdfSha256 = $referencePdfHash
        RasterPageCount = $pageHashes.Count
        RasterPages = $pageHashes
        ImportedAtUtc = [DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)
    }
    $metadata | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $tempDir "reference-metadata.json") -Encoding UTF8

    New-Item -ItemType Directory -Force -Path $cacheRoot | Out-Null
    if (Test-Path -LiteralPath $cacheDir) {
        Remove-Item -LiteralPath $cacheDir -Recurse -Force
    }

    Move-Item -LiteralPath $tempDir -Destination $cacheDir
    Set-Content -LiteralPath (Join-Path $cacheDir "complete.txt") -Value ("inputSha256={0}`ndpi={1}`nvariant={2}`nreferencePdfSha256={3}`n" -f $metadata.InputSha256, $Dpi, $cacheVariant, $referencePdfHash)
}
finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}

Write-Host "Imported DOCX markup reference cache: $cacheDir"
