# Trusted DOCX markup reference import. Thin wrapper over the shared
# Import-ReferenceCacheEntry path: derives the cache variant from the visual
# case manifest (or explicit overrides) and delegates keying, verification,
# atomic publish, and provenance recording to ReferenceCache.ps1.

param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [string] $ReferencePdf,

    [string] $ReferenceDirectory,

    [ValidateSet("final", "original", "simple", "all", "simple-markup", "all-markup")]
    [string] $DocxMarkup,

    [ValidateSet("preserve", "preserve-layout", "preserve-document-layout", "reserve", "reserve-margin", "markup-margin", "reserve-markup-margin", "word", "word-compatible", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup")]
    [string] $DocxMarkupGeometry = "preserve",

    [string] $OfficeApp,

    [string] $OfficeVersion,

    [string] $ExportSettings,

    [int] $Dpi = 144,

    [switch] $Force
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

. (Join-Path $PSScriptRoot "ReferenceCache.ps1")

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
$resolvedDocxMarkup = ConvertTo-CanonicalDocxMarkup $(if (-not [string]::IsNullOrWhiteSpace($DocxMarkup)) {
        $DocxMarkup
    }
    elseif ($manifest.PSObject.Properties.Name -contains "docxMarkup") {
        [string]$manifest.docxMarkup
    }
    else {
        $null
    })
if ([string]::IsNullOrWhiteSpace($resolvedDocxMarkup)) {
    throw "DOCX markup reference cache import requires docxMarkup in the visual case or -DocxMarkup."
}

$resolvedDocxMarkupGeometry = ConvertTo-CanonicalDocxMarkupGeometry $(if (-not [string]::IsNullOrWhiteSpace($DocxMarkupGeometry)) {
        $DocxMarkupGeometry
    }
    elseif ($manifest.PSObject.Properties.Name -contains "docxMarkupGeometry") {
        [string]$manifest.docxMarkupGeometry
    }
    else {
        $null
    })
if ([string]::IsNullOrWhiteSpace($resolvedDocxMarkupGeometry)) {
    $resolvedDocxMarkupGeometry = "preserve"
}

$caseId = [string]$manifest.id
$cacheVariant = "docxMarkup={0};docxMarkupGeometry={1}" -f $resolvedDocxMarkup, $resolvedDocxMarkupGeometry
Import-ReferenceCacheEntry `
    -InputPath $inputFull `
    -ReferencePdf $ReferencePdf `
    -ReferenceDirectory $ReferenceDirectory `
    -DpiValue $Dpi `
    -CacheVariant $cacheVariant `
    -CaseId $caseId `
    -Producer @{ Producer = "import"; OfficeApp = $OfficeApp; OfficeVersion = $OfficeVersion; ExportSettings = $ExportSettings } `
    -Force:$Force
