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
    [string] $DocxMarkupGeometry,

    [string] $OfficeApp,

    [string] $OfficeVersion,

    [string] $ExportSettings,

    [int] $Dpi = 144,

    [switch] $Force
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "ReferenceCache.ps1")

. (Join-Path $PSScriptRoot "DocxMarkupReference.ps1")

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
$settings = Resolve-DocxMarkupReferenceSettings $manifest $DocxMarkup $DocxMarkupGeometry -RequireMode

$caseId = [string]$manifest.id
Import-ReferenceCacheEntry `
    -InputPath $inputFull `
    -ReferencePdf $ReferencePdf `
    -ReferenceDirectory $ReferenceDirectory `
    -DpiValue $Dpi `
    -CacheVariant $settings.Variant `
    -CaseId $caseId `
    -Producer @{ Producer = "import"; OfficeApp = $OfficeApp; OfficeVersion = $OfficeVersion; ExportSettings = $ExportSettings } `
    -Force:$Force
