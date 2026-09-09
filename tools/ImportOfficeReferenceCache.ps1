# Trusted Office-reference import for ordinary DOCX/PPTX inputs (R02).
#
# The caller vouches that the reference PDF was exported by Office from the
# given input. The import verifies PDF structure, binds raster derivatives to
# the PDF hash, publishes atomically, and re-verifies. Use this for normal
# visual cases; markup-variant cases should prefer
# ImportDocxMarkupReferenceCache.ps1, which derives the cache variant from the
# case manifest.

param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [string] $ReferencePdf,

    [string] $ReferenceDirectory,

    [string] $CacheVariant,

    [string] $CaseId,

    [string] $OfficeApp,

    [string] $OfficeVersion,

    [string] $ExportSettings,

    [int] $Dpi = 144,

    [switch] $Force
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "ReferenceCache.ps1")

if ([string]::IsNullOrWhiteSpace($ReferencePdf) -and [string]::IsNullOrWhiteSpace($ReferenceDirectory)) {
    throw "Provide either -ReferencePdf or -ReferenceDirectory."
}

if (-not [string]::IsNullOrWhiteSpace($ReferencePdf) -and -not [string]::IsNullOrWhiteSpace($ReferenceDirectory)) {
    throw "Use either -ReferencePdf or -ReferenceDirectory, not both."
}

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
$extension = [System.IO.Path]::GetExtension($inputFull).ToLowerInvariant()
if ($extension -ne ".docx" -and $extension -ne ".pptx") {
    throw "Unsupported import input extension '$extension'. Expected .docx or .pptx."
}

Import-ReferenceCacheEntry `
    -InputPath $inputFull `
    -ReferencePdf $ReferencePdf `
    -ReferenceDirectory $ReferenceDirectory `
    -DpiValue $Dpi `
    -CacheVariant $CacheVariant `
    -CaseId $CaseId `
    -Producer @{ Producer = "import"; OfficeApp = $OfficeApp; OfficeVersion = $OfficeVersion; ExportSettings = $ExportSettings } `
    -Force:$Force
