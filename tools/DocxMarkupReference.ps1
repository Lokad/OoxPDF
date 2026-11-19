# Canonical view settings bind a DOCX candidate and its trusted reference to the
# same cache variant. The settings identify an export; they do not drive Office.
function Resolve-DocxMarkupReferenceSettings($Manifest, [string] $Mode, [string] $Geometry, [switch] $RequireMode) {
    if ([string]::IsNullOrWhiteSpace($Mode)) { $Mode = [string]$Manifest.docxMarkup }
    if ([string]::IsNullOrWhiteSpace($Geometry)) { $Geometry = [string]$Manifest.docxMarkupGeometry }
    if ($RequireMode -and [string]::IsNullOrWhiteSpace($Mode)) {
        throw "DOCX markup reference requires docxMarkup in the case or -DocxMarkup."
    }

    $modeName = switch ($Mode.Trim().ToLowerInvariant()) {
        "" { "final" }
        "simple-markup" { "simple" }
        "all-markup" { "all" }
        { $_ -in @("final", "original", "simple", "all") } { $_ }
        default { throw "Unknown DOCX markup view: $Mode" }
    }
    $geometryName = switch ($Geometry.Trim().ToLowerInvariant()) {
        { $_ -in @("", "preserve", "preserve-layout", "preserve-document-layout") } { "preserve" }
        { $_ -in @("reserve", "reserve-margin", "markup-margin", "reserve-markup-margin") } { "reserve-margin" }
        { $_ -in @("word", "word-compatible", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup") } { "word-compatible" }
        default { throw "Unknown DOCX markup geometry: $Geometry" }
    }
    [pscustomobject]@{
        Mode = $modeName
        Geometry = $geometryName
        Variant = "docxMarkup=$modeName;docxMarkupGeometry=$geometryName"
    }
}

function Get-DocxMarkupReferenceCases([string[]] $Cases, [string] $PrivatePath) {
    $defaultSelection = $Cases.Count -eq 0
    $paths = if ($defaultSelection) {
        Get-ChildItem -LiteralPath (Join-Path $repoRoot "visual-cases/cases") -Directory | ForEach-Object {
            $path = Join-Path $_.FullName "case.json"
            if (Test-Path -LiteralPath $path -PathType Leaf) { $path }
        }
    }
    else {
        $Cases
    }
    $publicCases = foreach ($path in $paths) {
        $full = if ($defaultSelection) { $path } else { (Resolve-Path -LiteralPath $path).Path }
        $manifest = Get-Content -Raw -LiteralPath $full | ConvertFrom-Json
        $hasMarkup = $manifest.kind -eq "docx" -and -not [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)
        if ($defaultSelection -and (!$hasMarkup -or $manifest.tags -notcontains "docx-markup")) { continue }
        if (!$hasMarkup) { throw "Expected a DOCX case with docxMarkup: $full" }
        [pscustomobject]@{
            Id = [string]$manifest.id
            Path = $full
            Manifest = $manifest
            Settings = Resolve-DocxMarkupReferenceSettings $manifest -RequireMode
            PrivateSafe = $false
        }
    }
    if ($defaultSelection) { $publicCases | Sort-Object Id }
    else { $publicCases }

    if ($PrivatePath) {
        if (!(Test-Path -LiteralPath $PrivatePath -PathType Leaf)) { throw "Private DOCX markup case was requested but does not exist." }
        $full = (Resolve-Path -LiteralPath $PrivatePath).Path
        $manifest = Get-Content -Raw -LiteralPath $full | ConvertFrom-Json
        if ($manifest.kind -ne "docx" -or [string]::IsNullOrWhiteSpace([string]$manifest.input)) {
            throw "Private DOCX markup status requires a DOCX manifest with an input path."
        }
        $views = if ([string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
            Resolve-DocxMarkupReferenceSettings $manifest "final" "preserve"
            Resolve-DocxMarkupReferenceSettings $manifest "original" "preserve"
            Resolve-DocxMarkupReferenceSettings $manifest "simple" "preserve"
            Resolve-DocxMarkupReferenceSettings $manifest "all" "word-compatible"
        }
        else {
            Resolve-DocxMarkupReferenceSettings $manifest -RequireMode
        }
        foreach ($settings in $views) {
            [pscustomobject]@{
                Id = "private-docx-markup-$($settings.Mode)"
                Path = $full
                Manifest = $manifest
                Settings = $settings
                PrivateSafe = $true
            }
        }
    }
}

function Get-DocxMarkupReferenceStatus($CaseInfo, [int] $DpiValue) {
    $inputFull = Get-CaseInputPath $CaseInfo
    $inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $settings = $CaseInfo.Settings
    $cacheKey = Get-ReferenceIdentityKey $inputFull $settings.Variant $inputHash
    $cacheDir = Join-Path (Get-ReferenceCacheRoot) $cacheKey
    $exists = Test-Path -LiteralPath $cacheDir -PathType Container
    $entry = if ($exists) { Test-ReferenceIdentityEntry $cacheDir $inputHash $settings.Variant }
    $caseLabel = if ($CaseInfo.PrivateSafe) { $CaseInfo.Id } else { ConvertTo-RepoPath $CaseInfo.Path }
    $caseArgument = if ($CaseInfo.PrivateSafe) { "<private-case-manifest>" } else { "'" + $caseLabel.Replace("'", "''") + "'" }
    $overrides = if ($CaseInfo.PrivateSafe) { " -DocxMarkup $($settings.Mode) -DocxMarkupGeometry $($settings.Geometry)" } else { "" }

    [pscustomobject]@{
        Id = $CaseInfo.Id
        Case = $caseLabel
        PrivateSafe = [bool]$CaseInfo.PrivateSafe
        InputExtension = [IO.Path]::GetExtension($inputFull).TrimStart(".").ToLowerInvariant()
        InputSha256 = $inputHash
        Dpi = $DpiValue
        DocxMarkup = $settings.Mode
        DocxMarkupGeometry = $settings.Geometry
        CacheVariant = $settings.Variant
        CacheKey = $cacheKey
        ExpectedCacheDirectory = "artifacts/reference-cache/$cacheKey"
        CacheDirectoryExists = $exists
        CompleteMarkerPresent = Test-Path -LiteralPath (Join-Path $cacheDir "complete.txt") -PathType Leaf
        ReferencePdfPresent = Test-Path -LiteralPath (Join-Path $cacheDir "reference.pdf") -PathType Leaf
        ReferenceMetadataPresent = Test-Path -LiteralPath (Join-Path $cacheDir "reference-metadata.json") -PathType Leaf
        ReferencePdfSha256 = $entry.Metadata.ReferencePdfSha256
        MetadataRasterPageCount = $entry.Metadata.RasterPageCount
        CacheState = if ($exists) { $entry.State } else { "Missing" }
        Ready = $entry.State -eq "Complete"
        ImportCommand = "pwsh tools/ImportDocxMarkupReferenceCache.ps1 -Case $caseArgument -ReferencePdf <trusted-reference-pdf>$overrides -Dpi $DpiValue -Force"
    }
}
