param(
    [string[]] $Case = @(),

    [string] $PrivateCase = "private-cases/docx-markup-track-changes.json",

    [string] $OutputRoot,

    [int] $Dpi = 144,

    [switch] $IncludePrivate,

    [switch] $ValidateOnly,

    [switch] $CacheStatusOnly,

    [switch] $FailOnMissingCache,

    [switch] $SkipRasterDiff,

    [switch] $FailOnDeltas,

    [switch] $ContinueOnFailure
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

function Read-Manifest([string] $Path) {
    $full = (Resolve-Path -LiteralPath $Path).Path
    $manifest = Get-Content -Raw -LiteralPath $full | ConvertFrom-Json
    if ($manifest.kind -ne "docx") {
        throw "DOCX markup reference gate only supports DOCX cases: $full"
    }

    if ($manifest.PSObject.Properties.Name -notcontains "docxMarkup" -or [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
        throw "DOCX markup reference gate requires docxMarkup in case: $full"
    }

    [pscustomobject]@{
        Id = [string]$manifest.id
        Path = $full
        Manifest = $manifest
    }
}

function New-PrivateMarkupCaseInfo($Manifest, [string] $Path, [string] $Id, [string] $MarkupMode, [string] $MarkupGeometry) {
    $syntheticManifest = [pscustomobject]@{
        id = $Id
        kind = "docx"
        input = [string]$Manifest.input
        docxMarkup = $MarkupMode
        docxMarkupGeometry = $MarkupGeometry
    }

    [pscustomobject]@{
        Id = $Id
        Path = $Path
        Manifest = $syntheticManifest
        DocxMarkupOverride = $MarkupMode
        DocxMarkupGeometryOverride = $MarkupGeometry
        PrivateSafe = $true
    }
}

function Get-PrivateMarkupCaseId([string] $MarkupMode) {
    $suffix = if ([string]::IsNullOrWhiteSpace($MarkupMode)) {
        "mode"
    }
    else {
        $MarkupMode.ToLowerInvariant() -replace "[^a-z0-9]+", "-"
    }

    $suffix = $suffix.Trim("-")
    if ([string]::IsNullOrWhiteSpace($suffix)) {
        $suffix = "mode"
    }

    "private-docx-markup-{0}" -f $suffix
}

function Read-PrivateCacheStatusManifests([string] $Path) {
    $full = (Resolve-Path -LiteralPath $Path).Path
    $manifest = Get-Content -Raw -LiteralPath $full | ConvertFrom-Json
    if ($manifest.kind -ne "docx" -or [string]::IsNullOrWhiteSpace([string]$manifest.input)) {
        throw "Private DOCX markup cache status requires a DOCX private manifest with an input path."
    }

    if ($manifest.PSObject.Properties.Name -contains "docxMarkup" -and -not [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
        $markupMode = [string]$manifest.docxMarkup
        $markupGeometry = if ($manifest.PSObject.Properties.Name -contains "docxMarkupGeometry" -and -not [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkupGeometry)) {
            [string]$manifest.docxMarkupGeometry
        }
        else {
            "preserve"
        }

        return ,(New-PrivateMarkupCaseInfo $manifest $full (Get-PrivateMarkupCaseId $markupMode) $markupMode $markupGeometry)
    }

    @(
        New-PrivateMarkupCaseInfo $manifest $full "private-docx-markup-final" "final" "preserve"
        New-PrivateMarkupCaseInfo $manifest $full "private-docx-markup-original" "original" "preserve"
        New-PrivateMarkupCaseInfo $manifest $full "private-docx-markup-simple" "simple" "preserve"
        New-PrivateMarkupCaseInfo $manifest $full "private-docx-markup-all" "all" "word-compatible"
    )
}

function Read-JsonObjectIfExists([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    return Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
}

function Get-JsonPropertyValue($Object, [string] $Name) {
    if ($null -eq $Object -or $Object.PSObject.Properties.Name -notcontains $Name) {
        return $null
    }

    return $Object.PSObject.Properties[$Name].Value
}

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

function ConvertTo-RepoPath([string] $Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    if ($fullPath.StartsWith($fullRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullPath.Substring($fullRoot.Length + 1).Replace([System.IO.Path]::DirectorySeparatorChar, "/")
    }

    return $fullPath
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

function Get-CaseInputPath($CaseInfo) {
    $caseDirectory = Split-Path -Parent $CaseInfo.Path
    return (Resolve-Path -LiteralPath (Join-Path $caseDirectory $CaseInfo.Manifest.input)).Path
}

function Get-CaseCacheVariant($CaseInfo) {
    $docxMarkup = [string]$CaseInfo.Manifest.docxMarkup
    $docxMarkupGeometry = if ($CaseInfo.Manifest.PSObject.Properties.Name -contains "docxMarkupGeometry" -and -not [string]::IsNullOrWhiteSpace([string]$CaseInfo.Manifest.docxMarkupGeometry)) {
        [string]$CaseInfo.Manifest.docxMarkupGeometry
    }
    else {
        "preserve"
    }

    "docxMarkup={0};docxMarkupGeometry={1}" -f $docxMarkup, $docxMarkupGeometry
}

function New-ReferenceImportCommand($CaseInfo, [int] $DpiValue) {
    $docxMarkup = [string]$CaseInfo.Manifest.docxMarkup
    $docxMarkupGeometry = if ($CaseInfo.Manifest.PSObject.Properties.Name -contains "docxMarkupGeometry" -and -not [string]::IsNullOrWhiteSpace([string]$CaseInfo.Manifest.docxMarkupGeometry)) {
        [string]$CaseInfo.Manifest.docxMarkupGeometry
    }
    else {
        "preserve"
    }
    $isPrivateSafe = $CaseInfo.PSObject.Properties.Name -contains "PrivateSafe" -and $CaseInfo.PrivateSafe -eq $true
    $caseArgument = if ($isPrivateSafe) { "<private-case-manifest>" } else { ConvertTo-RepoPath $CaseInfo.Path }
    $overrideArguments = if ($isPrivateSafe) {
        " -DocxMarkup $docxMarkup -DocxMarkupGeometry $docxMarkupGeometry"
    }
    else {
        ""
    }

    "pwsh tools/ImportDocxMarkupReferenceCache.ps1 -Case $caseArgument -ReferencePdf <trusted-reference-pdf>$overrideArguments -Dpi $DpiValue -Force"
}

function Get-CacheMetadata([string] $CacheDirectory) {
    $metadataPath = Join-Path $CacheDirectory "reference-metadata.json"
    if (-not (Test-Path -LiteralPath $metadataPath)) {
        return $null
    }

    Get-Content -Raw -LiteralPath $metadataPath | ConvertFrom-Json
}

function Get-CaseReferenceCacheStatus($CaseInfo, [int] $DpiValue) {
    $inputFull = Get-CaseInputPath $CaseInfo
    $inputHash = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $cacheVariant = Get-CaseCacheVariant $CaseInfo
    $cacheKey = Get-ReferenceCacheKey $inputFull $DpiValue $cacheVariant
    $cacheDir = Join-Path (Join-Path $repoRoot "artifacts/reference-cache") $cacheKey
    $completeMarker = Join-Path $cacheDir "complete.txt"
    $metadata = Get-CacheMetadata $cacheDir
    $referencePdfPath = Join-Path $cacheDir "reference.pdf"
    $rasterPageCount = if (Test-Path -LiteralPath $cacheDir) {
        @(Get-ChildItem -LiteralPath $cacheDir -Filter "page-*.png" -ErrorAction SilentlyContinue).Count
    }
    else {
        0
    }
    $ready = (Test-Path -LiteralPath $completeMarker) -and (Test-Path -LiteralPath $referencePdfPath)
    $isPrivateSafe = $CaseInfo.PSObject.Properties.Name -contains "PrivateSafe" -and $CaseInfo.PrivateSafe -eq $true

    [pscustomobject]@{
        Id = $CaseInfo.Id
        Case = if ($isPrivateSafe) { $CaseInfo.Id } else { ConvertTo-RepoPath $CaseInfo.Path }
        PrivateSafe = [bool]$isPrivateSafe
        InputExtension = [System.IO.Path]::GetExtension($inputFull).TrimStart(".").ToLowerInvariant()
        InputSha256 = $inputHash
        Dpi = $DpiValue
        CacheVariant = $cacheVariant
        CacheKey = $cacheKey
        ExpectedCacheDirectory = "artifacts/reference-cache/$cacheKey"
        CacheDirectory = $cacheDir
        CacheDirectoryExists = Test-Path -LiteralPath $cacheDir
        CompleteMarkerPresent = Test-Path -LiteralPath $completeMarker
        ReferencePdfPresent = Test-Path -LiteralPath $referencePdfPath
        ReferenceMetadataPresent = $null -ne $metadata
        ReferencePdfSha256 = Get-JsonPropertyValue $metadata "ReferencePdfSha256"
        RasterPageCount = $rasterPageCount
        MetadataRasterPageCount = Get-JsonPropertyValue $metadata "RasterPageCount"
        Ready = $ready
        MissingImportCommand = if ($ready) { $null } else { New-ReferenceImportCommand $CaseInfo $DpiValue }
    }
}

function Get-CaseComparisonSummary([string] $OutputDirectory) {
    $summaryPath = Join-Path $OutputDirectory "summary.json"
    $gateSummaryPath = Join-Path $OutputDirectory "comparison/gate-summary.json"
    $summary = Read-JsonObjectIfExists $summaryPath
    $gateSummary = Read-JsonObjectIfExists $gateSummaryPath
    $blockingPhase = Get-JsonPropertyValue $gateSummary "CurrentBlockingPhase"
    if ($null -eq $blockingPhase) {
        $blockingPhase = Get-JsonPropertyValue $summary "CurrentBlockingGatePhase"
    }
    $gateFailureCount = Get-JsonPropertyValue $gateSummary "FailureCount"
    if ($null -eq $gateFailureCount) {
        $gateFailureCount = Get-JsonPropertyValue $summary "GateFailureCount"
    }

    [pscustomobject]@{
        SummaryPath = if (Test-Path -LiteralPath $summaryPath) { $summaryPath } else { $null }
        GateSummaryPath = if (Test-Path -LiteralPath $gateSummaryPath) { $gateSummaryPath } else { $null }
        ReferencePageCount = Get-JsonPropertyValue $summary "ReferencePageCount"
        CandidatePageCount = Get-JsonPropertyValue $summary "CandidatePageCount"
        PageCountDelta = Get-JsonPropertyValue $summary "PageCountDelta"
        VisualDeltaCount = Get-JsonPropertyValue $summary "VisualDeltaCount"
        RasterRegionDeltaCount = Get-JsonPropertyValue $summary "RasterRegionDeltaCount"
        AnnotationDeltaCount = Get-JsonPropertyValue $summary "AnnotationDeltaCount"
        AnnotationTargetDeltaCount = Get-JsonPropertyValue $summary "AnnotationTargetDeltaCount"
        BalloonGeometryDeltaCount = Get-JsonPropertyValue $summary "BalloonGeometryDeltaCount"
        GateFailureCount = $gateFailureCount
        BlockingGatePhaseId = Get-JsonPropertyValue $blockingPhase "Id"
        BlockingGatePhaseName = Get-JsonPropertyValue $blockingPhase "Name"
    }
}

function Get-DefaultDocxMarkupCases {
    $caseRoot = Join-Path $repoRoot "visual-cases/cases"
    Get-ChildItem -LiteralPath $caseRoot -Directory |
        ForEach-Object {
            $manifestPath = Join-Path $_.FullName "case.json"
            if (-not (Test-Path -LiteralPath $manifestPath)) {
                return
            }

            $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
            $tags = @($manifest.tags)
            if ($manifest.kind -eq "docx" -and $tags -contains "docx-markup" -and
                $manifest.PSObject.Properties.Name -contains "docxMarkup" -and
                -not [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
                Read-Manifest $manifestPath
            }
        } |
        Sort-Object Id
}

$cases = if ($Case.Count -ne 0) {
    @($Case | ForEach-Object { Read-Manifest $_ })
}
else {
    @(Get-DefaultDocxMarkupCases)
}

if ($IncludePrivate) {
    $privatePath = Join-Path $repoRoot $PrivateCase
    if (-not (Test-Path -LiteralPath $privatePath)) {
        throw "Private DOCX markup case was requested but does not exist: $privatePath"
    }

    $cases = @($cases) + @(Read-PrivateCacheStatusManifests $privatePath)
}

if ($cases.Count -eq 0) {
    throw "No DOCX markup cases were selected."
}

if ($ValidateOnly) {
    Write-Host ("DOCX markup reference gate validation passed for {0} case(s)." -f $cases.Count)
    $cases | ForEach-Object { Write-Host ("- {0}: {1}" -f $_.Id, $_.Path) }
    return
}

if ($CacheStatusOnly) {
    $outputRootFull = if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
        Join-Path $repoRoot ("artifacts/docx-markup-reference-cache-status/" + (Get-Date -Format "yyyyMMdd-HHmmss"))
    }
    else {
        [System.IO.Path]::GetFullPath($OutputRoot)
    }
    New-Item -ItemType Directory -Force -Path $outputRootFull | Out-Null

    $results = @($cases | ForEach-Object { Get-CaseReferenceCacheStatus $_ $Dpi })
    $commandsPath = Join-Path $outputRootFull "missing-import-commands.ps1"
    @(
        "# Import trusted DOCX markup references after controlled external Office export."
        "# Replace <trusted-reference-pdf> and, for private entries, <private-case-manifest> before running."
        $results | Where-Object { $_.Ready -ne $true } | ForEach-Object { $_.MissingImportCommand }
    ) | Set-Content -LiteralPath $commandsPath -Encoding UTF8
    $summary = [pscustomobject]@{
        CaseCount = $results.Count
        ReadyCount = @($results | Where-Object { $_.Ready -eq $true }).Count
        MissingCount = @($results | Where-Object { $_.Ready -ne $true }).Count
        Dpi = $Dpi
        MissingImportCommands = ConvertTo-RepoPath $commandsPath
        Cases = $results
    }
    $summaryPath = Join-Path $outputRootFull "summary.json"
    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $summaryPath -Encoding UTF8

    Write-Host ("DOCX markup reference cache status: {0} ready, {1} missing. See {2}" -f $summary.ReadyCount, $summary.MissingCount, $summaryPath)
    if ($FailOnMissingCache -and $summary.MissingCount -ne 0) {
        throw "DOCX markup reference cache status found missing cache entries. See $summaryPath"
    }

    return
}

$outputRootFull = if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    Join-Path $repoRoot ("artifacts/docx-markup-reference-gate/" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}
else {
    [System.IO.Path]::GetFullPath($OutputRoot)
}
New-Item -ItemType Directory -Force -Path $outputRootFull | Out-Null

$results = New-Object System.Collections.Generic.List[object]
foreach ($caseInfo in $cases) {
    $caseOutput = Join-Path $outputRootFull $caseInfo.Id
    $arguments = @{
        Case = $caseInfo.Path
        OutputDirectory = $caseOutput
        Dpi = $Dpi
    }
    if ($caseInfo.PSObject.Properties.Name -contains "DocxMarkupOverride") {
        $arguments.DocxMarkup = [string]$caseInfo.DocxMarkupOverride
        $arguments.DocxMarkupGeometry = [string]$caseInfo.DocxMarkupGeometryOverride
        $arguments.CaseId = [string]$caseInfo.Id
        $arguments.PrivateSafeSummary = $true
    }
    if ($FailOnDeltas) {
        $arguments.FailOnDeltas = $true
    }
    if ($SkipRasterDiff) {
        $arguments.SkipRasterDiff = $true
    }

    $exitCode = 0
    $errorMessage = $null
    try {
        & (Join-Path $PSScriptRoot "CompareCachedDocxMarkupReference.ps1") @arguments
        $exitCode = $LASTEXITCODE
    }
    catch {
        $exitCode = 1
        $errorMessage = if ($caseInfo.PSObject.Properties.Name -contains "PrivateSafe" -and $caseInfo.PrivateSafe -eq $true) {
            "Private DOCX markup reference comparison failed. See private-safe cache status for expected hashes and cache keys."
        }
        else {
            $_.Exception.Message
        }
    }

    $comparisonSummary = Get-CaseComparisonSummary $caseOutput
    $results.Add([pscustomobject]@{
        Id = $caseInfo.Id
        Case = if ($caseInfo.PSObject.Properties.Name -contains "PrivateSafe" -and $caseInfo.PrivateSafe -eq $true) { $caseInfo.Id } else { $caseInfo.Path }
        OutputDirectory = $caseOutput
        ExitCode = $exitCode
        ErrorMessage = $errorMessage
        SummaryPath = $comparisonSummary.SummaryPath
        GateSummaryPath = $comparisonSummary.GateSummaryPath
        ReferencePageCount = $comparisonSummary.ReferencePageCount
        CandidatePageCount = $comparisonSummary.CandidatePageCount
        PageCountDelta = $comparisonSummary.PageCountDelta
        VisualDeltaCount = $comparisonSummary.VisualDeltaCount
        RasterRegionDeltaCount = $comparisonSummary.RasterRegionDeltaCount
        AnnotationDeltaCount = $comparisonSummary.AnnotationDeltaCount
        AnnotationTargetDeltaCount = $comparisonSummary.AnnotationTargetDeltaCount
        BalloonGeometryDeltaCount = $comparisonSummary.BalloonGeometryDeltaCount
        GateFailureCount = $comparisonSummary.GateFailureCount
        BlockingGatePhaseId = $comparisonSummary.BlockingGatePhaseId
        BlockingGatePhaseName = $comparisonSummary.BlockingGatePhaseName
    }) | Out-Null
    if ($exitCode -ne 0 -and -not $ContinueOnFailure) {
        break
    }
}

$summaryPath = Join-Path $outputRootFull "summary.json"
$results | ConvertTo-Json -Depth 6 -AsArray | Set-Content -LiteralPath $summaryPath -Encoding UTF8

$failed = @($results | Where-Object { $_.ExitCode -ne 0 })
if ($failed.Count -ne 0) {
    throw "DOCX markup reference gate failed. See $summaryPath"
}

Write-Host "DOCX markup reference gate artifacts: $outputRootFull"
