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

    [switch] $ContinueOnFailure,
    [string] $CliPath,
    [string] $DocxInspectorPath,
    [string] $PdfInspectorPath,
    [switch] $SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

. (Join-Path $PSScriptRoot "JsonObject.ps1")

function Get-JsonPropertyValue($Object, [string] $Name) {
    if ($null -eq $Object -or $Object.PSObject.Properties.Name -notcontains $Name) {
        return $null
    }

    return $Object.PSObject.Properties[$Name].Value
}

. (Join-Path $PSScriptRoot "ReferenceCache.ps1")
. (Join-Path $PSScriptRoot "PathHelpers.ps1")
. (Join-Path $PSScriptRoot "DocxMarkupReference.ps1")

function Get-CaseReferenceCacheStatus($CaseInfo, [int] $DpiValue, [string] $DerivativeName) {
    $row = Get-DocxMarkupReferenceStatus $CaseInfo $DpiValue
    $cacheDir = Join-Path (Get-ReferenceCacheRoot) $row.CacheKey
    $derivativeDir = Join-Path (Join-Path $cacheDir "raster") $DerivativeName
    $rasterPageCount = if (Test-Path -LiteralPath $derivativeDir -PathType Container) {
        @(Get-ChildItem -LiteralPath $derivativeDir -Filter "page-*.png" -ErrorAction SilentlyContinue).Count
    }
    elseif ($row.CacheDirectoryExists) {
        @(Get-ChildItem -LiteralPath $cacheDir -Filter "page-*.png" -ErrorAction SilentlyContinue).Count
    }
    else {
        0
    }

    $row | Add-Member -NotePropertyName CacheDirectory -NotePropertyValue $cacheDir
    $row | Add-Member -NotePropertyName RasterPageCount -NotePropertyValue $rasterPageCount
    $row | Add-Member -NotePropertyName MissingImportCommand -NotePropertyValue $(if (!$row.Ready) { $row.ImportCommand })
    $row
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
        MarkupTextWitnessStatus = Get-JsonPropertyValue $summary 'MarkupTextWitnessStatus'
        MissingMarkupTextWitnessCount = Get-JsonPropertyValue $summary 'MissingMarkupTextWitnessCount'
        GateFailureCount = $gateFailureCount
        BlockingGatePhaseId = Get-JsonPropertyValue $blockingPhase "Id"
        BlockingGatePhaseName = Get-JsonPropertyValue $blockingPhase "Name"
    }
}

$privatePath = if ($IncludePrivate) { Join-Path $repoRoot $PrivateCase }
$cases = @(Get-DocxMarkupReferenceCases $Case $privatePath)

if ($cases.Count -eq 0) {
    throw "No DOCX markup cases were selected."
}

if ($ValidateOnly) {
    Write-Host ("DOCX markup reference gate validation passed for {0} case(s)." -f $cases.Count)
    foreach ($caseInfo in $cases) {
        if ($caseInfo.PrivateSafe) { Write-Host ("- {0}" -f $caseInfo.Id) }
        else { Write-Host ("- {0}: {1}" -f $caseInfo.Id, $caseInfo.Path) }
    }
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

    $derivativeName = Get-ReferenceDerivativeName $Dpi
    $results = @($cases | ForEach-Object { Get-CaseReferenceCacheStatus $_ $Dpi $derivativeName })
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
        CliPath = $CliPath
        DocxInspectorPath = $DocxInspectorPath
        PdfInspectorPath = $PdfInspectorPath
        SkipBuild = [bool]$SkipBuild
    }
    if ($caseInfo.PrivateSafe) {
        $arguments.DocxMarkup = $caseInfo.Settings.Mode
        $arguments.DocxMarkupGeometry = $caseInfo.Settings.Geometry
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
        $errorMessage = if ($caseInfo.PrivateSafe) {
            "Private DOCX markup reference comparison failed. See private-safe cache status for expected hashes and cache keys."
        }
        else {
            $_.Exception.Message
        }
    }

    $comparisonSummary = Get-CaseComparisonSummary $caseOutput
    $results.Add([pscustomobject]@{
        Id = $caseInfo.Id
        Case = if ($caseInfo.PrivateSafe) { $caseInfo.Id } else { $caseInfo.Path }
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
        MarkupTextWitnessStatus = $comparisonSummary.MarkupTextWitnessStatus
        MissingMarkupTextWitnessCount = $comparisonSummary.MissingMarkupTextWitnessCount
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
