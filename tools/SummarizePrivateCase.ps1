param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [string] $Run
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$privateRoot = Join-Path $repoRoot "private-cases"
$artifactRoot = Join-Path $repoRoot "artifacts/private-visual"

function Test-UnderDirectory([string] $Path, [string] $Directory) {
    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $fullDirectory = [System.IO.Path]::GetFullPath($Directory).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return $fullPath.Equals($fullDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

$caseFull = (Resolve-Path -LiteralPath $Case).Path
if (-not (Test-UnderDirectory $caseFull $privateRoot)) {
    throw "Private case manifest must be under $privateRoot."
}

$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace($manifest.id)) {
    throw "Private case manifest must contain an id."
}

$caseId = [string]$manifest.id
if ($caseId.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $caseId.Contains("/") -or $caseId.Contains("\")) {
    throw "Private case id must be a single filename-safe path segment."
}

$caseArtifactRoot = Join-Path $artifactRoot $caseId
if ([string]::IsNullOrWhiteSpace($Run)) {
    $runDirectory = Get-ChildItem -LiteralPath $caseArtifactRoot -Directory |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1
    if ($runDirectory -eq $null) {
        throw "No private visual runs found under $caseArtifactRoot."
    }
}
else {
    if ($Run.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $Run.Contains("/") -or $Run.Contains("\")) {
        throw "Run must be a single filename-safe path segment."
    }

    $runDirectory = Get-Item -LiteralPath (Join-Path $caseArtifactRoot $Run)
}

$runRoot = $runDirectory.FullName
if (-not (Test-UnderDirectory $runRoot $artifactRoot)) {
    throw "Run directory must be under $artifactRoot."
}

$metricsPath = Join-Path $runRoot "comparison/metrics.json"
$diagnosticsPath = Join-Path $runRoot "candidate/diagnostics.json"
$referencePageCount = @(Get-ChildItem -LiteralPath (Join-Path $runRoot "reference") -Filter "page-*.png" -ErrorAction SilentlyContinue).Count
$candidatePageCount = @(Get-ChildItem -LiteralPath (Join-Path $runRoot "candidate") -Filter "page-*.png" -ErrorAction SilentlyContinue).Count
$metrics = if (Test-Path -LiteralPath $metricsPath) { @(Get-Content -Raw -LiteralPath $metricsPath | ConvertFrom-Json) } else { @() }
$diagnostics = if (Test-Path -LiteralPath $diagnosticsPath) { @(Get-Content -Raw -LiteralPath $diagnosticsPath | ConvertFrom-Json) } else { @() }

$dimensionMismatches = @($metrics | Where-Object { -not $_.DimensionsMatch })
$worstPages = @($metrics |
    Where-Object { $_.MeanAbsoluteError -ne $null } |
    Sort-Object MeanAbsoluteError -Descending |
    Select-Object -First 5 Page, MeanAbsoluteError, ChangedPixelRatioAtThreshold16, DimensionsMatch)
$diagnosticGroups = @($diagnostics |
    Group-Object Id, Severity |
    Sort-Object Count -Descending |
    ForEach-Object {
        [pscustomobject]@{
            Id = $_.Group[0].Id
            Severity = $_.Group[0].Severity
            Count = $_.Count
        }
    })

[pscustomobject]@{
    CaseId = $caseId
    Run = Split-Path -Leaf $runRoot
    ReferencePages = $referencePageCount
    CandidatePages = $candidatePageCount
    ComparedPages = $metrics.Count
    DimensionMismatchCount = $dimensionMismatches.Count
    MeanAbsoluteError = if ($metrics.Count -gt 0) { [math]::Round(($metrics | Measure-Object MeanAbsoluteError -Average).Average, 6) } else { $null }
    MaxMeanAbsoluteError = if ($metrics.Count -gt 0) { [math]::Round(($metrics | Measure-Object MeanAbsoluteError -Maximum).Maximum, 6) } else { $null }
    MeanChangedPixelRatioAtThreshold16 = if ($metrics.Count -gt 0) { [math]::Round(($metrics | Measure-Object ChangedPixelRatioAtThreshold16 -Average).Average, 6) } else { $null }
    Diagnostics = $diagnosticGroups
    WorstPages = $worstPages
} | ConvertTo-Json -Depth 5
