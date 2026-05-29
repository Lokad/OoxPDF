param(
    [Parameter(Mandatory = $true)]
    [string] $Family,

    [switch] $List,

    [switch] $OnlyUnsupported,

    [switch] $OnlyChanged,

    [switch] $UpdateCatalog,

    [string[]] $CasePattern = @(),

    [string[]] $Tag = @(),

    [int] $Limit = 0
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$familyPath = Join-Path $repoRoot ("visual-cases/families/{0}.json" -f $Family)
if (-not (Test-Path -LiteralPath $familyPath)) {
    throw "Unknown visual family '$Family'. Expected $familyPath."
}

$familyManifest = Get-Content -Raw -LiteralPath $familyPath | ConvertFrom-Json
$caseRoot = Join-Path $repoRoot "visual-cases/cases"
$artifactRoot = Join-Path $repoRoot "artifacts/visual"
$reportRoot = Join-Path $artifactRoot "reports"
$supportCatalogPath = Join-Path $artifactRoot "support-catalog.json"

function Read-SupportCatalog {
    if (-not (Test-Path -LiteralPath $supportCatalogPath)) {
        return @{}
    }

    $catalog = Get-Content -Raw -LiteralPath $supportCatalogPath | ConvertFrom-Json
    $map = @{}
    foreach ($entry in @($catalog.cases)) {
        $map[$entry.id] = $entry
    }

    return $map
}

function Get-ChangedPaths {
    $paths = @()
    $status = & git -C $repoRoot status --short
    if ($LASTEXITCODE -ne 0) {
        return $paths
    }

    foreach ($line in $status) {
        if ($line.Length -lt 4) {
            continue
        }

        $path = $line.Substring(3).Trim()
        if ($path.Contains(" -> ")) {
            $path = ($path -split " -> ")[-1]
        }

        $paths += $path.Replace("\", "/")
    }

    return $paths
}

function Get-CaseInputPath([string] $caseManifestPath) {
    $caseManifest = Get-Content -Raw -LiteralPath $caseManifestPath | ConvertFrom-Json
    $caseDirectory = Split-Path -Parent $caseManifestPath
    $inputPath = Join-Path $caseDirectory $caseManifest.input
    try {
        return (Resolve-Path -LiteralPath $inputPath).Path
    }
    catch {
        return $inputPath
    }
}

function Convert-ToRepoPath([string] $path) {
    return (Resolve-Path -LiteralPath $path).Path.Substring($repoRoot.Length).TrimStart("\", "/").Replace("\", "/")
}

function Get-LatestCaseRun([string] $caseId) {
    $caseArtifactRoot = Join-Path $artifactRoot $caseId
    if (-not (Test-Path -LiteralPath $caseArtifactRoot)) {
        return $null
    }

    return Get-ChildItem -LiteralPath $caseArtifactRoot -Directory |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1
}

function Get-ManifestDouble($object, [string] $name, $defaultValue) {
    if ($object -ne $null -and $object.PSObject.Properties.Name -contains $name -and $object.$name -ne $null) {
        return [double]$object.$name
    }

    return $defaultValue
}

function Get-CaseClassification($caseManifest) {
    $tags = @($caseManifest.tags)
    if ($tags -contains "locked") {
        return "locked"
    }

    if ($tags -contains "locked-text-ops") {
        return "locked-text-ops"
    }

    if ($tags -contains "approximate") {
        return "approximate"
    }

    if ($tags -contains "needs-review") {
        return "needs-review"
    }

    return "unclassified"
}

function New-ReportRow($case, [bool] $passed, [string] $errorMessage) {
    $caseManifestPath = Join-Path $case.FullName "case.json"
    $caseManifest = Get-Content -Raw -LiteralPath $caseManifestPath | ConvertFrom-Json
    $latestRun = Get-LatestCaseRun $caseManifest.id
    $metrics = @()
    $diagnostics = @()
    if ($latestRun -ne $null) {
        $metricsPath = Join-Path $latestRun.FullName "comparison/metrics.json"
        if (Test-Path -LiteralPath $metricsPath) {
            $metrics = @(Get-Content -Raw -LiteralPath $metricsPath | ConvertFrom-Json)
        }

        $diagnosticsPath = Join-Path $latestRun.FullName "candidate/diagnostics.json"
        if (Test-Path -LiteralPath $diagnosticsPath) {
            $diagnosticsJson = Get-Content -Raw -LiteralPath $diagnosticsPath
            if ($diagnosticsJson.Trim() -ne "[]") {
                $diagnostics = @(ConvertFrom-Json -InputObject $diagnosticsJson)
            }
        }
    }

    $maeValues = @($metrics | Where-Object { $_.MeanAbsoluteError -ne $null } | ForEach-Object { [double]$_.MeanAbsoluteError })
    $changed16Values = @($metrics | Where-Object { $_.ChangedPixelRatioAtThreshold16 -ne $null } | ForEach-Object { [double]$_.ChangedPixelRatioAtThreshold16 })
    $ssimValues = @($metrics | Where-Object { $_.StructuralSimilarity -ne $null } | ForEach-Object { [double]$_.StructuralSimilarity })
    $histValues = @($metrics | Where-Object { $_.ForegroundColorHistogramCorrelation -ne $null } | ForEach-Object { [double]$_.ForegroundColorHistogramCorrelation })
    $dimensionMismatchCount = @($metrics | Where-Object { $_.DimensionsMatch -ne $true }).Count
    $minStructuralSimilarity = if ($ssimValues.Count -eq 0) { $null } else { ($ssimValues | Measure-Object -Minimum).Minimum }
    $minForegroundColorHistogramCorrelation = if ($histValues.Count -eq 0) { $null } else { ($histValues | Measure-Object -Minimum).Minimum }
    $support = $caseManifest.support
    $familySupport = $familyManifest.support
    $requiredStructuralSimilarity = Get-ManifestDouble $support "minStructuralSimilarity" (Get-ManifestDouble $familySupport "minStructuralSimilarity" 0.995)
    $requiredColorHistogram = Get-ManifestDouble $support "minForegroundColorHistogramCorrelation" (Get-ManifestDouble $familySupport "minForegroundColorHistogramCorrelation" 0.98)
    $status = if (-not $passed) {
        "unsupported"
    }
    elseif ($minStructuralSimilarity -eq $null -or
        $minForegroundColorHistogramCorrelation -eq $null -or
        $minStructuralSimilarity -lt $requiredStructuralSimilarity -or
        $minForegroundColorHistogramCorrelation -lt $requiredColorHistogram) {
        "needs-review"
    }
    else {
        "supported"
    }

    [pscustomobject]@{
        id = $caseManifest.id
        family = $Family
        kind = $caseManifest.kind
        classification = Get-CaseClassification $caseManifest
        passed = $passed
        status = $status
        runPath = if ($latestRun -ne $null) { Convert-ToRepoPath $latestRun.FullName } else { $null }
        pageCount = $metrics.Count
        dimensionMismatchCount = $dimensionMismatchCount
        diagnosticsCount = $diagnostics.Count
        diagnosticIds = @($diagnostics | ForEach-Object { $_.Id } | Sort-Object -Unique) -join ";"
        minStructuralSimilarity = $minStructuralSimilarity
        minForegroundColorHistogramCorrelation = $minForegroundColorHistogramCorrelation
        requiredStructuralSimilarity = $requiredStructuralSimilarity
        requiredForegroundColorHistogramCorrelation = $requiredColorHistogram
        maxMeanAbsoluteError = if ($maeValues.Count -eq 0) { $null } else { ($maeValues | Measure-Object -Maximum).Maximum }
        maxChangedPixelRatioAtThreshold16 = if ($changed16Values.Count -eq 0) { $null } else { ($changed16Values | Measure-Object -Maximum).Maximum }
        error = $errorMessage
    }
}

$patterns = @($familyManifest.casePatterns)
$cases = @(
    Get-ChildItem -LiteralPath $caseRoot -Directory |
        Where-Object {
            $caseName = $_.Name
            $patterns | Where-Object { $caseName -like $_ }
        } |
        Sort-Object Name
)

$supportCatalog = Read-SupportCatalog
if ($CasePattern.Count -ne 0) {
    $cases = @($cases | Where-Object {
        $caseName = $_.Name
        $CasePattern | Where-Object { $caseName -like $_ }
    })
}

if ($Tag.Count -ne 0) {
    $cases = @($cases | Where-Object {
        $caseManifestPath = Join-Path $_.FullName "case.json"
        $caseManifest = Get-Content -Raw -LiteralPath $caseManifestPath | ConvertFrom-Json
        $caseTags = @($caseManifest.tags)
        $missingTags = @($Tag | Where-Object { $caseTags -notcontains $_ })
        $missingTags.Count -eq 0
    })
}

if ($OnlyUnsupported) {
    $cases = @($cases | Where-Object {
        $id = $_.Name
        -not $supportCatalog.ContainsKey($id) -or $supportCatalog[$id].status -ne "supported"
    })
}

if ($OnlyChanged) {
    $changedPaths = @(Get-ChangedPaths)
    $cases = @($cases | Where-Object {
        $manifestPath = Join-Path $_.FullName "case.json"
        $casePath = Convert-ToRepoPath $_.FullName
        $inputPath = Convert-ToRepoPath (Get-CaseInputPath $manifestPath)
        $changedPaths | Where-Object { $_.StartsWith($casePath, [StringComparison]::Ordinal) -or $_ -eq $inputPath }
    })
}

if ($Limit -gt 0) {
    $cases = @($cases | Select-Object -First $Limit)
}

if ($List) {
    $cases | ForEach-Object { $_.Name }
    return
}

if ($cases.Count -eq 0) {
    throw "Visual family '$Family' did not match any cases."
}

Write-Host ("Visual family: {0} ({1} cases)" -f $Family, $cases.Count)
$rows = @()
foreach ($case in $cases) {
    $manifest = Join-Path $case.FullName "case.json"
    Write-Host ("==> {0}" -f $case.Name)
    $passed = $true
    $errorMessage = ""
    try {
        & (Join-Path $PSScriptRoot "CheckVisualCase.ps1") -Case $manifest
    }
    catch {
        $passed = $false
        $errorMessage = $_.Exception.Message
        Write-Host ("FAILED {0}: {1}" -f $case.Name, $errorMessage)
    }

    $rows += New-ReportRow $case $passed $errorMessage
}

New-Item -ItemType Directory -Force -Path $reportRoot | Out-Null
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$report = [pscustomobject]@{
    family = $Family
    generatedAt = (Get-Date).ToString("o")
    caseCount = $rows.Count
    passedCount = @($rows | Where-Object { $_.passed -eq $true }).Count
    failedCount = @($rows | Where-Object { $_.passed -ne $true }).Count
    cases = $rows
}

$json = $report | ConvertTo-Json -Depth 8
$reportJsonPath = Join-Path $reportRoot ("{0}.json" -f $Family)
$reportCsvPath = Join-Path $reportRoot ("{0}.csv" -f $Family)
$timestampedJsonPath = Join-Path $reportRoot ("{0}-{1}.json" -f $Family, $runId)
$timestampedCsvPath = Join-Path $reportRoot ("{0}-{1}.csv" -f $Family, $runId)
Set-Content -LiteralPath $reportJsonPath -Value $json
Set-Content -LiteralPath $timestampedJsonPath -Value $json
$rows | Export-Csv -NoTypeInformation -LiteralPath $reportCsvPath
$rows | Export-Csv -NoTypeInformation -LiteralPath $timestampedCsvPath

if ($UpdateCatalog) {
    foreach ($row in $rows) {
        $supportCatalog[$row.id] = [pscustomobject]@{
            id = $row.id
            family = $row.family
            kind = $row.kind
            classification = $row.classification
            status = $row.status
            lastRunAt = $report.generatedAt
            lastRunPath = $row.runPath
            minStructuralSimilarity = $row.minStructuralSimilarity
            minForegroundColorHistogramCorrelation = $row.minForegroundColorHistogramCorrelation
            requiredStructuralSimilarity = $row.requiredStructuralSimilarity
            requiredForegroundColorHistogramCorrelation = $row.requiredForegroundColorHistogramCorrelation
            maxMeanAbsoluteError = $row.maxMeanAbsoluteError
            maxChangedPixelRatioAtThreshold16 = $row.maxChangedPixelRatioAtThreshold16
        }
    }

    $catalog = [pscustomobject]@{
        generatedAt = (Get-Date).ToString("o")
        cases = @($supportCatalog.Values | Sort-Object id)
    }
    $catalogDirectory = Split-Path -Parent $supportCatalogPath
    New-Item -ItemType Directory -Force -Path $catalogDirectory | Out-Null
    Set-Content -LiteralPath $supportCatalogPath -Value ($catalog | ConvertTo-Json -Depth 8)
}

Write-Host "Visual family report: $reportJsonPath"
if ($report.failedCount -ne 0) {
    throw ("Visual family '{0}' had {1} failing case(s)." -f $Family, $report.failedCount)
}
