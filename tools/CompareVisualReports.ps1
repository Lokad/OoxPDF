param(
    [Parameter(Mandatory = $true)]
    [string] $Baseline,

    [Parameter(Mandatory = $true)]
    [string] $Current,

    [double] $MaxStructuralSimilarityDrop = 0.005,

    [double] $MaxColorHistogramDrop = 0.02,

    [double] $MaxMeanAbsoluteErrorIncrease = 0.02,

    [double] $MaxChangedPixelRatioIncrease = 0.002
)

$ErrorActionPreference = "Stop"

function Read-Report([string] $path) {
    $resolved = (Resolve-Path -LiteralPath $path).Path
    return Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json
}

function Build-CaseMap($report) {
    $map = @{}
    foreach ($case in @($report.cases)) {
        $map[$case.id] = $case
    }

    return $map
}

function To-NullableDouble($value) {
    if ($null -eq $value -or $value -eq "") {
        return $null
    }

    return [double]$value
}

$baselineReport = Read-Report $Baseline
$currentReport = Read-Report $Current
$baselineMap = Build-CaseMap $baselineReport
$currentMap = Build-CaseMap $currentReport

$regressions = @()
foreach ($id in $baselineMap.Keys | Sort-Object) {
    if (-not $currentMap.ContainsKey($id)) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "case"
            baseline = "present"
            current = "missing"
            delta = $null
            limit = $null
        }
        continue
    }

    $before = $baselineMap[$id]
    $after = $currentMap[$id]
    if ($before.passed -eq $true -and $after.passed -ne $true) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "passed"
            baseline = $before.passed
            current = $after.passed
            delta = $null
            limit = $null
        }
    }

    $beforeSsim = To-NullableDouble $before.minStructuralSimilarity
    $afterSsim = To-NullableDouble $after.minStructuralSimilarity
    if ($null -ne $beforeSsim -and $null -ne $afterSsim -and $beforeSsim - $afterSsim -gt $MaxStructuralSimilarityDrop) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "minStructuralSimilarity"
            baseline = $beforeSsim
            current = $afterSsim
            delta = $afterSsim - $beforeSsim
            limit = -$MaxStructuralSimilarityDrop
        }
    }

    $beforeHist = To-NullableDouble $before.minForegroundColorHistogramCorrelation
    $afterHist = To-NullableDouble $after.minForegroundColorHistogramCorrelation
    if ($null -ne $beforeHist -and $null -ne $afterHist -and $beforeHist - $afterHist -gt $MaxColorHistogramDrop) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "minForegroundColorHistogramCorrelation"
            baseline = $beforeHist
            current = $afterHist
            delta = $afterHist - $beforeHist
            limit = -$MaxColorHistogramDrop
        }
    }

    $beforeMae = To-NullableDouble $before.maxMeanAbsoluteError
    $afterMae = To-NullableDouble $after.maxMeanAbsoluteError
    if ($null -ne $beforeMae -and $null -ne $afterMae -and $afterMae - $beforeMae -gt $MaxMeanAbsoluteErrorIncrease) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "maxMeanAbsoluteError"
            baseline = $beforeMae
            current = $afterMae
            delta = $afterMae - $beforeMae
            limit = $MaxMeanAbsoluteErrorIncrease
        }
    }

    $beforeChanged = To-NullableDouble $before.maxChangedPixelRatioAtThreshold16
    $afterChanged = To-NullableDouble $after.maxChangedPixelRatioAtThreshold16
    if ($null -ne $beforeChanged -and $null -ne $afterChanged -and $afterChanged - $beforeChanged -gt $MaxChangedPixelRatioIncrease) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "maxChangedPixelRatioAtThreshold16"
            baseline = $beforeChanged
            current = $afterChanged
            delta = $afterChanged - $beforeChanged
            limit = $MaxChangedPixelRatioIncrease
        }
    }
}

if ($regressions.Count -eq 0) {
    Write-Host "No visual report regressions detected."
    return
}

$regressions | Format-Table -AutoSize
throw ("Detected {0} visual report regression(s)." -f $regressions.Count)
