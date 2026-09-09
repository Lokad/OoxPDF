param(
    [Parameter(Mandatory = $true)]
    [string] $Baseline,

    [Parameter(Mandatory = $true)]
    [string] $Current,

    [double] $MaxStructuralSimilarityDrop = 0.005,

    [double] $MaxColorHistogramDrop = 0.02,

    [double] $MaxForegroundRecallDrop = 0.02,

    [double] $MaxMeanAbsoluteErrorIncrease = 0.02,

    [double] $MaxChangedPixelRatioIncrease = 0.002
)

$ErrorActionPreference = "Stop"

function Read-Report([string] $path) {
    $resolved = (Resolve-Path -LiteralPath $path).Path
    $report = Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json
    if ($null -eq $report -or $null -eq $report.cases) {
        throw "Visual report '$path' is empty or truncated: missing 'cases' array."
    }

    $cases = @($report.cases)
    if ($cases.Count -eq 0) {
        throw "Visual report '$path' contains zero cases: incomplete measurement."
    }

    return $report
}

function Get-CaseId($case) {
    return [string]$case.id
}

function Assert-UniqueIds($report, [string] $label) {
    $seen = @{}
    $duplicates = @()
    foreach ($case in @($report.cases)) {
        $id = Get-CaseId $case
        if ([string]::IsNullOrWhiteSpace($id)) {
            throw "Visual report '$label' contains a case with a missing or empty id: incomplete measurement."
        }

        if ($seen.ContainsKey($id)) {
            $duplicates += $id
        }
        else {
            $seen[$id] = $true
        }
    }

    if ($duplicates.Count -ne 0) {
        throw ("Visual report '{0}' contains duplicate case ids: {1}." -f $label, (($duplicates | Sort-Object -Unique) -join ", "))
    }
}

function Build-CaseMap($report) {
    $map = @{}
    foreach ($case in @($report.cases)) {
        $map[[string]$case.id] = $case
    }

    return $map
}

function To-FiniteDouble($value) {
    if ($null -eq $value -or $value -eq "") {
        return $null
    }

    try {
        $number = [double]$value
    }
    catch {
        return $null
    }

    if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {
        return $null
    }

    return $number
}

function Get-CaseCompletenessError($case) {
    if ($case.passed -ne $true) {
        return $null
    }

    foreach ($field in @("minStructuralSimilarity", "minForegroundColorHistogramCorrelation", "maxMeanAbsoluteError", "maxChangedPixelRatioAtThreshold16")) {
        if ($null -eq (To-FiniteDouble $case.$field)) {
            return ("passed case '{0}' is missing finite metric '{1}': incomplete measurement." -f [string]$case.id, $field)
        }
    }

    if ($null -eq $case.pageCount) {
        return ("passed case '{0}' is missing 'pageCount': incomplete measurement." -f [string]$case.id)
    }

    try {
        $pages = [int]$case.pageCount
    }
    catch {
        return ("passed case '{0}' has a non-integer 'pageCount': incomplete measurement." -f [string]$case.id)
    }

    if ($pages -lt 1) {
        return ("passed case '{0}' reports pageCount {1}: incomplete measurement." -f [string]$case.id, $pages)
    }

    return $null
}

$baselineReport = Read-Report $Baseline
$currentReport = Read-Report $Current
Assert-UniqueIds $baselineReport $Baseline
Assert-UniqueIds $currentReport $Current

if (($null -ne $baselineReport.family -or $null -ne $currentReport.family) -and ([string]$baselineReport.family -ne [string]$currentReport.family)) {
    throw ("Incomparable visual reports: baseline family '{0}' differs from current family '{1}'. Compare matching families only." -f [string]$baselineReport.family, [string]$currentReport.family)
}

$baselineMap = Build-CaseMap $baselineReport
$currentMap = Build-CaseMap $currentReport

$regressions = @()
$incomplete = @()
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
    foreach ($field in @("family", "kind")) {
        if (($null -ne $before.$field -or $null -ne $after.$field) -and ([string]$before.$field -ne [string]$after.$field)) {
            throw ("Incomparable case '{0}': baseline {1} '{2}' differs from current {1} '{3}'. Regenerate both reports from matching inputs." -f $id, $field, [string]$before.$field, [string]$after.$field)
        }
    }

    $beforeError = Get-CaseCompletenessError $before
    $afterError = Get-CaseCompletenessError $after
    if ($null -ne $afterError) {
        $incomplete += ("current: " + $afterError)
    }

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

    if ($before.passed -eq $true -and $after.passed -eq $true -and ($null -ne $beforeError -or $null -ne $afterError)) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "completeness"
            baseline = $(if ($null -eq $beforeError) { "complete" } else { "incomplete" })
            current = $(if ($null -eq $afterError) { "complete" } else { "incomplete" })
            delta = $null
            limit = $null
        }
        continue
    }

    $beforePages = try { [int]$before.pageCount } catch { $null }
    $afterPages = try { [int]$after.pageCount } catch { $null }
    if ($null -ne $beforePages -and $null -ne $afterPages -and $beforePages -ne $afterPages) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "pageCount"
            baseline = $beforePages
            current = $afterPages
            delta = $afterPages - $beforePages
            limit = 0
        }
    }

    $beforeSsim = To-FiniteDouble $before.minStructuralSimilarity
    $afterSsim = To-FiniteDouble $after.minStructuralSimilarity
    if ($null -eq $beforeSsim -or $null -eq $afterSsim) {
        if ($before.passed -eq $true -and $after.passed -eq $true) {
            $regressions += [pscustomobject]@{
                id = $id
                metric = "minStructuralSimilarity"
                baseline = $(if ($null -eq $beforeSsim) { "missing" } else { $beforeSsim })
                current = $(if ($null -eq $afterSsim) { "missing" } else { $afterSsim })
                delta = $null
                limit = -$MaxStructuralSimilarityDrop
            }
        }
    }
    elseif ($beforeSsim - $afterSsim -gt $MaxStructuralSimilarityDrop) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "minStructuralSimilarity"
            baseline = $beforeSsim
            current = $afterSsim
            delta = $afterSsim - $beforeSsim
            limit = -$MaxStructuralSimilarityDrop
        }
    }

    $beforeHist = To-FiniteDouble $before.minForegroundColorHistogramCorrelation
    $afterHist = To-FiniteDouble $after.minForegroundColorHistogramCorrelation
    if ($null -eq $beforeHist -or $null -eq $afterHist) {
        if ($before.passed -eq $true -and $after.passed -eq $true) {
            $regressions += [pscustomobject]@{
                id = $id
                metric = "minForegroundColorHistogramCorrelation"
                baseline = $(if ($null -eq $beforeHist) { "missing" } else { $beforeHist })
                current = $(if ($null -eq $afterHist) { "missing" } else { $afterHist })
                delta = $null
                limit = -$MaxColorHistogramDrop
            }
        }
    }
    elseif ($beforeHist - $afterHist -gt $MaxColorHistogramDrop) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "minForegroundColorHistogramCorrelation"
            baseline = $beforeHist
            current = $afterHist
            delta = $afterHist - $beforeHist
            limit = -$MaxColorHistogramDrop
        }
    }

    # Foreground recall is newer than archived reports: a missing value on
    # either side is skipped (not a regression); only a finite drop beyond
    # the limit fails. Null means the metric did not apply (empty reference).
    $beforeRecall = To-FiniteDouble $before.minForegroundRecall
    $afterRecall = To-FiniteDouble $after.minForegroundRecall
    if ($null -ne $beforeRecall -and $null -ne $afterRecall -and ($beforeRecall - $afterRecall -gt $MaxForegroundRecallDrop)) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "minForegroundRecall"
            baseline = $beforeRecall
            current = $afterRecall
            delta = $afterRecall - $beforeRecall
            limit = -$MaxForegroundRecallDrop
        }
    }

    $beforeMae = To-FiniteDouble $before.maxMeanAbsoluteError
    $afterMae = To-FiniteDouble $after.maxMeanAbsoluteError
    if ($null -eq $beforeMae -or $null -eq $afterMae) {
        if ($before.passed -eq $true -and $after.passed -eq $true) {
            $regressions += [pscustomobject]@{
                id = $id
                metric = "maxMeanAbsoluteError"
                baseline = $(if ($null -eq $beforeMae) { "missing" } else { $beforeMae })
                current = $(if ($null -eq $afterMae) { "missing" } else { $afterMae })
                delta = $null
                limit = $MaxMeanAbsoluteErrorIncrease
            }
        }
    }
    elseif ($afterMae - $beforeMae -gt $MaxMeanAbsoluteErrorIncrease) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "maxMeanAbsoluteError"
            baseline = $beforeMae
            current = $afterMae
            delta = $afterMae - $beforeMae
            limit = $MaxMeanAbsoluteErrorIncrease
        }
    }

    $beforeChanged = To-FiniteDouble $before.maxChangedPixelRatioAtThreshold16
    $afterChanged = To-FiniteDouble $after.maxChangedPixelRatioAtThreshold16
    if ($null -eq $beforeChanged -or $null -eq $afterChanged) {
        if ($before.passed -eq $true -and $after.passed -eq $true) {
            $regressions += [pscustomobject]@{
                id = $id
                metric = "maxChangedPixelRatioAtThreshold16"
                baseline = $(if ($null -eq $beforeChanged) { "missing" } else { $beforeChanged })
                current = $(if ($null -eq $afterChanged) { "missing" } else { $afterChanged })
                delta = $null
                limit = $MaxChangedPixelRatioIncrease
            }
        }
    }
    elseif ($afterChanged - $beforeChanged -gt $MaxChangedPixelRatioIncrease) {
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

if ($incomplete.Count -ne 0) {
    Write-Host "Incomplete current measurements:"
    $incomplete | Sort-Object -Unique | ForEach-Object { Write-Host (" - " + $_) }
}

if ($regressions.Count -eq 0 -and $incomplete.Count -eq 0) {
    Write-Host "No visual report regressions detected."
    return
}

if ($regressions.Count -ne 0) {
    $regressions | Format-Table -AutoSize
    throw ("Detected {0} visual report regression(s)." -f $regressions.Count)
}

throw ("Current visual report has {0} incomplete measurement(s); refusing to report no regressions." -f $incomplete.Count)