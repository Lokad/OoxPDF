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

function Build-CaseMap($report, [string] $label) {
    $map = @{}
    foreach ($case in @($report.cases)) {
        $id = [string]$case.id
        if ([string]::IsNullOrWhiteSpace($id)) {
            throw "Visual report '$label' contains a case with a missing or empty id: incomplete measurement."
        }
        if ($map.ContainsKey($id)) {
            throw "Visual report '$label' contains duplicate case id '$id'."
        }
        $map[$id] = $case
    }
    return $map
}

function To-FiniteDouble($value) {
    if ($null -eq $value -or [string]$value -eq "") {
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

function To-PositiveInteger($value) {
    $number = To-FiniteDouble $value
    if ($value -is [bool] -or $null -eq $number -or $number -lt 1 -or
        $number -gt [int]::MaxValue -or $number -ne [Math]::Truncate($number)) {
        return $null
    }
    return [int]$number
}

$identityHashes = @("caseSha256", "inputSha256", "referencePdfSha256", "visualDiffSha256", "rasterizerSha256", "pdfiumPinSha256")
$comparableFields = $identityHashes + @("family", "kind", "dpi", "docxMarkup", "docxMarkupGeometry",
    "requiredStructuralSimilarity", "requiredForegroundColorHistogramCorrelation", "requiredForegroundRecall")

function Assert-CaseIdentity($case, [string] $side) {
    foreach ($name in $identityHashes) {
        if ($case.$name -isnot [string] -or $case.$name -notmatch '\A[0-9a-fA-F]{64}\z') {
            throw "Incomparable $side case '$($case.id)': missing or invalid $name. Regenerate the visual report."
        }
    }
    if ($null -eq (To-PositiveInteger $case.dpi)) {
        throw "Incomparable $side case '$($case.id)': missing or invalid dpi. Regenerate the visual report."
    }
    foreach ($name in @("requiredStructuralSimilarity", "requiredForegroundColorHistogramCorrelation")) {
        if ($null -eq (To-FiniteDouble $case.$name)) {
            throw "Incomparable $side case '$($case.id)': missing or invalid $name. Regenerate the visual report."
        }
    }
}

function Get-CaseCompletenessError($case, $baselineCase, [bool] $bound) {
    if (!$bound -and $case.passed -ne $true) {
        return $null
    }

    $fields = @("minStructuralSimilarity", "minForegroundColorHistogramCorrelation", "maxMeanAbsoluteError", "maxChangedPixelRatioAtThreshold16")
    $requiresRecall = $null -ne (To-FiniteDouble $case.requiredForegroundRecall) -or
        ($null -ne (To-FiniteDouble $baselineCase.requiredForegroundRecall) -and
         $null -ne (To-FiniteDouble $baselineCase.minForegroundRecall))
    if ($requiresRecall) { $fields += 'minForegroundRecall' }
    foreach ($field in $fields) {
        if ($null -eq (To-FiniteDouble $case.$field)) {
            return ("case '{0}' is missing finite metric '{1}': incomplete measurement." -f [string]$case.id, $field)
        }
    }

    $rows = To-PositiveInteger $case.pageCount
    if ($null -eq $rows) {
        return ("case '{0}' requires a positive integer 'pageCount': incomplete measurement." -f [string]$case.id)
    }
    if ($bound) {
        $referencePages = To-PositiveInteger $case.referencePageCount
        $candidatePages = To-PositiveInteger $case.candidatePageCount
        if ($null -eq $referencePages -or $null -eq $candidatePages -or
            $rows -ne [Math]::Max($referencePages, $candidatePages)) {
            return ("case '{0}' requires separate page counts consistent with its comparison rows: incomplete measurement." -f [string]$case.id)
        }
    }

    return $null
}

$baselineReport = Read-Report $Baseline
$currentReport = Read-Report $Current
$schema = "ooxpdf-visual-family-report/2"
foreach ($report in @($baselineReport, $currentReport)) {
    if ($report.schema -ceq "ooxpdf-visual-family-report/1") {
        throw "Historical visual family report lacks separate candidate and reference page counts. Regenerate it for comparison."
    }
    if ($null -ne $report.schema -and $report.schema -cne $schema) {
        throw "Unknown visual report schema '$($report.schema)'."
    }
}
if ($baselineReport.schema -cne $currentReport.schema) {
    throw "Incomparable visual reports: both must record the comparison inputs. Regenerate the historical report."
}
$bound = $currentReport.schema -ceq $schema
if ($bound) { $comparableFields += 'referencePageCount' }
$pageField = if ($bound) { 'candidatePageCount' } else { 'pageCount' }

if (($null -ne $baselineReport.family -or $null -ne $currentReport.family) -and ([string]$baselineReport.family -ne [string]$currentReport.family)) {
    throw ("Incomparable visual reports: baseline family '{0}' differs from current family '{1}'. Compare matching families only." -f [string]$baselineReport.family, [string]$currentReport.family)
}

$baselineMap = Build-CaseMap $baselineReport $Baseline
$currentMap = Build-CaseMap $currentReport $Current

$metricRules = @(
    @{Name = "minStructuralSimilarity"; Drop = $true; Tolerance = $MaxStructuralSimilarityDrop}
    @{Name = "minForegroundColorHistogramCorrelation"; Drop = $true; Tolerance = $MaxColorHistogramDrop}
    @{Name = "minForegroundRecall"; Drop = $true; Tolerance = $MaxForegroundRecallDrop}
    @{Name = "maxMeanAbsoluteError"; Drop = $false; Tolerance = $MaxMeanAbsoluteErrorIncrease}
    @{Name = "maxChangedPixelRatioAtThreshold16"; Drop = $false; Tolerance = $MaxChangedPixelRatioIncrease}
)

$regressions = @()
$incomplete = @()
foreach ($id in $currentMap.Keys | Sort-Object) {
    $after = $currentMap[$id]
    $afterError = Get-CaseCompletenessError $after $baselineMap[$id] $bound
    if ($null -ne $afterError) {
        $incomplete += "current: $afterError"
    }
    if ($bound -and ($null -eq $afterError -or $baselineMap.ContainsKey($id))) {
        Assert-CaseIdentity $after "current"
    }
}

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
    if ($bound) {
        Assert-CaseIdentity $before "baseline"
    }
    foreach ($field in $comparableFields) {
        if (($null -ne $before.$field -or $null -ne $after.$field) -and ([string]$before.$field -ne [string]$after.$field)) {
            throw ("Incomparable case '{0}': baseline {1} '{2}' differs from current {1} '{3}'. Regenerate both reports from matching inputs." -f $id, $field, [string]$before.$field, [string]$after.$field)
        }
    }

    $beforeError = Get-CaseCompletenessError $before $null $bound
    if ($bound -and $null -ne $beforeError) {
        $incomplete += "baseline: $beforeError"
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

    if (!$bound -and $before.passed -eq $true -and $after.passed -eq $true -and $null -ne $beforeError) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = "completeness"
            baseline = "incomplete"
            current = "reported passing"
            delta = $null
            limit = $null
        }
        continue
    }

    $beforePages = To-PositiveInteger $before.$pageField
    $afterPages = To-PositiveInteger $after.$pageField
    if ($null -ne $beforePages -and $null -ne $afterPages -and $beforePages -ne $afterPages) {
        $regressions += [pscustomobject]@{
            id = $id
            metric = $pageField
            baseline = $beforePages
            current = $afterPages
            delta = $afterPages - $beforePages
            limit = 0
        }
    }

    foreach ($rule in $metricRules) {
        $beforeValue = To-FiniteDouble $before.($rule.Name)
        $afterValue = To-FiniteDouble $after.($rule.Name)
        if ($null -eq $beforeValue -or $null -eq $afterValue) { continue }

        $delta = $afterValue - $beforeValue
        $worsening = if ($rule.Drop) { -$delta } else { $delta }
        if ($worsening -gt $rule.Tolerance) {
            $regressions += [pscustomobject]@{
                id = $id
                metric = $rule.Name
                baseline = $beforeValue
                current = $afterValue
                delta = $delta
                limit = $(if ($rule.Drop) { -$rule.Tolerance } else { $rule.Tolerance })
            }
        }
    }
}

if ($incomplete.Count -ne 0) {
    Write-Host "Incomplete measurements:"
    $incomplete | Sort-Object -Unique | ForEach-Object { Write-Host (" - " + $_) }
}

if ($regressions.Count -eq 0 -and $incomplete.Count -eq 0) {
    if ($bound) {
        Write-Host "No visual report regressions detected."
    }
    else {
        Write-Warning "Historical visual comparison found no metric regressions. These reports do not bind the comparison inputs."
    }
    return
}

if ($regressions.Count -ne 0) {
    $regressions | Format-Table -AutoSize
    throw ("Detected {0} visual report regression(s)." -f $regressions.Count)
}

throw ("Visual reports have {0} incomplete measurement(s); refusing to report no regressions." -f $incomplete.Count)
