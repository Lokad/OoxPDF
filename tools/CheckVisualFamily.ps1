[CmdletBinding()]
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory = $true)]
    [string] $Family,

    [switch] $List,

    [switch] $OnlyUnsupported,

    [switch] $UpdateCatalog,

    [switch] $CacheOnlyReference,

    [string] $CliPath,

    [string] $RasterizerPath,
    [string] $VisualDiffPath,
    [string] $PdfInspectorPath,

    [switch] $SkipBuild,

    [string[]] $CasePattern = @(),

    [string[]] $Tag = @(),

    [int] $Limit = 0
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'VisualCaseGates.ps1')

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

function Use-SupportCatalogLock([scriptblock] $Action) {
    $directory = Split-Path -Parent $supportCatalogPath
    $physical = [IO.Path]::GetPathRoot($directory)
    foreach ($segment in ($directory.Substring($physical.Length) -split '[\\/]')) {
        if (!$segment) { continue }
        $next = Join-Path $physical $segment
        $entry = Get-Item -LiteralPath $next -Force -ErrorAction Ignore
        $target = if ($null -ne $entry) { $entry.ResolveLinkTarget($true) }
        $physical = if ($null -ne $target) { $target.FullName } else { $next }
    }
    if ($IsWindows) { $physical = $physical.ToUpperInvariant() }
    $key = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($physical)))
    $mutex = [Threading.Mutex]::new($false, 'Global\OoxPdfSupportCatalog-' + $key)
    try {
        if (!$mutex.WaitOne([TimeSpan]::FromMinutes(5))) {
            throw 'Timed out waiting for the support-catalog lock.'
        }
        try { & $Action }
        finally { $mutex.ReleaseMutex() }
    }
    finally { $mutex.Dispose() }
}

function Convert-ToRepoPath([string] $path) {
    return (Resolve-Path -LiteralPath $path).Path.Substring($repoRoot.Length).TrimStart("\", "/").Replace("\", "/")
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

function New-ReportRow($caseBinding, [bool]$passed, [string]$errorMessage, [string]$explicitRunRoot) {
    $caseManifest = $caseBinding.Manifest
    $latestRun = if (-not [string]::IsNullOrWhiteSpace($explicitRunRoot) -and (Test-Path -LiteralPath $explicitRunRoot)) { Get-Item -LiteralPath $explicitRunRoot } else { $null }
    $runInfo = $null
    $metrics = @()
    $diagnostics = @()
    if ($latestRun -ne $null) {
        $runInfoPath = Join-Path $latestRun.FullName "run-info.json"
        if (Test-Path -LiteralPath $runInfoPath) {
            $runInfo = Get-Content -Raw -LiteralPath $runInfoPath | ConvertFrom-Json
            if ($runInfo.RunId -cne $latestRun.Name -or $runInfo.CaseId -cne $caseManifest.id -or
                $runInfo.CaseSha256 -ne $caseBinding.Sha256 -or
                (Get-FileHash -LiteralPath $caseBinding.Path).Hash -ne $caseBinding.Sha256 -or
                $runInfo.CliSha256 -ne $cliHash -or $runInfo.LibrarySha256 -ne $libraryHash) {
                throw "Visual case identity changed during the family run. $($caseManifest.id)"
            }
            if (($rasterizerBinding -and $runInfo.RasterizerSha256 -ne $rasterizerBinding.Hash) -or
                ($visualDiffBinding -and $runInfo.VisualDiffSha256 -ne $visualDiffBinding.Hash) -or
                ($inspectorBinding -and $runInfo.PdfInspectorSha256 -ne $inspectorBinding.Hash)) {
                throw "Visual case tool binding differs from the family run. $($caseManifest.id)"
            }
        }
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
    if ($passed -and $null -eq $runInfo) {
        throw "Passing case '$($caseManifest.id)' did not record its visual run identity."
    }

    $maeValues = @($metrics | Where-Object { $_.MeanAbsoluteError -ne $null } | ForEach-Object { [double]$_.MeanAbsoluteError })
    $changed16Values = @($metrics | Where-Object { $_.ChangedPixelRatioAtThreshold16 -ne $null } | ForEach-Object { [double]$_.ChangedPixelRatioAtThreshold16 })
    $ssimValues = @($metrics | Where-Object { $_.StructuralSimilarity -ne $null } | ForEach-Object { [double]$_.StructuralSimilarity })
    $histValues = @($metrics | Where-Object { $_.ForegroundColorHistogramCorrelation -ne $null } | ForEach-Object { [double]$_.ForegroundColorHistogramCorrelation })
    $recallValues = @($metrics | Where-Object { $_.ForegroundRecall -ne $null } | ForEach-Object { [double]$_.ForegroundRecall })
    $dimensionMismatchCount = @($metrics | Where-Object { $_.DimensionsMatch -ne $true }).Count
    $minStructuralSimilarity = if ($ssimValues.Count -eq 0) { $null } else { ($ssimValues | Measure-Object -Minimum).Minimum }
    $minForegroundColorHistogramCorrelation = if ($histValues.Count -eq 0) { $null } else { ($histValues | Measure-Object -Minimum).Minimum }
    $minForegroundRecall = if ($recallValues.Count -eq 0) { $null } else { ($recallValues | Measure-Object -Minimum).Minimum }
    $support = $caseManifest.support
    $familySupport = $familyManifest.support
    $requiredStructuralSimilarity = Get-ManifestDouble $support "minStructuralSimilarity" (Get-ManifestDouble $familySupport "minStructuralSimilarity" 0.995)
    $requiredColorHistogram = Get-ManifestDouble $support "minForegroundColorHistogramCorrelation" (Get-ManifestDouble $familySupport "minForegroundColorHistogramCorrelation" 0.98)
    $requiredForegroundRecall = Get-ManifestDouble $support "minForegroundRecall" (Get-ManifestDouble $familySupport "minForegroundRecall" $null)
    $status = if (-not $passed) {
        "unsupported"
    }
    elseif ($minStructuralSimilarity -eq $null -or
        $minForegroundColorHistogramCorrelation -eq $null -or
        $minStructuralSimilarity -lt $requiredStructuralSimilarity -or
        $minForegroundColorHistogramCorrelation -lt $requiredColorHistogram -or
        ($requiredForegroundRecall -ne $null -and ($minForegroundRecall -eq $null -or $minForegroundRecall -lt $requiredForegroundRecall))) {
        "needs-review"
    }
    else {
        "supported"
    }

    [pscustomobject]@{
        id = $caseManifest.id
        family = $Family
        kind = $caseManifest.kind
        caseSha256 = $runInfo.CaseSha256
        inputSha256 = $runInfo.InputSha256
        referencePdfSha256 = $runInfo.ReferencePdfSha256
        visualDiffSha256 = $runInfo.VisualDiffSha256
        rasterizerSha256 = $runInfo.RasterizerSha256
        pdfInspectorSha256 = $runInfo.PdfInspectorSha256
        pdfiumPinSha256 = $runInfo.PdfiumPinSha256
        dpi = $runInfo.Dpi
        docxMarkupGeometry = $runInfo.DocxMarkupGeometry
        docxMarkup = if ($caseManifest.kind -eq "docx") {
            if ($caseManifest.PSObject.Properties.Name -contains "docxMarkup" -and -not [string]::IsNullOrWhiteSpace([string]$caseManifest.docxMarkup)) {
                [string]$caseManifest.docxMarkup
            }
            else {
                "final"
            }
        }
        else {
            $null
        }
        classification = Get-CaseClassification $caseManifest
        passed = $passed
        status = $status
        runPath = if ($latestRun -ne $null) { Convert-ToRepoPath $latestRun.FullName } else { $null }
        referencePageCount = $runInfo.ReferencePageCount
        candidatePageCount = $runInfo.CandidatePageCount
        # VisualDiff emits one row for each position on either side.
        pageCount = $metrics.Count
        dimensionMismatchCount = $dimensionMismatchCount
        diagnosticsCount = $diagnostics.Count
        diagnosticIds = @($diagnostics | ForEach-Object { $_.Id } | Sort-Object -Unique) -join ";"
        minStructuralSimilarity = $minStructuralSimilarity
        minForegroundColorHistogramCorrelation = $minForegroundColorHistogramCorrelation
        minForegroundRecall = $minForegroundRecall
        requiredStructuralSimilarity = $requiredStructuralSimilarity
        requiredForegroundColorHistogramCorrelation = $requiredColorHistogram
        requiredForegroundRecall = $requiredForegroundRecall
        maxMeanAbsoluteError = if ($maeValues.Count -eq 0) { $null } else { ($maeValues | Measure-Object -Maximum).Maximum }
        maxChangedPixelRatioAtThreshold16 = if ($changed16Values.Count -eq 0) { $null } else { ($changed16Values | Measure-Object -Maximum).Maximum }
        error = $errorMessage
    }
}

$patterns = @($familyManifest.casePatterns)
$excludePatterns = @($familyManifest.excludePatterns | Where-Object { $_ -is [string] -and $_.Length -ne 0 })
$cases = @(
    Get-ChildItem -LiteralPath $caseRoot -Directory |
        Where-Object {
            $caseName = $_.Name
            ($patterns | Where-Object { $caseName -like $_ }) -and
                -not ($excludePatterns | Where-Object { $caseName -like $_ })
        } |
        Sort-Object Name
)

if ($CasePattern.Count -ne 0) {
    $cases = @($cases | Where-Object {
        $caseName = $_.Name
        $CasePattern | Where-Object { $caseName -like $_ }
    })
}

$caseBindings = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
if ($Tag.Count -ne 0) {
    $cases = @($cases | Where-Object {
        $caseBinding = Read-VisualCaseManifest (Join-Path $_.FullName "case.json")
        $caseTags = @($caseBinding.Manifest.tags)
        $missingTags = @($Tag | Where-Object { $caseTags -notcontains $_ })
        if ($missingTags.Count -ne 0) { return $false }
        $caseBindings[$_.FullName] = $caseBinding
        return $true
    })
}

if ($OnlyUnsupported) {
    $supportCatalog = Use-SupportCatalogLock { Read-SupportCatalog }
    $cases = @($cases | Where-Object {
        $id = $_.Name
        -not $supportCatalog.ContainsKey($id) -or $supportCatalog[$id].status -ne "supported"
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

# Inspect complete selected manifests before preparing tools or starting any case.
$needsPdfInspection = $false
$inspectionThresholds = @('maxTextOperationPositionDelta', 'maxTextLineStartDelta',
    'maxGraphicsOperationBoundsDelta', 'maxChartTextStructurePositionDelta',
    'maxChartTextStructurePositionDeltaByKind')
foreach ($case in $cases) {
    if (!$caseBindings.ContainsKey($case.FullName)) {
        $caseBindings[$case.FullName] = Read-VisualCaseManifest (Join-Path $case.FullName 'case.json')
    }
    $manifest = $caseBindings[$case.FullName].Manifest
    $expected = $manifest.expected
    $gateIssue = Get-VisualChartGraphicsGateIssue $expected $manifest.id
    if ($gateIssue) { throw $gateIssue }
    if ((Test-VisualChartGraphicsGate $expected) -or
        $expected.candidateFontsMustBeEmbedded -eq $true -or
        $expected.candidateFontsMustHaveToUnicode -eq $true -or
        $expected.candidateFontsMustBeSubset -eq $true) {
        $needsPdfInspection = $true
    }
    foreach ($key in $inspectionThresholds) {
        if ($null -ne $expected.$key) { $needsPdfInspection = $true }
    }
}

Write-Host ("Visual family: {0} ({1} cases)" -f $Family, $cases.Count)
$rasterizerBinding = if ($RasterizerPath) { Get-FileHash -LiteralPath $RasterizerPath }
$visualDiffBinding = if ($VisualDiffPath) { Get-FileHash -LiteralPath $VisualDiffPath }
$inspectorBinding = if ($PdfInspectorPath) { Get-FileHash -LiteralPath $PdfInspectorPath }
# Bind every case to the same CLI. A supplied frozen CLI never builds the library;
# independent validation tools can still be built once unless -SkipBuild is set.
if (!$CliPath) {
    $CliPath = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/$Framework/Lokad.OoxPdf.Cli.dll"
    if (!$SkipBuild) {
        & (Join-Path $PSScriptRoot "EnsureDotnetBuild.ps1") -Framework $Framework -Project (Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj") -OutputDll $CliPath -Description "CLI"
    }
}
$CliPath = (Resolve-Path -LiteralPath $CliPath).Path
$libraryPath = Join-Path (Split-Path $CliPath) 'Lokad.OoxPdf.dll'
$cliHash = (Get-FileHash -LiteralPath $CliPath).Hash
$libraryHash = (Get-FileHash -LiteralPath $libraryPath).Hash
if (!$SkipBuild -and !$visualDiffBinding) {
    & (Join-Path $PSScriptRoot "EnsureDotnetBuild.ps1") -Framework $Framework -Project (Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj") -OutputDll (Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/bin/Debug/$Framework/Lokad.OoxPdf.VisualDiff.dll") -Description "VisualDiff"
}
if (!$SkipBuild -and !$rasterizerBinding) {
    & (Join-Path $PSScriptRoot "EnsureDotnetBuild.ps1") -Framework $Framework -Project (Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer/Lokad.OoxPdf.PdfiumRasterizer.csproj") -OutputDll (Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/$Framework/Lokad.OoxPdf.PdfiumRasterizer.dll") -Description "PDFium rasterizer"
}
# Only inspection gates need this tool. Cases consume it prebuilt, just like the
# rasterizer and VisualDiff, rather than invoking MSBuild for each inspected PDF.
if ($needsPdfInspection -and !$inspectorBinding) {
    $inspector = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/bin/Debug/$Framework/Lokad.OoxPdf.PdfInspect.dll"
    if (!$SkipBuild) {
        & (Join-Path $PSScriptRoot 'EnsureDotnetBuild.ps1') -Framework $Framework -Project (Join-Path $repoRoot 'tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj') -OutputDll $inspector -Description 'PDF inspect'
    }
    elseif (!(Test-Path -LiteralPath $inspector -PathType Leaf)) {
        throw ('PDF inspector is missing and -SkipBuild was requested: ' + $inspector)
    }
}
$rows = @()
foreach ($case in $cases) {
    $caseBinding = $caseBindings[$case.FullName]
    Write-Host ("==> {0}" -f $case.Name)
    # Explicit run binding: the family mints a collision-resistant run id and
    # hands the resulting run root to reporting instead of rediscovering the
    # latest run directory by mtime (which races concurrent and failed runs).
    $caseRunId = "{0}-p{1}-{2}" -f (Get-Date -Format "yyyyMMdd-HHmmss-fff"), $PID, ([System.Guid]::NewGuid().ToString("N").Substring(0, 8))
    $caseRunRoot = Join-Path $artifactRoot ([System.IO.Path]::Combine([string]$caseBinding.Manifest.id, $caseRunId))
    $passed = $true
    $errorMessage = ""
    try {
        $caseArgs = @{
            Case = $caseBinding.Path
            RunId = $caseRunId
            SkipBuild = $true
            CliPath = $CliPath
        }
        if ($CacheOnlyReference) {
            $caseArgs.CacheOnlyReference = $true
        }
        if ($rasterizerBinding) { $caseArgs.RasterizerPath = $rasterizerBinding.Path }
        if ($visualDiffBinding) { $caseArgs.VisualDiffPath = $visualDiffBinding.Path }
        if ($inspectorBinding) { $caseArgs.PdfInspectorPath = $inspectorBinding.Path }

        & (Join-Path $PSScriptRoot "CheckVisualCase.ps1") -Framework $Framework @caseArgs
    }
    catch {
        $passed = $false
        $errorMessage = $_.Exception.Message
        Write-Host ("FAILED {0}: {1}" -f $case.Name, $errorMessage)
    }

    $rows += New-ReportRow $caseBinding $passed $errorMessage $caseRunRoot
}
if ((Get-FileHash -LiteralPath $CliPath).Hash -ne $cliHash -or
    (Get-FileHash -LiteralPath $libraryPath).Hash -ne $libraryHash) {
    throw 'The family CLI or library changed during validation'
}
foreach ($binding in @($rasterizerBinding, $visualDiffBinding, $inspectorBinding)) {
    if ($binding -and (Get-FileHash -LiteralPath $binding.Path).Hash -ne $binding.Hash) {
        throw "Visual validation tool changed during the family: $($binding.Path)"
    }
}

New-Item -ItemType Directory -Force -Path $reportRoot | Out-Null
$runId = "{0}-p{1}-{2}" -f (Get-Date -Format "yyyyMMdd-HHmmss-fff"), $PID, ([System.Guid]::NewGuid().ToString("N").Substring(0, 8))
$report = [pscustomobject]@{
    schema = "ooxpdf-visual-family-report/2"
    family = $Family
    cliPath = $CliPath
    cliSha256 = $cliHash
    librarySha256 = $libraryHash
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
    Use-SupportCatalogLock {
        $supportCatalog = Read-SupportCatalog
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
                minForegroundRecall = $row.minForegroundRecall
                requiredStructuralSimilarity = $row.requiredStructuralSimilarity
                requiredForegroundColorHistogramCorrelation = $row.requiredForegroundColorHistogramCorrelation
                requiredForegroundRecall = $row.requiredForegroundRecall
                maxMeanAbsoluteError = $row.maxMeanAbsoluteError
                maxChangedPixelRatioAtThreshold16 = $row.maxChangedPixelRatioAtThreshold16
            }
        }

        $catalog = [pscustomobject]@{
            generatedAt = (Get-Date).ToString("o")
            cases = @($supportCatalog.Values | Sort-Object id)
        }
        $staging = $supportCatalogPath + '.staging-' + [Guid]::NewGuid().ToString('N')
        try {
            Set-Content -LiteralPath $staging -Value ($catalog | ConvertTo-Json -Depth 8)
            [IO.File]::Move($staging, $supportCatalogPath, $true)
        }
        finally { [IO.File]::Delete($staging) }
    }
}

Write-Host "Visual family report: $reportJsonPath"
if ($report.failedCount -ne 0) {
    throw ("Visual family '{0}' had {1} failing case(s)." -f $Family, $report.failedCount)
}
