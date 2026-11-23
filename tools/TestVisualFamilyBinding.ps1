# Exercise the real family orchestration with fake builds/conversions, no Office or library execution.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) { $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/visual-family-test-' + [Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo = Join-Path $OutputDirectory 'repository'
New-Item -ItemType Directory -Path (Join-Path $repo 'tools') -Force | Out-Null
$repo = (Resolve-Path -LiteralPath $repo).Path
foreach ($name in 'CheckVisualFamily.ps1', 'ValidateVisualCases.ps1', 'VisualCaseGates.ps1') {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $repo 'tools')
}
$buildLog = Join-Path $repo 'builds.jsonl'
$caseLog = Join-Path $repo 'cases.jsonl'
@'
param($Project, $OutputDll, $Description, $RecordPath, $Framework)
@{Project=$Project;Description=$Description} | ConvertTo-Json -Compress | Add-Content (Join-Path (Split-Path $PSScriptRoot) 'builds.jsonl')
New-Item -ItemType Directory -Path (Split-Path $OutputDll) -Force | Out-Null
[IO.File]::WriteAllText($OutputDll, 'fake CLI or validation tool')
if ($Description -eq 'CLI') { [IO.File]::WriteAllText((Join-Path (Split-Path $OutputDll) 'Lokad.OoxPdf.dll'), 'fake library') }
'@ | Set-Content (Join-Path $repo 'tools/EnsureDotnetBuild.ps1')
@'
param($Case, $RunId, [switch]$SkipBuild, $CliPath, [switch]$CacheOnlyReference, $Framework)
@{Case=$Case;RunId=$RunId;SkipBuild=[bool]$SkipBuild;CliPath=$CliPath;CacheOnlyReference=[bool]$CacheOnlyReference} |
    ConvertTo-Json -Compress | Add-Content (Join-Path (Split-Path $PSScriptRoot) 'cases.jsonl')
if (!$SkipBuild -or !(Test-Path -LiteralPath $CliPath)) { throw 'Family did not bind an existing CLI without per-case builds' }
$id=(Get-Content -LiteralPath $Case -Raw | ConvertFrom-Json).id
$run=Join-Path (Split-Path $PSScriptRoot) "artifacts/visual/$id/$RunId"
New-Item -ItemType Directory -Path $run -Force | Out-Null
@{RunId=$RunId;CaseId=$id;CaseSha256=(Get-FileHash -LiteralPath $Case).Hash;
    CliSha256=(Get-FileHash -LiteralPath $CliPath).Hash;
    LibrarySha256=(Get-FileHash -LiteralPath (Join-Path (Split-Path $CliPath) 'Lokad.OoxPdf.dll')).Hash} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'run-info.json')
if ($overlapCatalogUpdates -and $id -ceq 'probe-a') {
    & (Join-Path $PSScriptRoot 'CheckVisualFamily.ps1') -Family probe -CasePattern probe-b -CliPath $CliPath -SkipBuild -UpdateCatalog | Out-Null
}
'@ | Set-Content (Join-Path $repo 'tools/CheckVisualCase.ps1')
New-Item -ItemType Directory -Path (Join-Path $repo 'visual-cases/families') -Force | Out-Null
'{"casePatterns":["probe-*"]}' | Set-Content (Join-Path $repo 'visual-cases/families/probe.json')
foreach ($id in @('probe-a','probe-b')) {
    $directory = Join-Path $repo "visual-cases/cases/$id"
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    @{id=$id;kind='pptx';tags=@('approximate')} | ConvertTo-Json | Set-Content (Join-Path $directory 'case.json')
}
$frozenCli = Join-Path $repo 'frozen/Lokad.OoxPdf.Cli.dll'
New-Item -ItemType Directory -Path (Split-Path $frozenCli) -Force | Out-Null
[IO.File]::WriteAllText($frozenCli, 'frozen fake CLI')
$frozenLibrary = Join-Path (Split-Path $frozenCli) 'Lokad.OoxPdf.dll'
[IO.File]::WriteAllText($frozenLibrary, 'frozen fake library')
$script = Join-Path $repo 'tools/CheckVisualFamily.ps1'
$listed = @(& $script -Family probe -List -CliPath 'missing.dll')
if ($listed.Count -ne 2 -or (Test-Path $buildLog) -or (Test-Path $caseLog)) { throw 'Listing performed builds or required a runtime' }
& $script -Family probe -CliPath $frozenCli -CacheOnlyReference | Out-Null
$builds = @(Get-Content $buildLog | ConvertFrom-Json)
$cases = @(Get-Content $caseLog | ConvertFrom-Json)
if ($builds.Count -ne 2 -or @($builds | Where-Object Description -eq 'CLI').Count -or
    $cases.Count -ne 2 -or @($cases | Where-Object { $_.CliPath -ne $frozenCli -or !$_.CacheOnlyReference }).Count) {
    throw 'Frozen family binding rebuilt the CLI or lost case arguments'
}
$report = Get-Content (Join-Path $repo 'artifacts/visual/reports/probe.json') -Raw | ConvertFrom-Json
if ($report.cliSha256 -ne (Get-FileHash $frozenCli).Hash -or $report.librarySha256 -ne (Get-FileHash $frozenLibrary).Hash) {
    throw 'Family report omitted the actual runtime identity'
}
& $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null
if (@(Get-Content $buildLog).Count -ne 2 -or @(Get-Content $caseLog).Count -ne 4) { throw 'SkipBuild rebuilt tools or missed cases' }
& $script -Family probe | Out-Null
$builds = @(Get-Content $buildLog | ConvertFrom-Json)
$cases = @(Get-Content $caseLog | ConvertFrom-Json)
if ($builds.Count -ne 5 -or @($builds | Where-Object Description -eq 'CLI').Count -ne 1 -or $cases.Count -ne 6) {
    throw 'Default family execution no longer builds each required project once'
}
$defaultCli = $cases[-1].CliPath
& $script -Family probe -SkipBuild | Out-Null
$cases = @(Get-Content $caseLog | ConvertFrom-Json)
if (@(Get-Content $buildLog).Count -ne 5 -or $cases.Count -ne 8 -or $cases[-1].CliPath -ne $defaultCli) { throw 'Prebuilt default binding failed' }
# Every inspection capability prepares the tool once for both cases. False font
# flags leave raster-only families free of unused inspection builds.
$inspectionKeys=@('maxTextOperationPositionDelta','maxTextLineStartDelta','maxGraphicsOperationBoundsDelta',
    'maxChartGraphicsStructureBoundsDelta','maxChartGraphicsStructureBoundsDeltaByKind',
    'maxChartTextStructurePositionDelta','maxChartTextStructurePositionDeltaByKind',
    'candidateFontsMustBeEmbedded','candidateFontsMustHaveToUnicode','candidateFontsMustBeSubset')
foreach($key in $inspectionKeys) {
    foreach($id in @('probe-a','probe-b')) {
        $expected=@{}; $expected[$key]=if($key.StartsWith('candidateFonts')) {$true} else {0}
        @{id=$id;kind='docx';tags=@('approximate');expected=$expected} | ConvertTo-Json -Depth 4 |
            Set-Content (Join-Path $repo "visual-cases/cases/$id/case.json")
    }
    $before=@(Get-Content $buildLog).Count
    & $script -Family probe -CliPath $frozenCli | Out-Null
    $added=@(Get-Content $buildLog | Select-Object -Skip $before | ConvertFrom-Json)
    if($added.Count -ne 3 -or @($added | Where-Object Description -eq 'PDF inspect').Count -ne 1) {throw ('Inspection tool was not prepared exactly once: '+$key)}
}
$before=@(Get-Content $buildLog).Count
& $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null
if(@(Get-Content $buildLog).Count -ne $before) {throw 'An inspecting prebuilt family requested a build'}
$inspector=Join-Path $repo 'tools/Lokad.OoxPdf.PdfInspect/bin/Debug/net11.0/Lokad.OoxPdf.PdfInspect.dll'
if(!$inspector.StartsWith($repo+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {throw 'Unexpected synthetic inspector path'}
Remove-Item -LiteralPath $inspector
$caseCount=@(Get-Content $caseLog).Count
$failure=$null
try { & $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null } catch {$failure=$_.Exception.Message}
if($failure -notlike '*PDF inspector is missing*' -or @(Get-Content $buildLog).Count -ne $before -or @(Get-Content $caseLog).Count -ne $caseCount) {throw 'Missing prebuilt inspector did not fail before cases/builds'}
[IO.File]::WriteAllText($inspector,'prebuilt synthetic PDF inspector')
foreach($id in @('probe-a','probe-b')) {
    @{id=$id;kind='docx';tags=@('approximate');expected=@{candidateFontsMustBeEmbedded=$false}} | ConvertTo-Json -Depth 4 |
        Set-Content (Join-Path $repo "visual-cases/cases/$id/case.json")
}
& $script -Family probe -CliPath $frozenCli | Out-Null
if(@(Get-Content $buildLog).Count -ne $before+2) {throw 'A false font flag caused an unused inspection build'}

# Interleave disjoint writers deterministically. The nested family publishes
# probe-b while the first family still has probe-a in progress.
$overlapCatalogUpdates = $true
& $script -Family probe -CasePattern probe-a -Tag approximate -CliPath $frozenCli -SkipBuild -OnlyUnsupported -UpdateCatalog | Out-Null
$overlapCatalogUpdates = $false
$catalogPath = Join-Path $repo 'artifacts/visual/support-catalog.json'
$catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
if (($catalog.cases.id -join ',') -cne 'probe-a,probe-b') { throw 'An overlapping family update lost a catalog row' }
if (@(Get-ChildItem -LiteralPath (Split-Path -Parent $catalogPath) -Filter 'support-catalog.json.staging-*').Count) {
    throw 'A successful catalog update left staging files'
}

# Cross the real case -> inspector boundary. Only native executions and Office/
# raster helpers are stubbed. The real scripts decide preparation and forwarding.
foreach($name in @('CheckVisualCase.ps1','DocxMarkupReference.ps1','InspectPdf.ps1','JsonArray.ps1','CompareMath.ps1','PathHelpers.ps1',
    'PdfTextGeometry.ps1','ComparePdfTextLineStarts.ps1','ComparePdfTextOperations.ps1',
    'ComparePdfGraphicsOperations.ps1','ClassifyPdfChartGraphics.ps1','ClassifyPdfChartText.ps1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $repo 'tools') -Force
}
$vendor=Join-Path $repo 'tools/vendor/pdfium/win-x64'
New-Item -ItemType Directory -Path $vendor -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'vendor/pdfium/win-x64/pdfium.sha256') -Destination $vendor
@'
param($InputPath,$OutputDirectory,$Dpi,$CaseId,[switch]$SkipBuild,$Framework,[switch]$CacheOnly,[string]$CacheVariant,$RasterizerPath)
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'reference.pdf'),'%PDF-fake reference')
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'page-001.png'),'synthetic raster')
if ($CacheOnly) { [IO.File]::WriteAllText((Join-Path $OutputDirectory 'requested-variant.txt'),$CacheVariant) }
if ($RasterizerPath) { [IO.File]::WriteAllText((Join-Path $OutputDirectory 'rasterizer-path.txt'),$RasterizerPath) }
'@ | Set-Content (Join-Path $repo 'tools/RenderCachedReference.ps1')
@'
param($InputPdf,$OutputDirectory,$Dpi,[switch]$SkipBuild,$Framework,$RasterizerPath)
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'page-001.png'),'synthetic raster')
if ($RasterizerPath) { [IO.File]::WriteAllText((Join-Path $OutputDirectory 'rasterizer-path.txt'),$RasterizerPath) }
'@ | Set-Content (Join-Path $repo 'tools/RasterizePdf.ps1')
foreach($id in @('probe-a','probe-b')) {
    $directory=Join-Path $repo "visual-cases/cases/$id"
    [IO.File]::WriteAllText((Join-Path $directory 'input.docx'),'synthetic orchestration input')
    @{id=$id;kind='docx';input='input.docx';tags=@('approximate');docxMarkup='all-markup';docxMarkupGeometry='office';expected=@{
        maxTextLineStartDelta=0.15;maxTextLineStartYDelta=0.5;candidateFontsMustBeEmbedded=$true}} | ConvertTo-Json -Depth 4 |
        Set-Content (Join-Path $directory 'case.json')
}
$inspectionLog=Join-Path $repo 'inspection-calls.jsonl'
$candidateBaseline = 700
$candidateChartPageTwoDelta = 0
$emitGraphicsOperations = $true
$mutateVisualDiff = $false
$rasterPage = @{Page=1;DimensionsMatch=$true;MeanAbsoluteError=0;ChangedPixelRatioAtThreshold16=0;
    StructuralSimilarity=1;ForegroundColorHistogramCorrelation=1}
$rasterMetrics = @($rasterPage)
function dotnet {
    $values=@($args)
    if($values[0] -like '*Lokad.OoxPdf.Cli.dll') {
        [IO.File]::WriteAllText($values[3],"%PDF-fake candidate baseline $candidateBaseline")
        [IO.File]::WriteAllText($values[5],'[]')
    } elseif($values[0] -like '*Lokad.OoxPdf.VisualDiff.dll') {
        ConvertTo-Json -InputObject $rasterMetrics -Compress |
            Set-Content (Join-Path $values[3] 'metrics.json')
        if ($mutateVisualDiff) { [IO.File]::AppendAllText($values[0], 'changed') }
    } elseif($values[0] -like '*Lokad.OoxPdf.PdfInspect.dll') {
        $output=$values[2]
        $baseline=if($values[1] -like '*output.pdf') { $candidateBaseline } else { 700 }
        @(@{PageNumber=1;X=72;Y=$baseline;DecodedText='Visible content'},
            @{PageNumber=2;X=150;Y=150;DecodedText='42'}) | ConvertTo-Json -AsArray | Set-Content (Join-Path $output 'text-operations.json')
        if ($emitGraphicsOperations) {
            @(foreach ($page in 1, 2) {
                $maxX = if ($page -eq 2 -and $values[1] -like '*output.pdf') {300 + $candidateChartPageTwoDelta} else {300}
                @{PageNumber=$page;Kind='Stroke';Operator='S';SegmentCount=1;MoveCount=1;LineCount=1;CurveCount=0;CloseCount=0;LineWidth=1;MinX=100;MinY=100;MaxX=$maxX;MaxY=100}
                @{PageNumber=$page;Kind='Stroke';Operator='S';SegmentCount=1;MoveCount=1;LineCount=1;CurveCount=0;CloseCount=0;LineWidth=1;MinX=100;MinY=100;MaxX=100;MaxY=300}
            }) | ConvertTo-Json -AsArray | Set-Content (Join-Path $output 'graphics-operations.json')
        }
        '[{"PageNumber":1,"ResourceName":"F1","Embedded":true,"HasToUnicode":true,"IsSubset":true}]' | Set-Content (Join-Path $output 'font-resources.json')
        @{Inspector=$values[0];InspectorSha256=(Get-FileHash -LiteralPath $values[0]).Hash;Output=$output} | ConvertTo-Json -Compress |
            Add-Content $inspectionLog
    } else {throw 'Unexpected native execution request'}
    $global:LASTEXITCODE=0
}
$realCase=Join-Path $repo 'tools/CheckVisualCase.ps1'
$realManifest=Join-Path $repo 'visual-cases/cases/probe-a/case.json'
try {
    $before=@(Get-Content $buildLog).Count
    & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -CacheOnlyReference -RunId 'real-prebuilt' | Out-Null
    if(@(Get-Content $buildLog).Count -ne $before) {throw 'Real prebuilt case requested a build'}
    $priorRun = Join-Path $repo 'artifacts/visual/probe-a/real-prebuilt'
    if ([IO.File]::ReadAllText((Join-Path $priorRun 'reference/requested-variant.txt')) -cne 'docxMarkup=all;docxMarkupGeometry=word-compatible') {
        throw 'The public visual case used raw aliases for the cached reference key'
    }
    $priorPdf = Join-Path $priorRun 'candidate/output.pdf'
    $priorReport = Join-Path $priorRun 'run-info.json'
    $priorPdfHash = (Get-FileHash -LiteralPath $priorPdf).Hash
    $priorReportHash = (Get-FileHash -LiteralPath $priorReport).Hash
    $beforeInspection = @(Get-Content $inspectionLog).Count
    $candidateBaseline = 500
    $failure = $null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-prebuilt' | Out-Null } catch {$failure=$_.Exception.Message}
    if ($failure -notlike 'Visual run directory already exists.*' -or
        (Get-FileHash -LiteralPath $priorPdf).Hash -ne $priorPdfHash -or
        (Get-FileHash -LiteralPath $priorReport).Hash -ne $priorReportHash -or
        @(Get-Content $inspectionLog).Count -ne $beforeInspection -or @(Get-Content $buildLog).Count -ne $before) {
        throw 'The real case reopened an occupied run or modified its evidence'
    }
    $candidateBaseline = 700
    & $realCase -Case $realManifest -CliPath $frozenCli -RunId 'real-prepare-once' | Out-Null
    $added=@(Get-Content $buildLog | Select-Object -Skip $before | ConvertFrom-Json)
    if($added.Count -ne 2 -or @($added | Where-Object Description -eq 'PDF inspect').Count -ne 1) {throw 'Real direct case did not prepare inspector once'}
    $before=@(Get-Content $buildLog).Count
    & $script -Family probe -CliPath $frozenCli | Out-Null
    $added=@(Get-Content $buildLog | Select-Object -Skip $before | ConvertFrom-Json)
    if($added.Count -ne 3 -or @($added | Where-Object Description -eq 'PDF inspect').Count -ne 1) {throw 'Real family repeated per-case inspector builds'}
    $before=@(Get-Content $buildLog).Count
    & $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null
    if(@(Get-Content $buildLog).Count -ne $before) {throw 'Real prebuilt family requested a build'}
    $reportPath=Join-Path $repo 'artifacts/visual/reports/probe.json'
    $bound=Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
    if ($bound.schema -cne 'ooxpdf-visual-family-report/2' -or $bound.cases.Count -ne 2) {
        throw 'The real family did not publish a versioned comparison identity'
    }
    foreach ($row in $bound.cases) {
        $run=Join-Path $repo $row.runPath
        $caseDir=Join-Path $repo "visual-cases/cases/$($row.id)"
        $reference=Join-Path $run 'reference/reference.pdf'
        $info=Get-Content -LiteralPath (Join-Path $run 'run-info.json') -Raw | ConvertFrom-Json
        if ($row.inputSha256 -ne (Get-FileHash -LiteralPath (Join-Path $caseDir 'input.docx')).Hash -or
            $row.caseSha256 -ne (Get-FileHash -LiteralPath (Join-Path $caseDir 'case.json')).Hash -or
            $row.referencePdfSha256 -ne (Get-FileHash -LiteralPath $reference).Hash -or
            $row.dpi -ne 144 -or $row.docxMarkup -cne 'all-markup' -or $row.docxMarkupGeometry -cne 'office' -or
            $row.visualDiffSha256 -ne (Get-FileHash -LiteralPath $info.VisualDiffDll).Hash -or
            $row.rasterizerSha256 -ne (Get-FileHash -LiteralPath $info.RasterizerDll).Hash -or
            $row.pdfiumPinSha256 -ne (Get-FileHash -LiteralPath (Join-Path $vendor 'pdfium.sha256')).Hash) {
            throw 'The family row did not identify the actual comparison inputs'
        }
        if ($null -ne $info.GitCommit -or $null -ne $info.GitDirty) {
            throw 'The synthetic repository inherited an enclosing checkout identity'
        }
        if ($row.referencePageCount -ne 1 -or $row.candidatePageCount -ne 1 -or $row.pageCount -ne 1 -or
            $info.ReferencePageCount -ne 1 -or $info.CandidatePageCount -ne 1) {
            throw 'The case and family report did not retain separate raster page counts'
        }
    }
    & (Join-Path $PSScriptRoot 'CompareVisualReports.ps1') -Baseline $reportPath -Current $reportPath | Out-Null
    $failure=$null
    $candidateBaseline = 500
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-y-shift' | Out-Null } catch {$failure=$_.Exception.Message}
    if($failure -ne 'PDF text line-start gate failed.' -or @(Get-Content $buildLog).Count -ne $before) {throw 'The real case did not enforce its manifest Y bound without building'}
    $failedRun=Join-Path $repo 'artifacts/visual/probe-a/real-y-shift'
    $failedInfo=Get-Content -LiteralPath (Join-Path $failedRun 'run-info.json') -Raw | ConvertFrom-Json
    if ($failedInfo.ReferencePdfSha256 -ne (Get-FileHash -LiteralPath (Join-Path $failedRun 'reference/reference.pdf')).Hash) {
        throw 'The measured case lost its comparison identity on a gate failure'
    }
    $candidateBaseline = 700
    $combined = Get-Content -Raw -LiteralPath $realManifest | ConvertFrom-Json -AsHashtable
    $combined.expected.maxTextOperationPositionDelta = 0
    $combined.expected.maxGraphicsOperationBoundsDelta = 0
    $combined.expected.maxChartGraphicsStructureBoundsDelta = 0
    $combined.expected.maxChartTextStructurePositionDelta = 0
    foreach ($id in 'real-shared-inspection', 'real-filtered-chart') {
        if ($id -eq 'real-filtered-chart') { $combined.expected.compareChartTextStructurePageNumber = 2 }
        $combined | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $realManifest
        $beforeInspection = @(Get-Content $inspectionLog).Count
        & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId $id | Out-Null
        if (@(Get-Content $inspectionLog).Count -ne $beforeInspection + 2) {
            throw 'The combined text, graphics, chart and font gates must inspect each PDF once'
        }
        $layout = Join-Path $repo "artifacts/visual/probe-a/$id/comparison"
        foreach ($side in 'reference', 'candidate') {
            $allPages = @(Get-Content -Raw -LiteralPath (Join-Path $layout "pdf-graphics/$side/chart-structures.json") | ConvertFrom-Json)
            $plotPath = Join-Path $layout "pdf-graphics/$side/chart-text-plot-structures.json"
            if (@($allPages.PageNumber | Sort-Object -Unique).Count -ne 2) {
                throw 'The chart graphics report must retain both pages'
            }
            if ($id -eq 'real-shared-inspection') {
                if (Test-Path -LiteralPath $plotPath) { throw 'Matching chart classifiers must share one report' }
            } else {
                $plot = @(Get-Content -Raw -LiteralPath $plotPath | ConvertFrom-Json)
                if (!$plot.Count -or @($plot | Where-Object PageNumber -ne 2).Count) {
                    throw 'The chart text plot report must retain its own page selection'
                }
            }
        }
    }
    $chart = $combined.Clone()
    $chart.expected = @{additionalChartGraphicsStructureComparisons=@(@{pageNumber=2;maxBoundsDelta=0;kinds=@('HorizontalLine')})}
    $chart | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $realManifest
    $familyPath = Join-Path $repo 'visual-cases/families/probe.json'
    @{id='probe';kind='docx';casePatterns=@('probe-*')} | ConvertTo-Json | Set-Content -LiteralPath $familyPath
    $validator = Join-Path $repo 'tools/ValidateVisualCases.ps1'
    $caseRoot = Join-Path $repo 'artifacts/visual/probe-a'
    $beforeRuns = @(Get-ChildItem -LiteralPath $caseRoot -Directory).Count
    $beforeBuilds = @(Get-Content $buildLog).Count
    foreach ($entry in @(
        @{Script=$realCase;Arguments=@{Case=$realManifest;RunId='real-unbound-chart'}},
        @{Script=$script;Arguments=@{Family='probe'}},
        @{Script=$validator;Arguments=@{}}
    )) {
        $arguments = $entry.Arguments
        $failure = $null
        try { & $entry.Script @arguments | Out-Null } catch { $failure = $_.Exception.Message }
        if ($failure -cne "Visual case 'probe-a' requires maxChartGraphicsStructureBoundsDelta or maxChartGraphicsStructureBoundsDeltaByKind for additionalChartGraphicsStructureComparisons.") {
            throw "Missing chart bounds preflight failed unexpectedly: $failure"
        }
    }
    if (@(Get-ChildItem -LiteralPath $caseRoot -Directory).Count -ne $beforeRuns -or
        @(Get-Content $buildLog).Count -ne $beforeBuilds) { throw 'Incomplete chart gates started work before rejection' }
    $chart.expected.maxChartGraphicsStructureBoundsDelta = 0
    $chart | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $realManifest
    & $validator | Out-Null
    & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-chart-zero' | Out-Null
    $chart.expected.Remove('maxChartGraphicsStructureBoundsDelta')
    $chart.expected.maxChartGraphicsStructureBoundsDeltaByKind = @{HorizontalLine=2}
    $chart | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $realManifest
    $candidateChartPageTwoDelta = 1
    $failure = $null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-chart-page-two' | Out-Null }
    catch { $failure = $_.Exception.Message }
    finally { $candidateChartPageTwoDelta = 0 }
    if ($failure -cne 'Additional PDF chart graphics structure gate failed.') {
        throw "The page-two chart gate did not enforce its separate zero bound: $failure"
    }
    $fontOnly = $combined.Clone()
    $fontOnly.expected = @{candidateFontsMustBeEmbedded=$true}
    $fontOnly | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $realManifest
    $beforeInspection = @(Get-Content $inspectionLog).Count
    & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-font-only' | Out-Null
    $inspections = @(Get-Content $inspectionLog | Select-Object -Skip $beforeInspection | ConvertFrom-Json)
    if ($inspections.Count -ne 1 -or (Split-Path -Leaf $inspections[0].Output) -ne 'candidate') {
        throw 'A font-only case should inspect only its candidate'
    }
    $rasterOnly = $combined.Clone()
    $rasterOnly.expected = @{pageCountMustMatch=$true;dimensionsMustMatch=$true;maxMeanAbsoluteError=0;maxChangedPixelRatioAtThreshold16=0;
        compareChartGraphicsStructurePathGeometry=$false;additionalChartGraphicsStructureComparisons=@()}
    $rasterOnly | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $realManifest
    & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-zero-raster-errors' | Out-Null
    foreach ($control in @(
        @{Field='MeanAbsoluteError'; Value=$null; Error='Mean absolute error gate failed. Page 1 was unavailable, limit is 0.'},
        @{Field='ChangedPixelRatioAtThreshold16'; Value=$null; Error='Changed-pixel ratio gate failed. Page 1 was unavailable, limit is 0.'},
        @{Field='DimensionsMatch'; Value=$false; Error='Dimension mismatch on page(s): 1.'}
    )) {
        $original = $rasterPage[$control.Field]
        $rasterPage[$control.Field] = $control.Value
        $failure = $null
        try {
            & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId ('real-raster-' + $control.Field) | Out-Null
        }
        catch { $failure = $_.Exception.Message }
        finally { $rasterPage[$control.Field] = $original }
        if ($failure -cne $control.Error) { throw ('Incorrect raster gate outcome for ' + $control.Field + ': ' + $failure) }
    }
    $rasterMetrics = @()
    $failure = $null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-empty-raster' | Out-Null }
    catch { $failure = $_.Exception.Message }
    finally { $rasterMetrics = @($rasterPage) }
    if ($failure -cne 'The visual comparison produced no page measurements.') { throw 'An empty raster report must fail' }
    $combined | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $realManifest
    $emitGraphicsOperations = $false
    $failure = $null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-incomplete-inspection' | Out-Null } catch {$failure=$_.Exception.Message}
    if ($failure -ne 'PDF inspection did not produce graphics-operations.json for the reference PDF.') {
        throw 'An incomplete inspection must fail before the comparison gates'
    }
    $emitGraphicsOperations = $true
    Remove-Item -LiteralPath $inspector
    $failure=$null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-missing-tool' | Out-Null } catch {$failure=$_.Exception.Message}
    if($failure -notlike '*PDF inspect DLL is missing*' -or @(Get-Content $buildLog).Count -ne $before) {throw 'Real missing-tool case requested a build or ignored refusal'}

    $frozenTools = Join-Path $repo 'Frozen tools [pinned]'
    New-Item -ItemType Directory -Path $frozenTools | Out-Null
    $overrides = @{}
    foreach ($tool in @{
        RasterizerPath='PdfiumRasterizer'; VisualDiffPath='VisualDiff'; PdfInspectorPath='PdfInspect'
    }.GetEnumerator()) {
        $name='Lokad.OoxPdf.' + $tool.Value
        $path=Join-Path $frozenTools ($name + '.dll')
        [IO.File]::WriteAllText($path, 'frozen ' + $name)
        $overrides[$tool.Key]=$path
        $local=Join-Path $repo "tools/$name/bin/Debug/net11.0/$name.dll"
        if (Test-Path -LiteralPath $local) { [IO.File]::Delete($local) }
    }
    $boundHashes=@{}
    foreach ($key in $overrides.Keys) { $boundHashes[$key]=(Get-FileHash -LiteralPath $overrides[$key]).Hash }
    $beforeInspection=@(Get-Content $inspectionLog).Count
    & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -CacheOnlyReference -RunId 'real-frozen-tools' @overrides | Out-Null
    & $script -Family probe -CliPath $frozenCli -CacheOnlyReference @overrides | Out-Null
    $frozenReport=Get-Content -Raw -LiteralPath $reportPath | ConvertFrom-Json
    if ($frozenReport.passedCount -ne 2 -or @($frozenReport.cases | Where-Object {
        $_.visualDiffSha256 -ne $boundHashes.VisualDiffPath -or $_.rasterizerSha256 -ne $boundHashes.RasterizerPath -or
        $_.pdfInspectorSha256 -ne $boundHashes.PdfInspectorPath
    }).Count) { throw 'The frozen family did not bind its comparison tools for each case' }
    $roots=@('artifacts/visual/probe-a/real-frozen-tools') + @($frozenReport.cases.runPath)
    foreach ($root in $roots) {
        $run=Join-Path $repo $root
        $info=Get-Content -Raw -LiteralPath (Join-Path $run 'run-info.json') | ConvertFrom-Json
        if ($info.VisualDiffDll -cne $overrides.VisualDiffPath -or $info.RasterizerDll -cne $overrides.RasterizerPath -or
            $info.PdfInspectorDll -cne $overrides.PdfInspectorPath -or $info.VisualDiffSha256 -ne $boundHashes.VisualDiffPath -or
            $info.RasterizerSha256 -ne $boundHashes.RasterizerPath -or $info.PdfInspectorSha256 -ne $boundHashes.PdfInspectorPath) {
            throw 'The frozen case recorded a different tool binding'
        }
        foreach ($side in 'reference','candidate') {
            if ([IO.File]::ReadAllText((Join-Path $run "$side/rasterizer-path.txt")) -cne $overrides.RasterizerPath) {
                throw "The $side rasterizer did not use the frozen path"
            }
        }
    }
    $frozenInspections=@(Get-Content $inspectionLog | Select-Object -Skip $beforeInspection | ConvertFrom-Json)
    if ($frozenInspections.Count -ne 6 -or @($frozenInspections | Where-Object {
        $_.Inspector -cne $overrides.PdfInspectorPath -or $_.InspectorSha256 -ne $boundHashes.PdfInspectorPath
    }).Count -or @(Get-Content $buildLog).Count -ne $before) { throw 'Frozen tools built or used a different inspector' }

    $mutateVisualDiff=$true
    $failure=$null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-changed-diff' @overrides | Out-Null }
    catch { $failure=$_.Exception.Message }
    if ($failure -notlike 'Visual validation tool changed during the case:*') { throw 'A changed visual tool must fail the case' }
}
finally {Remove-Item Function:/dotnet}
$inspections=@(Get-Content (Join-Path $repo 'inspection-calls.jsonl') | ConvertFrom-Json)
if($inspections.Count -ne 30) {throw 'Real inspection coverage is incomplete'}
foreach($inspection in $inspections) {
    $record=Get-Content (Join-Path $inspection.Output 'pdf-inspect-build-info.json') -Raw | ConvertFrom-Json
    if($record.BuildRequested -ne $false -or $record.OutputSha256 -ne $inspection.InspectorSha256) {throw 'Inspector provenance or explicit prebuilt binding was lost'}
}
[ordered]@{Checks=@('List without runtime','Frozen CLI forwarding','Validation-only builds','Runtime identity',
    'Skip all builds','Default builds once','Prebuilt default CLI','All ten inspection capabilities',
    'No unused inspector','Missing prebuilt tool preflight','Real case/inspector forwarding',
    'Real once-only preparation','Real family reuse','Real missing-tool refusal','Occupied run refusal and evidence integrity',
    'Fresh shifted PDF fails the manifest Y bound','Combined gates inspect each PDF once',
    'Chart classification reuse and independent page scopes','Shared chart gate preflight',
    'Zero chart bounds admit equal page scopes','Secondary page-two bound rejects displaced chart ink','Candidate-only font inspection',
    'Incomplete inspection refusal','Inspector provenance','Canonical markup cache key',
    'Bound visual report identity','Separate raster page counts','Measured failure identity',
    'Zero upper limits admit exact metrics','Missing upper-bound metrics fail independently','Raster dimensions fail independently',
    'Empty raster report fails','Frozen comparison tools without local builds','Changed frozen tool refusal','Synthetic repository isolation');
    InspectionCapabilities=$inspectionKeys.Count;RealInspectionCalls=$inspections.Count;LibraryBuilds=0;LibraryExecutions=0;OfficeExports=0} |
    ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'evidence.json')
Write-Output 'Visual-family and real case/inspector binding checks pass; no actual build, library execution or Office export.'
