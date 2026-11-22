# Exercise report comparison semantics with small reports; never builds or renders.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/visual-report-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$baseline=@{id='probe';family='synthetic';kind='pptx';passed=$true;pageCount=1;
    minStructuralSimilarity=1.0;minForegroundColorHistogramCorrelation=1.0;
    minForegroundRecall=1.0;requiredForegroundRecall=.90;
    maxMeanAbsoluteError=0.0;maxChangedPixelRatioAtThreshold16=0.0}
$cases=[Collections.Generic.List[object]]::new()
$cases.Add(@{Id='numeric-zero';Before=$baseline;After=$baseline;Accept=$true})
$strings=$baseline.Clone();$strings.maxMeanAbsoluteError='0';$strings.maxChangedPixelRatioAtThreshold16='0'
$cases.Add(@{Id='string-zero';Before=$strings;After=$strings;Accept=$true})
$drop=$baseline.Clone();$drop.minForegroundRecall=.5
$cases.Add(@{Id='finite-recall-drop';Before=$baseline;After=$drop;Accept=$false})
$zero=$baseline.Clone();$zero.minForegroundRecall=0.0
$cases.Add(@{Id='zero-recall-drop';Before=$baseline;After=$zero;Accept=$false})
foreach ($value in @($null,'not-measured','NaN','Infinity','-Infinity')) {
    $after=$baseline.Clone();$after.minForegroundRecall=$value
    $id=if ($null -eq $value) {'null'} else {$value}
    $cases.Add(@{Id=('required-recall-'+$id);Before=$baseline;After=$after;Accept=$false})
}
$absent=$baseline.Clone();$absent.Remove('minForegroundRecall')
$cases.Add(@{Id='required-recall-absent';Before=$baseline;After=$absent;Accept=$false})
$lost=$absent.Clone();$lost.Remove('requiredForegroundRecall')
$cases.Add(@{Id='known-requirement-lost';Before=$baseline;After=$lost;Accept=$false})
$legacy=$lost.Clone()
$cases.Add(@{Id='legacy-unmeasured';Before=$legacy;After=$legacy;Accept=$true})
$inapplicable=$legacy.Clone();$inapplicable.minForegroundRecall=$null
$cases.Add(@{Id='inapplicable-unrequired';Before=$inapplicable;After=$inapplicable;Accept=$true})
foreach ($value in @($null,'not-measured','NaN','Infinity')) {
    $after=$baseline.Clone();$after.maxMeanAbsoluteError=$value
    $id=if ($null -eq $value) {'null'} else {$value}
    $cases.Add(@{Id=('invalid-mae-'+$id);Before=$baseline;After=$after;Accept=$false})
}
$bound=$baseline.Clone()
$bound.kind='docx'
$bound.docxMarkup='all'
$bound.docxMarkupGeometry='word-compatible'
$bound.referencePageCount=1
$bound.candidatePageCount=1
$bound.dpi=144
$bound.caseSha256='A'*64
$bound.inputSha256='B'*64
$bound.referencePdfSha256='C'*64
$bound.visualDiffSha256='D'*64
$bound.rasterizerSha256='E'*64
$bound.pdfiumPinSha256='F'*64
$bound.requiredStructuralSimilarity=.995
$bound.requiredForegroundColorHistogramCorrelation=.98
$cases.Add(@{Id='bound-different-runtime';Before=$bound;After=$bound;Bound=$true;Accept=$true})
$changes=@{
    caseSha256=('1'*64); inputSha256=('2'*64); referencePdfSha256=('3'*64)
    visualDiffSha256=('4'*64); rasterizerSha256=('5'*64); pdfiumPinSha256=('6'*64)
    dpi=96; docxMarkup='final'; docxMarkupGeometry='preserve'
    requiredStructuralSimilarity=.990; requiredForegroundColorHistogramCorrelation=.97; requiredForegroundRecall=.89
}
foreach ($field in $changes.Keys | Sort-Object) {
    $after=$bound.Clone()
    $after[$field]=$changes[$field]
    $cases.Add(@{Id="different-$field";Before=$bound;After=$after;Bound=$true;Accept=$false;ErrorLike="*differs from current $field*"})
}
$unbound=$bound.Clone()
$unbound.Remove('referencePdfSha256')
$cases.Add(@{Id='missing-bound-reference';Before=$bound;After=$unbound;Bound=$true;Accept=$false;ErrorLike='*missing or invalid referencePdfSha256*'})
$cases.Add(@{Id='mixed-schemas';Before=$bound;After=$bound;Bound=$true;LegacyBaseline=$true;Accept=$false;ErrorLike='*both must record the comparison inputs*'})
$cases.Add(@{Id='old-bound-schema';Before=$bound;After=$bound;Bound=$true;OldSchema=$true;Accept=$false;ErrorLike='*lacks separate candidate and reference page counts*'})
foreach ($count in @(1.25, $true, 0, 2147483648)) {
    $after=$bound.Clone()
    $after.pageCount=$count
    $cases.Add(@{Id="non-page-count-$count";Before=$bound;After=$after;Bound=$true;Accept=$false;ErrorLike='*incomplete measurement*'})
}
$added=$bound.Clone()
$added.id='new-case'
$cases.Add(@{Id='complete-addition';Before=$bound;After=@($bound,$added);Bound=$true;Accept=$true})
$incomplete=$added.Clone()
$incomplete.Remove('minStructuralSimilarity')
$cases.Add(@{Id='incomplete-addition';Before=$bound;After=@($bound,$incomplete);Bound=$true;Accept=$false;ErrorLike='*incomplete measurement*'})
$twoPages=$bound.Clone()
$twoPages.pageCount=2
$twoPages.candidatePageCount=2
$cases.Add(@{Id='different-page-count';Before=$bound;After=$twoPages;Bound=$true;Accept=$false;ErrorLike='*visual report regression*'})
$failed=$bound.Clone();$failed.passed=$false;$failed.minStructuralSimilarity=.90
$cases.Add(@{Id='complete-inherited-failure';Before=$failed;After=$failed;Bound=$true;Accept=$true})
$twoPageFailure=$failed.Clone();$twoPageFailure.pageCount=2;$twoPageFailure.referencePageCount=2;$twoPageFailure.candidatePageCount=2
$missingPage=$twoPageFailure.Clone();$missingPage.candidatePageCount=1
$extraPage=$failed.Clone();$extraPage.pageCount=2;$extraPage.candidatePageCount=2
$cases.Add(@{Id='inherited-failure-missing-page';Before=$twoPageFailure;After=$missingPage;Bound=$true;Accept=$false;ErrorLike='*visual report regression*'})
$cases.Add(@{Id='inherited-failure-extra-page';Before=$failed;After=$extraPage;Bound=$true;Accept=$false;ErrorLike='*visual report regression*'})
$missingMetrics=$failed.Clone();$missingMetrics.minStructuralSimilarity=$null;$missingMetrics.maxMeanAbsoluteError=$null
$cases.Add(@{Id='inherited-failure-missing-metrics';Before=$failed;After=$missingMetrics;Bound=$true;Accept=$false;ErrorLike='*incomplete measurement*'})
$cases.Add(@{Id='inherited-failure-incomplete-baseline';Before=$missingMetrics;After=$failed;Bound=$true;Accept=$false;ErrorLike='*incomplete measurement*'})
$missingCount=$failed.Clone();$missingCount.Remove('candidatePageCount')
$cases.Add(@{Id='inherited-failure-missing-count';Before=$failed;After=$missingCount;Bound=$true;Accept=$false;ErrorLike='*incomplete measurement*'})
$missingRows=$twoPageFailure.Clone();$missingRows.pageCount=1
$cases.Add(@{Id='inherited-failure-missing-rows';Before=$twoPageFailure;After=$missingRows;Bound=$true;Accept=$false;ErrorLike='*incomplete measurement*'})
$cases.Add(@{Id='inconsistent-pinned-reference-count';Before=$failed;After=$missingPage;Bound=$true;Accept=$false;ErrorLike='*differs from current referencePageCount*'})
$metricRegressions=@{minStructuralSimilarity=.99;minForegroundColorHistogramCorrelation=.97;
    maxMeanAbsoluteError=.03;maxChangedPixelRatioAtThreshold16=.01}
foreach ($field in $metricRegressions.Keys | Sort-Object) {
    $after=$bound.Clone()
    $after[$field]=$metricRegressions[$field]
    $cases.Add(@{Id="worse-$field";Before=$bound;After=$after;Bound=$true;Accept=$false;ErrorLike='*visual report regression*'})
}

$results=@()
foreach ($case in $cases) {
    $folder=Join-Path $root $case.Id
    New-Item -ItemType Directory -Path $folder | Out-Null
    foreach ($side in @('Before','After')) {
        $report=@{family='synthetic';cases=@($case[$side])}
        if ($case.Bound -and ($side -eq 'After' -or !$case.LegacyBaseline)) {
            $report.schema=if ($case.OldSchema) {'ooxpdf-visual-family-report/1'} else {'ooxpdf-visual-family-report/2'}
            $report.librarySha256=if ($side -eq 'Before') {'7'*64} else {'8'*64}
        }
        $report | ConvertTo-Json -Depth 6 |
            Set-Content -LiteralPath (Join-Path $folder ($side+'.json')) -Encoding utf8
    }
    $log=Join-Path $folder 'comparison.log'
    $errorText=$null
    try {
        & (Join-Path $PSScriptRoot 'CompareVisualReports.ps1') -Baseline (Join-Path $folder 'Before.json') -Current (Join-Path $folder 'After.json') *> $log
    }
    catch { $errorText=$_.Exception.Message }
    $accepted=$null -eq $errorText
    if ($accepted -ne $case.Accept) { throw ($case.Id+': unexpected comparison result; '+$errorText) }
    if ($case.ErrorLike -and $errorText -notlike $case.ErrorLike) {
        throw ($case.Id+': unexpected refusal; '+$errorText)
    }
    if ($accepted -and ((Get-Content -LiteralPath $log -Raw) -match 'Historical visual comparison') -ne !$case.Bound) {
        throw ($case.Id+': historical comparison status differs')
    }
    $results+=@{Id=$case.Id;Accepted=$accepted;Expected=$case.Accept;Error=$errorText}
}
@{Checks=$results;Builds=0;Conversions=0;OfficeExports=0;ThresholdsChanged=$false} | ConvertTo-Json -Depth 6 |
    Set-Content -LiteralPath (Join-Path $root 'evidence.json') -Encoding utf8
"Visual report comparison: $($results.Count) checks pass; no build or render."
