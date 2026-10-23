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
$results=@()
foreach ($case in $cases) {
    $folder=Join-Path $root $case.Id
    New-Item -ItemType Directory -Path $folder | Out-Null
    foreach ($side in @('Before','After')) {
        @{family='synthetic';cases=@($case[$side])} | ConvertTo-Json -Depth 6 |
            Set-Content -LiteralPath (Join-Path $folder ($side+'.json')) -Encoding utf8
    }
    $errorText=$null
    try {
        & (Join-Path $PSScriptRoot 'CompareVisualReports.ps1') -Baseline (Join-Path $folder 'Before.json') -Current (Join-Path $folder 'After.json') *> (Join-Path $folder 'comparison.log')
    }
    catch { $errorText=$_.Exception.Message }
    $accepted=$null -eq $errorText
    if ($accepted -ne $case.Accept) { throw ($case.Id+': unexpected comparison result; '+$errorText) }
    $results+=@{Id=$case.Id;Accepted=$accepted;Expected=$case.Accept;Error=$errorText}
}
@{Checks=$results;Builds=0;Conversions=0;OfficeExports=0;ThresholdsChanged=$false} | ConvertTo-Json -Depth 6 |
    Set-Content -LiteralPath (Join-Path $root 'evidence.json') -Encoding utf8
"Visual report comparison: $($results.Count) checks pass; no build or render."
