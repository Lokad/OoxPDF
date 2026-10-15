# Adversarial checks for ComparePdfTextOperations (T05): reordered runs must
# match robustly while split, duplicated, deleted, or moved runs fail loudly
# with reported (not invented) deltas. Repeated identical rows at distinct
# positions must neither dedup-match nor cascade: deleting/duplicating one
# reports exactly one missing, and a near-neighbor substitution reports its
# real delta. No Office/COM or reference cache needed.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = Join-Path $repoRoot "artifacts/text-compare-probe"
$scratchFullPath = [IO.Path]::GetFullPath($scratch)
if (!$scratchFullPath.StartsWith((Resolve-Path -LiteralPath $repoRoot).Path + [IO.Path]::DirectorySeparatorChar + 'artifacts' + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Unexpected comparison-probe cleanup target'
}
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $scratch | Out-Null

function Write-Ops([string] $Name, [object[]] $Ops) {
    $path = Join-Path $scratch ($Name + ".json")
    ($Ops | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $path -Encoding UTF8
    return $path
}

function New-Op($X, $Y, $Text, $EffX = $null, $EffY = $null) {
    return [pscustomobject]@{ X = $X; Y = $Y; FontSize = 12; CharacterSpacing = 0; Payload = $Text; EffectiveX = $EffX; EffectiveY = $EffY }
}

$failures = @()
function Invoke-Case([string] $Name, [object[]] $Reference, [object[]] $Candidate, [int] $ExpectedExit, [int] $ExpectedMissing, [string] $ExtraArgs = "", [switch] $ByIndex) {
    $refPath = Write-Ops ($Name + "-ref") $Reference
    $candPath = Write-Ops ($Name + "-cand") $Candidate
    $outPath = Join-Path $scratch ($Name + "-out.txt")
    $reportPath = Join-Path $scratch ($Name + '-report.json')
    $tool = Join-Path $repoRoot "tools/ComparePdfTextOperations.ps1"
    $toolArgs = @("-NoProfile", "-File", $tool, "-Reference", $refPath, "-Candidate", $candPath)
    if (!$ByIndex) { $toolArgs += '-MatchByPosition' }
    $toolArgs += @($ExtraArgs -split " " | Where-Object { $_ -ne "" })
    $toolArgs += @('-OutputPath', $reportPath)
    $proc = Start-Process pwsh -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outPath
    $code = $proc.ExitCode
    $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
    if ($report.ReferenceSha256 -ne (Get-FileHash -LiteralPath $refPath).Hash -or
        $report.CandidateSha256 -ne (Get-FileHash -LiteralPath $candPath).Hash -or
        @($report.Rows | Where-Object Status -ne 'ok').Count -ne $report.FailureCount -or
        ($report.FailureCount -eq 0) -ne ($code -eq 0)) {
        throw ('Structured comparison report differs: ' + $Name)
    }
    $missing = @(Get-Content -LiteralPath $outPath | Where-Object { $_ -match "^\s*\d+\s+missing\b" }).Count
    if ($code -ne $ExpectedExit -or $missing -ne $ExpectedMissing) {
        $script:failures += ("{0}: exit={1} (want {2}), missing={3} (want {4})" -f $Name, $code, $ExpectedExit, $missing, $ExpectedMissing)
        Write-Host ("FAIL: {0}" -f $script:failures[-1])
    }
    else {
        Write-Host ("PASS: {0} (exit={1}, missing={2})" -f $Name, $code, $missing)
    }
}
$a = New-Op 10 20 "Alpha"
$b = New-Op 30 20 "Beta"

$c = New-Op 50 20 "Gamma"
Invoke-Case "identical" @($a, $b, $c) @($a, $b, $c) 0 0
Invoke-Case "reordered" @($a, $b, $c) @($c, $b, $a) 0 0
Invoke-Case "deleted" @($a, $b, $c) @($a, $c) 1 1
Invoke-Case "duplicated" @($a, $b) @($a, $b, (New-Op 30 20 "Beta")) 1 1
Invoke-Case "split" @((New-Op 10 20 "Hello")) @((New-Op 10 20 "Hel"), (New-Op 25 20 "lo")) 1 1
Invoke-Case "moved" @($a) @((New-Op 10.5 20 "Alpha")) 1 0
$r1 = New-Op 10 100 "Row"
$r2 = New-Op 10 200 "Row"
Invoke-Case "repeat-identical" @($r1, $r2) @($r1, $r2) 0 0
Invoke-Case "repeat-deleted" @($r1, $r2) @($r1) 1 1
Invoke-Case "repeat-duplicated" @($r1, $r2) @($r1, $r2, (New-Op 10 200 "Row")) 1 1
$mixA1 = New-Op 10 100 "A"
$mixB = New-Op 10 200 "B"
$mixA2 = New-Op 10 300 "A"
Invoke-Case "repeat-interleaved-deleted" @($mixA1, $mixB, $mixA2) @($mixA1, $mixA2) 1 1
Invoke-Case "repeat-neighbor-substituted" @($mixA1, $mixB) @($mixA1, (New-Op 10 205 "C")) 1 0
$mEff = New-Op 10 20 "Shifted" 110 120
$mCand = New-Op 110 120 "Shifted" 110 120
Invoke-Case "matrix-raw" @($mEff) @($mCand) 1 0
Invoke-Case "matrix-effective" @($mEff) @($mCand) 0 0 "-UseEffectiveMatrix"
Invoke-Case "matrix-fallback" @((New-Op 10 20 "Plain")) @((New-Op 10 20 "Plain")) 0 0 "-UseEffectiveMatrix"

function New-MetricOp($X, $Y, $Text, $Advance) {
    [pscustomobject]@{ PageNumber=1; X=$X; Y=$Y; FontSize=12; CharacterSpacing=0; Payload=$Text; DecodedText=$Text; EffectiveX=$X; EffectiveY=$Y; EmittedAdvancePoints=$Advance; NaturalWidthPoints=$Advance; AdjustmentTotalPoints=0; NetSpacingGapTotalPoints=0; CharacterSpacingGapTotalPoints=0; DecodedRuneCount=([string]$Text).Length; TextChunkCount=1 }
}
$hello = New-MetricOp 10 100 'Hello' 25
$space = New-MetricOp 35 100 ' ' 3
$world = New-MetricOp 38 100 'world' 25
$blank = New-MetricOp 10 80 ' ' 3
$whole = New-MetricOp 10 100 'Hello world' 53
$pageOne = New-MetricOp 10 100 'Page-owned content' 80
$pageTwo = $pageOne.PSObject.Copy()
$pageTwo.PageNumber = 2
Invoke-Case 'page-moved-nearest' @($pageOne) @($pageTwo) 1 2 '-CompareDecodedText'
Invoke-Case 'page-moved-indexed' @($pageOne) @($pageTwo) 1 0 '-CompareDecodedText' -ByIndex
Invoke-Case 'page-reordered' @($pageOne, $pageTwo) @($pageTwo, $pageOne) 0 0 '-CompareDecodedText'
$visibleArgs = '-MergeSameLineOperations -IgnoreWhitespaceOnlyLines -CompareDecodedText -UseEffectiveMatrix'
function New-ScaledMetricOp($X, $Text, $FontSize, $Advance, $Scale, $Spacing = 0d) {
    $op = New-MetricOp $X 100 $Text $Advance
    $op.FontSize = $FontSize
    $op.CharacterSpacing = $Spacing
    foreach ($pair in @(@('EffectiveA',$Scale), @('EffectiveB',0d), @('EffectiveC',0d), @('EffectiveD',$Scale))) {
        $op | Add-Member -NotePropertyName $pair[0] -NotePropertyValue $pair[1]
    }
    return $op
}
Invoke-Case 'effective-font-and-spacing' @((New-ScaledMetricOp 10 'Alpha' 6 5 1 .2)) @((New-ScaledMetricOp 10 'Alpha' 12 10 .5 .4)) 0 0 '-UseEffectiveMatrix -CompareDecodedText'
Invoke-Case 'effective-scaled-gap' @((New-ScaledMetricOp 10 'A' 6 5 1), (New-ScaledMetricOp 20 'B' 6 5 1)) @((New-ScaledMetricOp 10 'A' 12 10 .5), (New-ScaledMetricOp 20 'B' 12 10 .5)) 0 0 $visibleArgs
Invoke-Case 'effective-expanded-adjacent' @((New-ScaledMetricOp 10 'A' 12 10 1), (New-ScaledMetricOp 20 'B' 12 10 1)) @((New-ScaledMetricOp 10 'A' 6 5 2), (New-ScaledMetricOp 20 'B' 6 5 2)) 0 0 $visibleArgs
Invoke-Case 'decoded-neighbor-columns' @((New-MetricOp 10 100 'Body' 20), (New-MetricOp 400 99 'Comment' 40)) @((New-MetricOp 400 101 'Comment' 40), (New-MetricOp 10 102 'Body' 20)) 0 0 '-CompareDecodedText -PositionTolerance 3'
Invoke-Case 'decoded-neighbor-change' @((New-MetricOp 10 100 'Body' 20), (New-MetricOp 400 99 'Comment' 40)) @((New-MetricOp 400 101 'Changed' 40), (New-MetricOp 10 102 'Body' 20)) 1 0 '-CompareDecodedText -PositionTolerance 3'
Invoke-Case 'decoded-case-change' @((New-MetricOp 10 100 'Alpha' 20)) @((New-MetricOp 10 100 'alpha' 20)) 1 0 '-CompareDecodedText'
Invoke-Case 'rounded-baseline-boundary' @((New-MetricOp 38 100.02 'world' 25), (New-MetricOp 35 100.04 ' ' 3), (New-MetricOp 10 100.06 'Hello' 25)) @((New-MetricOp 10 100.04 'Hello world' 53)) 0 0 $visibleArgs
Invoke-Case 'separate-baselines' @($hello, (New-MetricOp 38 100.3 'world' 25)) @($whole) 1 1 $visibleArgs
Invoke-Case 'blank-line-default' @($whole, $blank) @($whole) 1 1 '-MergeSameLineOperations -CompareDecodedText'
Invoke-Case 'blank-line-geometry' @($hello, $space, $world, $blank) @($whole) 0 0 $visibleArgs
Invoke-Case 'blank-line-moved' @($whole, $blank) @((New-MetricOp 10 99 'Hello world' 53)) 1 0 $visibleArgs
Invoke-Case 'blank-line-missing-text' @($whole, $blank) @($blank) 1 1 $visibleArgs
$unknown = New-MetricOp 10 80 $null 3
$unknown.Payload = 'undecoded'
Invoke-Case 'blank-line-unknown' @($unknown) @($blank) 1 1 $visibleArgs
Invoke-Case 'blank-line-partly-unknown' @($unknown, (New-MetricOp 13 80 ' ' 3)) @($blank) 1 1 $visibleArgs

if ($failures.Count -ne 0) {
    throw ("Text comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Text comparison adversarial checks passed (28 cases)."
