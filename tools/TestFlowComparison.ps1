# Adversarial checks for CompareDocxLayoutPdfFlow (T05): reordered op files
# must compare identically while repeated, deleted, or split rows shift
# exactly one mapping verdict. No Office/COM or reference cache needed.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = Join-Path $repoRoot "artifacts/flow-compare-probe"
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $scratch | Out-Null

function New-FlowOp($X, $Y, $Text) {
    return [pscustomobject]@{ PageNumber = 1; X = $X; Y = $Y; FontSize = 12; CharacterSpacing = 0; DecodedText = $Text }
}

function New-LayoutLine($Y, $Length, $Block) {
    return [pscustomobject]@{ Kind = "TextLine"; Y = $Y; X = 10; SourceBlockIndex = $Block; SourceLineIndex = 0; TextLength = $Length }
}

$failures = @()
function Write-CaseDir([string] $Name, [object[]] $CandidateOps) {
    $dir = Join-Path $scratch $Name
    $candDir = Join-Path $dir "comparison/pdf-text/candidate"
    $refDir = Join-Path $dir "comparison/pdf-text/reference"
    New-Item -ItemType Directory -Force -Path $candDir, $refDir | Out-Null
    ($CandidateOps | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath (Join-Path $candDir "text-operations.json") -Encoding UTF8
    Copy-Item -LiteralPath (Join-Path $scratch "base-ref-ops.json") -Destination (Join-Path $refDir "text-operations.json")
    return $dir
}

$a1 = New-FlowOp 10 100 "A"
$a2 = New-FlowOp 20 100 "A"
$b1 = New-FlowOp 10 90 "B"
$b2 = New-FlowOp 20 90 "B"
$c1 = New-FlowOp 10 80 "C"
$c2 = New-FlowOp 20 80 "C"
$baseOps = @($a1, $a2, $b1, $b2, $c1, $c2)
($baseOps | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath (Join-Path $scratch "base-ref-ops.json") -Encoding UTF8
$layout = [pscustomobject]@{ Pages = @([pscustomobject]@{ Items = @(
  (New-LayoutLine 100 2 0),
  (New-LayoutLine 90 2 1),
  (New-LayoutLine 80 2 2)
) }) }
($layout | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $scratch "layout.json") -Encoding UTF8

function Invoke-FlowCase([string] $Name, [object[]] $CandidateOps) {
    $dir = Write-CaseDir $Name $CandidateOps
    $out = Join-Path $dir "summary.json"
    $tool = Join-Path $repoRoot "tools/CompareDocxLayoutPdfFlow.ps1"
    $proc = Start-Process pwsh -ArgumentList "-NoProfile", "-File", $tool, "-RunDirectory", $dir, "-LayoutSnapshot", (Join-Path $scratch "layout.json"), "-OutputJson", $out -NoNewWindow -Wait -PassThru -RedirectStandardOutput (Join-Path $dir "stdout.txt")
    if ($proc.ExitCode -ne 0) { throw ("Flow case " + $Name + " exited " + $proc.ExitCode) }
    return Get-Content -Raw -LiteralPath $out | ConvertFrom-Json
}

function Assert-Counts([string] $Name, $Summary, [int] $Matched, [int] $MissingRef, [int] $MissingCand, [int] $CandLines, [int] $MiddleBlock) {
    $mid = @($Summary.MissingReferenceMatches | Where-Object { [int]$_.SourceBlockIndex -eq $MiddleBlock })
    $midOk = if ($MissingRef -eq 0) { ($Summary.MissingReferenceMatches.Count -eq 0) } else { ($mid.Count -eq $MissingRef) }
    $ok = ($Summary.MatchedLineCount -eq $Matched) -and ($Summary.MissingReferenceMatchCount -eq $MissingRef) -and ($Summary.MissingCandidateRowCount -eq $MissingCand) -and ($Summary.CandidatePdfLineCount -eq $CandLines) -and ($Summary.AmbiguousReferenceMatchCount -eq 0) -and $midOk
    if ($ok) {
        Write-Host ("PASS: " + $Name)
    }
    else {
        $script:failures += $Name
        Write-Host ("FAIL: " + $Name)
    }
}

$base = Invoke-FlowCase "baseline" $baseOps
Assert-Counts "baseline" $base 3 0 0 3 1
$reorder = Invoke-FlowCase "reorder" @($c2, $c1, $b2, $b1, $a2, $a1)
$dupe = Invoke-FlowCase "dupe" @($a1, $a2, $b1, $b2, $b1, $b2, $c1, $c2)
$delete = Invoke-FlowCase "delete" @($a1, $a2, $c1, $c2)
$split = Invoke-FlowCase "split" @($a1, $a2, $b1, (New-FlowOp 20 92 "B"), $c1, $c2)
Assert-Counts "dupe" $dupe 2 1 0 3 1
Assert-Counts "delete" $delete 2 0 1 2 1
Assert-Counts "split" $split 2 1 0 4 1
$reorderDir = (Resolve-Path -LiteralPath (Join-Path $scratch "reorder")).Path.Replace("\", "\\")
$baselineDir = (Resolve-Path -LiteralPath (Join-Path $scratch "baseline")).Path.Replace("\", "\\")
$rb = (Get-Content -Raw -LiteralPath (Join-Path $scratch "reorder/summary.json")).Replace($reorderDir, "DIR")
$bb = (Get-Content -Raw -LiteralPath (Join-Path $scratch "baseline/summary.json")).Replace($baselineDir, "DIR")
if ($rb -cne $bb) { $script:failures += "reorder"; Write-Host "FAIL: reorder (summary differs from baseline)" }
else { Write-Host "PASS: reorder (summary identical to baseline)" }
if ($failures.Count -ne 0) {
    throw ("Flow comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Flow comparison adversarial checks passed (5 cases)."
