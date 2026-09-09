# Adversarial checks for ComparePdfGraphicsOperations (T05): reordered runs
# must match robustly while duplicated, deleted, kind-mismatched, or moved
# runs fail loudly with reported (not invented) deltas. No Office/COM or
# reference cache needed.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = Join-Path $repoRoot "artifacts/graphics-compare-probe"
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $scratch | Out-Null

function Write-Ops([string] $Name, [object[]] $Ops) {
    $path = Join-Path $scratch ($Name + ".json")
    ($Ops | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $path -Encoding UTF8
    return $path
}

function New-Rect($Kind, $MinX, $MinY, $MaxX, $MaxY) {
    return [pscustomobject]@{ PageNumber = 1; Kind = $Kind; Operator = "f*"; SegmentCount = 4; MoveCount = 1; LineCount = 3; CurveCount = 0; CloseCount = 1; MinX = $MinX; MinY = $MinY; MaxX = $MaxX; MaxY = $MaxY; LineWidth = 1 }
}

$failures = @()
function Invoke-Case([string] $Name, [object[]] $Reference, [object[]] $Candidate, [int] $ExpectedExit, [int] $ExpectedMissing) {
    $refPath = Write-Ops ($Name + "-ref") $Reference
    $candPath = Write-Ops ($Name + "-cand") $Candidate
    $outPath = Join-Path $scratch ($Name + "-out.txt")
    $tool = Join-Path $repoRoot "tools/ComparePdfGraphicsOperations.ps1"
    $proc = Start-Process pwsh -ArgumentList "-NoProfile", "-File", $tool, "-Reference", $refPath, "-Candidate", $candPath, "-MatchByBounds" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outPath
    $code = $proc.ExitCode
    $missing = @(Get-Content -LiteralPath $outPath | Where-Object { $_ -match "^\s*\d+\s+missing\b" }).Count
    if ($code -ne $ExpectedExit -or $missing -ne $ExpectedMissing) {
        $script:failures += ("{0}: exit={1} (want {2}), missing={3} (want {4})" -f $Name, $code, $ExpectedExit, $missing, $ExpectedMissing)
        Write-Host ("FAIL: {0}" -f $script:failures[-1])
    }
    else {
        Write-Host ("PASS: {0} (exit={1}, missing={2})" -f $Name, $code, $missing)
    }
}

$a = New-Rect "Fill" 10 10 30 30
$b = New-Rect "Fill" 50 50 70 70
$c = New-Rect "Fill" 90 90 110 110
Invoke-Case "identical" @($a, $b, $c) @($a, $b, $c) 0 0
Invoke-Case "reordered" @($a, $b, $c) @($c, $b, $a) 0 0
Invoke-Case "deleted" @($a, $b, $c) @($a, $c) 1 1
Invoke-Case "duplicated" @($a, $b) @($a, $b, (New-Rect "Fill" 50 50 70 70)) 1 1
Invoke-Case "kind-mismatch" @($a) @((New-Rect "Stroke" 10 10 30 30)) 1 0
Invoke-Case "moved" @($a) @((New-Rect "Fill" 20 10 40 30)) 1 0

if ($failures.Count -ne 0) {
    throw ("Graphics comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Graphics comparison adversarial checks passed (6 cases)."
