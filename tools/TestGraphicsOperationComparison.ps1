# Adversarial checks for ComparePdfGraphicsOperations (T05): reordered runs
# must match robustly while duplicated, deleted, kind-mismatched, or moved
# runs fail loudly with reported (not invented) deltas. Translated and scaled
# equivalents (Fill and Clip, with and without path-coordinate matching) must
# likewise fail loudly: the comparison has no transform normalization, so a
# shifted path reports its real delta instead of passing or crashing.
# No Office/COM or reference cache needed.

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

function New-Rect($Kind, $MinX, $MinY, $MaxX, $MaxY, $Operator = "f*") {
    return [pscustomobject]@{ PageNumber = 1; Kind = $Kind; Operator = $Operator; SegmentCount = 4; MoveCount = 1; LineCount = 3; CurveCount = 0; CloseCount = 1; MinX = $MinX; MinY = $MinY; MaxX = $MaxX; MaxY = $MaxY; LineWidth = 1 }
}

function New-PathRect($Kind, $MinX, $MinY, $MaxX, $MaxY, $Dx = 0, $Dy = 0, $Operator = "f*") {
    $x0 = $MinX + $Dx
    $y0 = $MinY + $Dy
    $x1 = $MaxX + $Dx
    $y1 = $MaxY + $Dy
    $commands = @(
        [pscustomobject]@{ Operator = "m"; Values = @($x0, $y0) },
        [pscustomobject]@{ Operator = "l"; Values = @($x1, $y0) },
        [pscustomobject]@{ Operator = "l"; Values = @($x1, $y1) },
        [pscustomobject]@{ Operator = "l"; Values = @($x0, $y1) },
        [pscustomobject]@{ Operator = "h"; Values = @() }
    )
    return [pscustomobject]@{ PageNumber = 1; Kind = $Kind; Operator = $Operator; SegmentCount = 4; MoveCount = 1; LineCount = 3; CurveCount = 0; CloseCount = 1; MinX = $x0; MinY = $y0; MaxX = $x1; MaxY = $y1; LineWidth = 1; PathCommands = $commands }
}

$failures = @()
function Invoke-Case([string] $Name, [object[]] $Reference, [object[]] $Candidate, [int] $ExpectedExit, [int] $ExpectedMissing, [string] $ExtraArgs = "") {
    $refPath = Write-Ops ($Name + "-ref") $Reference
    $candPath = Write-Ops ($Name + "-cand") $Candidate
    $outPath = Join-Path $scratch ($Name + "-out.txt")
    $tool = Join-Path $repoRoot "tools/ComparePdfGraphicsOperations.ps1"
    $toolArgs = @("-NoProfile", "-File", $tool, "-Reference", $refPath, "-Candidate", $candPath, "-MatchByBounds") + @($ExtraArgs -split " " | Where-Object { $_ -ne "" })
    $proc = Start-Process pwsh -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outPath
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
$k1 = New-Rect "Clip" 10 10 30 30
$k2 = New-Rect "Clip" 50 50 70 70
Invoke-Case "clip-identical" @($k1, $k2) @($k1, $k2) 0 0
Invoke-Case "clip-reordered" @($k1, $k2) @($k2, $k1) 0 0
Invoke-Case "clip-deleted" @($k1, $k2) @($k1) 1 1
Invoke-Case "operator-ignored" @($a) @((New-Rect "Fill" 10 10 30 30 "f")) 0 0
Invoke-Case "operator-mismatch" @($a) @((New-Rect "Fill" 10 10 30 30 "f")) 1 0 "-MatchOperator"
$pa = New-PathRect "Fill" 10 10 30 30
Invoke-Case "path-identical" @($pa) @((New-PathRect "Fill" 10 10 30 30)) 0 0
Invoke-Case "path-translated" @($pa) @((New-PathRect "Fill" 10 10 30 30 10 0)) 1 0
Invoke-Case "path-translated-coords" @($pa) @((New-PathRect "Fill" 10 10 30 30 10 0)) 1 0 "-MatchPathCommandCoordinates"
Invoke-Case "path-scaled" @($pa) @((New-PathRect "Fill" 10 10 50 50)) 1 0
$pc = New-PathRect "Clip" 10 10 30 30 0 0 "W*"
Invoke-Case "clip-translated" @($pc) @((New-PathRect "Clip" 10 10 30 30 10 0 "W*")) 1 0

if ($failures.Count -ne 0) {
    throw ("Graphics comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Graphics comparison adversarial checks passed (16 cases)."
