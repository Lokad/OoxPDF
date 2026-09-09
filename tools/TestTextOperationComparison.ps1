# Adversarial checks for ComparePdfTextOperations (T05): reordered runs must
# match robustly while split, duplicated, deleted, or moved runs fail loudly
# with reported (not invented) deltas. Repeated identical rows at distinct
# positions must neither dedup-match nor cascade: deleting/duplicating one
# reports exactly one missing, and a near-neighbor substitution reports its
# real delta. No Office/COM or reference cache needed.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = Join-Path $repoRoot "artifacts/text-compare-probe"
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
function Invoke-Case([string] $Name, [object[]] $Reference, [object[]] $Candidate, [int] $ExpectedExit, [int] $ExpectedMissing, [string] $ExtraArgs = "") {
    $refPath = Write-Ops ($Name + "-ref") $Reference
    $candPath = Write-Ops ($Name + "-cand") $Candidate
    $outPath = Join-Path $scratch ($Name + "-out.txt")
    $tool = Join-Path $repoRoot "tools/ComparePdfTextOperations.ps1"
    $toolArgs = @("-NoProfile", "-File", $tool, "-Reference", $refPath, "-Candidate", $candPath, "-MatchByPosition") + @($ExtraArgs -split " " | Where-Object { $_ -ne "" })
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

if ($failures.Count -ne 0) {
    throw ("Text comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Text comparison adversarial checks passed (14 cases)."
