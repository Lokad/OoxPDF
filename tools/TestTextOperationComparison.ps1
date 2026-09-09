# Adversarial checks for ComparePdfTextOperations (T05): reordered runs must
# match robustly while split, duplicated, deleted, or moved runs fail loudly
# with reported (not invented) deltas. No Office/COM or reference cache needed.

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

function New-Op($X, $Y, $Text) {
    return [pscustomobject]@{ X = $X; Y = $Y; FontSize = 12; CharacterSpacing = 0; Payload = $Text }
}

$failures = @()
function Invoke-Case([string] $Name, [object[]] $Reference, [object[]] $Candidate, [int] $ExpectedExit, [int] $ExpectedMissing) {
    $refPath = Write-Ops ($Name + "-ref") $Reference
    $candPath = Write-Ops ($Name + "-cand") $Candidate
    $outPath = Join-Path $scratch ($Name + "-out.txt")
    $tool = Join-Path $repoRoot "tools/ComparePdfTextOperations.ps1"
    $proc = Start-Process pwsh -ArgumentList "-NoProfile", "-File", $tool, "-Reference", $refPath, "-Candidate", $candPath, "-MatchByPosition" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outPath
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

if ($failures.Count -ne 0) {
    throw ("Text comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Text comparison adversarial checks passed (6 cases)."
