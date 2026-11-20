# Adversarial checks for ComparePptxTextEmission (T05): reordered emission
# runs must match robustly, extra/missing runs must report exactly one missing,
# the text-then-position filter must pair by text, and font-size drift must
# fail as a delta. No Office/COM or reference cache needed.

[CmdletBinding()]
param([string] $OutputDirectory)

$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
$repoRoot = Split-Path -Parent $PSScriptRoot
if (!$OutputDirectory) { $OutputDirectory = Join-Path $repoRoot ('artifacts/emission-compare-probe-' + [guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh emission comparison test directory' }
$scratch = (New-Item -ItemType Directory -Path $OutputDirectory).FullName

function Write-Ops([string] $Name, [object[]] $Ops) {
    $path = Join-Path $scratch ($Name + ".json")
    ($Ops | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $path -Encoding UTF8
    return $path
}

function New-RefOp($X, $Y, $Text) {
    return [pscustomobject]@{ PageNumber = 1; X = $X; Y = $Y; FontSize = 12; CharacterSpacing = 0; DecodedText = $Text }
}
function New-CandRun($X, $BaselineY, $Text, $PdfSize) {
    return [pscustomobject]@{ Slide = 1; X = $X; BaselineY = $BaselineY; PdfFontSize = $PdfSize; LayoutFontSize = 12; Text = $Text }
}

$failures = @()
function Invoke-Case([string] $Name, [object[]] $Reference, [object[]] $Candidate, [int] $ExpectedExit, [int] $ExpectedMissing, [string] $ExtraArgs = "") {
    $refPath = Write-Ops ($Name + "-ref") $Reference
    $candPath = Write-Ops ($Name + "-cand") $Candidate
    $outPath = Join-Path $scratch ($Name + "-out.txt")
    $tool = Join-Path $repoRoot "tools/ComparePptxTextEmission.ps1"
    $toolArgs = @("-NoProfile", "-File", $tool, "-ReferenceTextOperations", $refPath, "-CandidateGlyphRuns", $candPath, "-MatchByPosition") + @($ExtraArgs -split " " | Where-Object { $_ -ne "" })
    & pwsh @toolArgs > $outPath
    $code = $LASTEXITCODE
    $missing = @(Get-Content -LiteralPath $outPath | Where-Object { $_ -match "^\s*\d+\s+missing\b" }).Count
    if ($code -ne $ExpectedExit -or $missing -ne $ExpectedMissing) {
        $script:failures += ("{0}: exit={1} (want {2}), missing={3} (want {4})" -f $Name, $code, $ExpectedExit, $missing, $ExpectedMissing)
        Write-Host ("FAIL: {0}" -f $script:failures[-1])
    }
    else {
        Write-Host ("PASS: {0} (exit={1}, missing={2})" -f $Name, $code, $missing)
    }
}

$r1 = New-RefOp 10 20 "Alpha"
$r2 = New-RefOp 30 20 "Beta"
$c1 = New-CandRun 10 20 "Alpha" 12
$c2 = New-CandRun 30 20 "Beta" 12
Invoke-Case "identical" @($r1, $r2) @($c1, $c2) 0 0
Invoke-Case "reordered" @($r1, $r2) @($c2, $c1) 0 0
Invoke-Case "deleted" @($r1, $r2) @($c1) 1 1
Invoke-Case "duplicated" @($r1) @($c1, (New-CandRun 10 20 "Alpha" 12)) 1 1
Invoke-Case "textfilter-plain" @($r1, $r2) @((New-CandRun 10 20 "Beta" 12), (New-CandRun 30 20 "Alpha" 12)) 0 0
Invoke-Case "textfilter" @($r1, $r2) @((New-CandRun 10 20 "Beta" 12), (New-CandRun 30 20 "Alpha" 12)) 1 0 "-MatchByTextThenPosition"
Invoke-Case "fontsize" @($r1) @((New-CandRun 10 20 "Alpha" 12.5)) 1 0

$tool = Join-Path $PSScriptRoot 'ComparePptxTextEmission.ps1'
$referencePath = Join-Path $scratch 'deleted-ref.json'
$candidatePath = Join-Path $scratch 'deleted-cand.json'
& $tool -ReferenceTextOperations $referencePath -CandidateGlyphRuns $candidatePath *> (Join-Path $scratch 'in-process-failure.log')
if ($LASTEXITCODE -ne 1) { throw 'A missing glyph run must report a failed in-process comparison.' }
$noFailLog = Join-Path $scratch 'in-process-no-fail.log'
& $tool -ReferenceTextOperations $referencePath -CandidateGlyphRuns $candidatePath -NoFail *> $noFailLog
if ($LASTEXITCODE -ne 0 -or (Get-Content -LiteralPath $noFailLog -Raw) -notmatch 'deltas=1') {
    throw 'NoFail must report the delta and return success in-process.'
}

if ($failures.Count -ne 0) {
    throw ("Emission comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Emission comparison adversarial checks passed (9 cases)."
