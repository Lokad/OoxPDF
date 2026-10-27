# Adversarial checks for ComparePdfGraphicsOperations (T05): reordered runs
# must match robustly while duplicated, deleted, kind-mismatched, or moved
# runs fail loudly with reported (not invented) deltas. Translated and scaled
# equivalents (Fill and Clip, with and without path-coordinate matching) must
# likewise fail loudly: the comparison has no transform normalization, so a
# shifted path reports its real delta instead of passing or crashing.
# Page moves and unavailable required evidence must fail too.
# No build, Office/COM or reference cache needed.
[CmdletBinding()]
param([string] $OutputDirectory)

$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
$repoRoot = Split-Path -Parent $PSScriptRoot
if (!$OutputDirectory) { $OutputDirectory = Join-Path $repoRoot ('artifacts/graphics-compare-probe-' + [guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh graphics comparison test directory' }
$scratch = (New-Item -ItemType Directory -Path $OutputDirectory).FullName
$checks = [Collections.Generic.List[object]]::new()

function Write-Ops([string] $Name, [object[]] $Ops) {
    $path = Join-Path $scratch ($Name + ".json")
    ConvertTo-Json -InputObject @($Ops) -Depth 6 | Set-Content -LiteralPath $path -Encoding UTF8
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
function Invoke-Case([string] $Name, [object[]] $Reference, [object[]] $Candidate, [int] $ExpectedExit, [int] $ExpectedMissing, [string] $ExtraArgs = "", [switch] $ByIndex, [string] $ExpectedError) {
    $refPath = Write-Ops ($Name + "-ref") $Reference
    $candPath = Write-Ops ($Name + "-cand") $Candidate
    $outPath = Join-Path $scratch ($Name + "-out.txt")
    $tool = Join-Path $repoRoot "tools/ComparePdfGraphicsOperations.ps1"
    $toolArgs = @("-NoProfile", "-File", $tool, "-Reference", $refPath, "-Candidate", $candPath) + @($ExtraArgs -split " " | Where-Object { $_ -ne "" })
    if (!$ByIndex) { $toolArgs += '-MatchByBounds' }
    & pwsh @toolArgs *> $outPath
    $code = $LASTEXITCODE
    $missing = @(Get-Content -LiteralPath $outPath | Where-Object { $_ -match "^\s*\d+\s+missing\b" }).Count
    $errorOk = !$ExpectedError -or (Get-Content -Raw -LiteralPath $outPath) -match $ExpectedError
    $checks.Add([pscustomobject]@{Id=$Name;ExpectedExit=$ExpectedExit;ActualExit=$code;ExpectedMissing=$ExpectedMissing;ActualMissing=$missing;ErrorVerified=$errorOk})
    if ($code -ne $ExpectedExit -or $missing -ne $ExpectedMissing -or !$errorOk) {
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

$page2 = $a.PSObject.Copy(); $page2.PageNumber = 2
Invoke-Case 'second-page-identical' @($page2) @($page2) 0 0
Invoke-Case 'moved-page' @($a) @($page2) 1 2
Invoke-Case 'moved-page-index' @($a) @($page2) 1 0 -ByIndex
Invoke-Case 'page-local-reordering' @($a,$page2) @($page2,$a) 0 0
$page2Other = $b.PSObject.Copy(); $page2Other.PageNumber = 2
Invoke-Case 'pages-with-swapped-geometry' @($a,$page2Other) @($b,$page2) 1 0
Invoke-Case 'explicit-page-unchanged' @($a,$page2Other) @($a,$page2) 0 0 '-PageNumber 1'
Invoke-Case 'explicit-page-missing' @($a) @($page2) 1 1 '-PageNumber 1'

$zero = New-Rect 'Fill' 0 0 0 0
Invoke-Case 'valid-zero-bounds' @($zero) @($zero) 0 0
foreach ($field in @('PageNumber','MinX','MinY','MaxX','MaxY')) {
    $missing = $zero.PSObject.Copy(); $missing.PSObject.Properties.Remove($field)
    Invoke-Case ('missing-'+$field) @($zero) @($missing) 1 0 -ExpectedError ('finite '+$field)
}
foreach ($value in @($null,'','NaN','Infinity','-Infinity')) {
    $invalid = $zero.PSObject.Copy(); $invalid.MinX = $value
    $label = if ($null -eq $value) { 'null' } elseif ($value -eq '') { 'empty' } else { $value }
    Invoke-Case ('invalid-bounds-'+$label) @($zero) @($invalid) 1 0 -ExpectedError 'finite MinX'
}
foreach ($value in @(0,-1,1.5,2147483648)) {
    $invalid = $zero.PSObject.Copy(); $invalid.PageNumber = $value
    Invoke-Case ('invalid-page-'+$value) @($zero) @($invalid) 1 0 -ExpectedError 'positive integer PageNumber'
}
$hashed = $a.PSObject.Copy(); $hashed | Add-Member TextHash 'ABC'
$differentHash = $hashed.PSObject.Copy(); $differentHash.TextHash = 'DEF'
Invoke-Case 'hash-identical' @($hashed) @($hashed) 0 0 '-MatchTextHash'
Invoke-Case 'hash-different' @($hashed) @($differentHash) 1 0 '-MatchTextHash'
Invoke-Case 'hash-missing-candidate' @($hashed) @($a) 1 0 '-MatchTextHash' -ExpectedError 'TextHash is required'
Invoke-Case 'hash-missing-reference' @($a) @($hashed) 1 0 '-MatchTextHash' -ExpectedError 'TextHash is required'
Invoke-Case 'hash-missing-both' @($a) @($a) 1 0 '-MatchTextHash' -ExpectedError 'TextHash is required'
$emptyHash = $hashed.PSObject.Copy(); $emptyHash.TextHash = ' '
Invoke-Case 'hash-empty' @($hashed) @($emptyHash) 1 0 '-MatchTextHash' -ExpectedError 'TextHash is required'
Invoke-Case 'optional-hash-disabled' @($hashed) @($a) 0 0

$checks | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $scratch 'checks.json') -Encoding utf8
if ($failures.Count -ne 0) {
    throw ("Graphics comparison adversarial checks failed: " + ($failures -join "; "))
}
Write-Host ("Graphics comparison checks passed ({0} cases)." -f $checks.Count)
