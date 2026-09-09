# Adversarial checks for the chart classifiers (T05): reordered inputs must
# classify identically (deterministic sorted output) while a duplicated op
# yields exactly one extra structure. No Office/COM or reference cache needed.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = Join-Path $repoRoot "artifacts/chart-classifier-probe"
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $scratch | Out-Null

$failures = @()
function Invoke-Classifier([string] $Tool, [string] $Name, [object[]] $Ops, [string] $ExtraArgs = "") {
    $inPath = Join-Path $scratch ($Name + "-in.json")
    $outPath = Join-Path $scratch ($Name + "-out.json")
    ($Ops | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $inPath -Encoding UTF8
    if (-not (Test-Path -LiteralPath $inPath)) { throw ("Classifier input was not written for " + $Name) }
    $toolArgs = @("-NoProfile", "-File", (Join-Path $repoRoot $Tool), "-InputPath", $inPath, "-Output", $outPath) + @($ExtraArgs -split " " | Where-Object { $_ -ne "" })
    $proc = Start-Process pwsh -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru -RedirectStandardOutput (Join-Path $scratch ($Name + "-out.txt"))
    if ($proc.ExitCode -ne 0) { throw ("Classifier case " + $Name + " exited " + $proc.ExitCode) }
    return Get-Content -Raw -LiteralPath $outPath | ConvertFrom-Json
}

function Compare-JsonFile([string] $Name, [string] $First, [string] $Second) {
    $a = [IO.File]::ReadAllBytes($First)
    $b = [IO.File]::ReadAllBytes($Second)
    $same = ($a.Length -eq $b.Length)
    if ($same) { for ($i = 0; $i -lt $a.Length; $i++) { if ($a[$i] -ne $b[$i]) { $same = $false; break } } }
    if ($same) { Write-Host ("PASS: " + $Name + " identical") }
    else { $script:failures += ($Name + ": outputs differ"); Write-Host ("FAIL: " + $Name + " outputs differ") }
}

function Assert-OneExtraBlock([string] $Name, [object[]] $Base, [object[]] $Grown) {
    $rest = [System.Collections.Generic.List[object]]::new($Grown)
    foreach ($wanted in $Base) {
        $hit = -1
        for ($i = 0; $i -lt $rest.Count; $i++) {
            $x = ($rest[$i] | ConvertTo-Json -Compress -Depth 6)
            $y = ($wanted | ConvertTo-Json -Compress -Depth 6)
            if ($x -ceq $y) { $hit = $i; break }
        }
        if ($hit -lt 0) { $script:failures += ($Name + ": baseline block missing from grown output"); Write-Host ("FAIL: " + $Name + " baseline block missing"); return }
        $rest.RemoveAt($hit)
    }
    if ($rest.Count -eq 1) { Write-Host ("PASS: " + $Name + " exactly one extra block") }
    else { $script:failures += ($Name + ": extra count=" + $rest.Count); Write-Host ("FAIL: " + $Name + " extra count=" + $rest.Count) }
}

$gfxTool = "tools/ClassifyPdfChartGraphics.ps1"
$gfxList = @(Get-Content -Raw -LiteralPath "artifacts/a1-arial-final/graphics-operations.json" | ConvertFrom-Json)
$gfxBase = Invoke-Classifier $gfxTool "gfx-base" $gfxList
$gfxReorder = Invoke-Classifier $gfxTool "gfx-reorder" $gfxList[($gfxList.Count - 1)..0]
$gfxDupe = Invoke-Classifier $gfxTool "gfx-dupe" ($gfxList + $gfxList[0])
Compare-JsonFile "gfx-reorder" (Join-Path $scratch "gfx-reorder-out.json") (Join-Path $scratch "gfx-base-out.json")
Assert-OneExtraBlock "gfx-dupe" $gfxBase $gfxDupe

$textTool = "tools/ClassifyPdfChartText.ps1"
$textList = @(Get-Content -Raw -LiteralPath "artifacts/a1-arial-final/text-operations.json" | ConvertFrom-Json)
$textBase = Invoke-Classifier $textTool "text-base" $textList
$textReorder = Invoke-Classifier $textTool "text-reorder" $textList[($textList.Count - 1)..0]
$textDupe = Invoke-Classifier $textTool "text-dupe" ($textList + $textList[0])
Compare-JsonFile "text-reorder" (Join-Path $scratch "text-reorder-out.json") (Join-Path $scratch "text-base-out.json")
Assert-OneExtraBlock "text-dupe" $textBase $textDupe

if ($failures.Count -ne 0) {
    throw ("Chart classifier adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Chart classifier adversarial checks passed (4 cases)."
