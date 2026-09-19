# Adversarial checks for the chart classifiers (T05): reordered inputs must
# classify identically (deterministic sorted output) while a duplicated op
# yields exactly one extra structure. Synthetic operations keep this gate
# independent of previous renders, Office/COM, and the reference cache.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = Join-Path $repoRoot "artifacts/chart-classifier-probe"
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $scratch | Out-Null

$failures = @()
function Invoke-Classifier([string] $Tool, [string] $Name, [object[]] $Ops, [string[]] $ExtraArgs = @()) {
    $inPath = Join-Path $scratch ($Name + "-in.json")
    $outPath = Join-Path $scratch ($Name + "-out.json")
    ($Ops | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $inPath -Encoding UTF8
    if (-not (Test-Path -LiteralPath $inPath)) { throw ("Classifier input was not written for " + $Name) }
    $toolArgs = @("-NoProfile", "-File", (Join-Path $repoRoot $Tool), "-InputPath", $inPath, "-Output", $outPath) + $ExtraArgs
    & pwsh @toolArgs > (Join-Path $scratch ($Name + "-out.txt"))
    if ($LASTEXITCODE -ne 0) { throw ("Classifier case " + $Name + " exited " + $LASTEXITCODE) }
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

function New-Graphic([string] $Kind, [double] $MinX, [double] $MinY, [double] $MaxX, [double] $MaxY) {
    $isStroke = $Kind -eq "Stroke"
    return [pscustomobject]@{
        PageNumber = 1; Kind = $Kind; Operator = $(if ($isStroke) { "S" } else { "f" })
        SegmentCount = $(if ($isStroke) { 1 } else { 4 }); MoveCount = 1
        LineCount = $(if ($isStroke) { 1 } else { 3 }); CurveCount = 0
        CloseCount = $(if ($isStroke) { 0 } else { 1 }); LineWidth = 1
        MinX = $MinX; MinY = $MinY; MaxX = $MaxX; MaxY = $MaxY
    }
}

function New-Text([double] $X, [double] $Y, [string] $Text, [double] $FontSize = 12) {
    return [pscustomobject]@{
        PageNumber = 1; Operator = "Tj"; X = $X; Y = $Y
        DecodedText = $Text; Font = "F1"; FontSize = $FontSize; CharacterSpacing = 0
    }
}

function Assert-Kinds([string] $Name, [object[]] $Structures, [string[]] $Expected) {
    $actualKinds = @($Structures | ForEach-Object { $_.Kind } | Sort-Object) -join ","
    $expectedKinds = @($Expected | Sort-Object) -join ","
    if ($actualKinds -ceq $expectedKinds) { Write-Host ("PASS: " + $Name + " expected structures") }
    else { $script:failures += ($Name + ": expected " + $expectedKinds + ", got " + $actualKinds); Write-Host ("FAIL: " + $script:failures[-1]) }
}

$gfxTool = "tools/ClassifyPdfChartGraphics.ps1"
# A marker inside two axes. Duplicating the marker adds one structure without
# changing the derived plot box or its axes.
$gfxList = @(
    (New-Graphic "Fill" 146 146 154 154),
    (New-Graphic "Stroke" 100 100 300 100),
    (New-Graphic "Stroke" 100 100 100 300)
)
$gfxBase = Invoke-Classifier $gfxTool "gfx-base" $gfxList
Assert-Kinds "gfx-base" $gfxBase @("MarkerCandidate", "HorizontalLine", "VerticalLine", "AxisPairPlotBoxCandidate")
$gfxReorder = Invoke-Classifier $gfxTool "gfx-reorder" $gfxList[($gfxList.Count - 1)..0]
$gfxDupe = Invoke-Classifier $gfxTool "gfx-dupe" ($gfxList + $gfxList[0])
Compare-JsonFile "gfx-reorder" (Join-Path $scratch "gfx-reorder-out.json") (Join-Path $scratch "gfx-base-out.json")
Assert-OneExtraBlock "gfx-dupe" $gfxBase $gfxDupe

$textTool = "tools/ClassifyPdfChartText.ps1"
$textList = @(
    (New-Text 150 150 "42"),
    (New-Text 180 330 "Synthetic chart" 18),
    (New-Text 70 150 "10"),
    (New-Text 180 80 "Q1")
)
$textArgs = @("-ChartStructures", (Join-Path $scratch "gfx-base-out.json"))
$textBase = Invoke-Classifier $textTool "text-base" $textList $textArgs
Assert-Kinds "text-base" $textBase @("DataLabelText", "ChartTitleText", "ValueAxisTickLabel", "CategoryAxisTickLabel")
$textReorder = Invoke-Classifier $textTool "text-reorder" $textList[($textList.Count - 1)..0] $textArgs
$textDupe = Invoke-Classifier $textTool "text-dupe" ($textList + $textList[0]) $textArgs
Compare-JsonFile "text-reorder" (Join-Path $scratch "text-reorder-out.json") (Join-Path $scratch "text-base-out.json")
Assert-OneExtraBlock "text-dupe" $textBase $textDupe

if ($failures.Count -ne 0) {
    throw ("Chart classifier adversarial checks failed: " + ($failures -join "; "))
}
Write-Host "Chart classifier adversarial checks passed (6 cases)."
