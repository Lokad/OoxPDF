# Adversarial checks for the chart classifiers (T05): reordered inputs must
# classify identically (deterministic sorted output) while a duplicated op
# yields exactly one extra structure. Synthetic operations keep this gate
# independent of previous renders, Office/COM, and the reference cache.
[CmdletBinding()]
param([string] $OutputDirectory)

$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
$repoRoot = Split-Path -Parent $PSScriptRoot
if (!$OutputDirectory) { $OutputDirectory = Join-Path $repoRoot ('artifacts/chart-classifier-probe-' + [guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh classifier test directory' }
$scratch = (New-Item -ItemType Directory -Path $OutputDirectory).FullName
$invocations = [Collections.Generic.List[object]]::new()

$failures = @()
function Invoke-Classifier([string] $Tool, [string] $Name, [object[]] $Ops, [string[]] $ExtraArgs = @()) {
    $inPath = Join-Path $scratch ($Name + "-in.json")
    $outPath = Join-Path $scratch ($Name + "-out.json")
    ($Ops | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $inPath -Encoding UTF8
    if (-not (Test-Path -LiteralPath $inPath)) { throw ("Classifier input was not written for " + $Name) }
    $toolArgs = @("-NoProfile", "-File", (Join-Path $repoRoot $Tool), "-InputPath", $inPath, "-Output", $outPath) + $ExtraArgs
    & pwsh @toolArgs > (Join-Path $scratch ($Name + "-out.txt"))
    if ($LASTEXITCODE -ne 0) { throw ("Classifier case " + $Name + " exited " + $LASTEXITCODE) }
    $invocations.Add([pscustomobject]@{Id=$Name;Tool=$Tool;InputSha256=(Get-FileHash $inPath).Hash;OutputSha256=(Get-FileHash $outPath).Hash})
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

$swatch = New-Graphic "Fill" 312 146 320 154
$swatchOps = @($swatch) + $gfxList[1..2]
$swatchKinds = @("MarkerCandidate", "HorizontalLine", "VerticalLine", "AxisPairPlotBoxCandidate", "LegendSwatchCandidate")
$caption = New-Text 324 150 "North"
$number = New-Text 324 150 "42"
foreach ($variant in @(
    @{ Name = "plot"; Text = @(); Keys = @() },
    @{ Name = "caption"; Text = @($caption); Keys = @() },
    @{ Name = "number"; Text = @($number); Keys = @("DataLabelLegendKeyFillCandidate") },
    @{ Name = "both"; Text = @($caption, $number); Keys = @("DataLabelLegendKeyFillCandidate") }
)) {
    $name = "gfx-swatch-" + $variant.Name
    $textPath = Join-Path $scratch "$name-text.json"
    ConvertTo-Json -InputObject $variant.Text | Set-Content -LiteralPath $textPath -Encoding UTF8
    $classified = Invoke-Classifier $gfxTool $name $swatchOps @("-TextOperations", $textPath)
    Assert-Kinds $name $classified ($swatchKinds + $variant.Keys)
}
$captionPath = Join-Path $scratch "gfx-swatch-caption-text.json"
$twoSwatches = Invoke-Classifier $gfxTool "gfx-swatch-duplicate-paint" ($swatchOps + $swatch) @("-TextOperations", $captionPath)
Assert-Kinds "gfx-swatch-duplicate-paint" $twoSwatches ($swatchKinds + @("MarkerCandidate", "LegendSwatchCandidate"))

$textOnlyPath = Join-Path $scratch "gfx-text-only-swatch-text.json"
ConvertTo-Json -InputObject @((New-Text 158 150 "North"), (New-Text 158 150 "42")) |
    Set-Content -LiteralPath $textOnlyPath -Encoding UTF8
$textOnlySwatch = Invoke-Classifier $gfxTool "gfx-text-only-swatch" $gfxList @("-TextOperations", $textOnlyPath)
Assert-Kinds "gfx-text-only-swatch" $textOnlySwatch ($swatchKinds + "DataLabelLegendKeyFillCandidate")

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

function Assert-Rows([string] $Name, [object[]] $Expected, [object[]] $Actual) {
    $left = @($Expected | Sort-Object PageNumber, Kind, MinY, MinX, MaxY, MaxX | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 6 }) -join "`n"
    $right = @($Actual | Sort-Object PageNumber, Kind, MinY, MinX, MaxY, MaxX | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 6 }) -join "`n"
    if ($left -cne $right) { throw ($Name + ': adding another page changed page-owned classification') }
    Write-Host ('PASS: ' + $Name + ' retains independent page results')
}
function On-Page($Items, [int] $Page, [double] $Shift = 0) {
    foreach ($item in $Items) {
        $copy = $item.PSObject.Copy()
        $copy.PageNumber = $Page
        foreach ($name in @('MinX','MinY','MaxX','MaxY','X','Y')) {
            if ($null -ne $copy.PSObject.Properties[$name]) { $copy.$name = [double]$copy.$name + $Shift }
        }
        $copy
    }
}
$pageTwoGraphics = @(On-Page $gfxList 2 400)
$gfxTwo = Invoke-Classifier $gfxTool 'gfx-page-two' $pageTwoGraphics
$joinedGraphics = @($gfxList) + $pageTwoGraphics
$gfxJoined = Invoke-Classifier $gfxTool 'gfx-joined' $joinedGraphics
Assert-Rows 'graphics composition' (@($gfxBase) + @($gfxTwo)) $gfxJoined
Invoke-Classifier $gfxTool 'gfx-joined-reordered' $joinedGraphics[($joinedGraphics.Count-1)..0] | Out-Null
Compare-JsonFile 'graphics page reordering' (Join-Path $scratch 'gfx-joined-out.json') (Join-Path $scratch 'gfx-joined-reordered-out.json')
Invoke-Classifier $gfxTool 'gfx-explicit-one' $joinedGraphics @('-PageNumber','1') | Out-Null
Invoke-Classifier $gfxTool 'gfx-explicit-two' $joinedGraphics @('-PageNumber','2') | Out-Null
Compare-JsonFile 'graphics page-1 filter' (Join-Path $scratch 'gfx-base-out.json') (Join-Path $scratch 'gfx-explicit-one-out.json')
Compare-JsonFile 'graphics page-2 filter' (Join-Path $scratch 'gfx-page-two-out.json') (Join-Path $scratch 'gfx-explicit-two-out.json')
$identicalPageTwo = @(On-Page $gfxList 2)
$repeated = Invoke-Classifier $gfxTool 'gfx-repeated-geometry' (@($gfxList) + $identicalPageTwo)
Assert-Rows 'identical geometry on different pages' (@($gfxBase) + @(On-Page $gfxBase 2)) $repeated

$pageTwoText = @(On-Page $textList 2 400)
$textTwo = Invoke-Classifier $textTool 'text-page-two' $pageTwoText @('-ChartStructures',(Join-Path $scratch 'gfx-page-two-out.json'))
$joinedText = @($textList) + $pageTwoText
$joinedTextArgs = @('-ChartStructures',(Join-Path $scratch 'gfx-joined-out.json'))
$textJoined = Invoke-Classifier $textTool 'text-joined' $joinedText $joinedTextArgs
Assert-Rows 'text composition' (@($textBase) + @($textTwo)) $textJoined
Invoke-Classifier $textTool 'text-joined-reordered' $joinedText[($joinedText.Count-1)..0] $joinedTextArgs | Out-Null
Compare-JsonFile 'text page reordering' (Join-Path $scratch 'text-joined-out.json') (Join-Path $scratch 'text-joined-reordered-out.json')
Invoke-Classifier $textTool 'text-explicit-two' $joinedText ($joinedTextArgs + @('-PageNumber','2')) | Out-Null
Compare-JsonFile 'text page-2 filter' (Join-Path $scratch 'text-page-two-out.json') (Join-Path $scratch 'text-explicit-two-out.json')
$withoutPlot = Invoke-Classifier $textTool 'text-without-own-plot' @(On-Page @($textList[0]) 3) $joinedTextArgs
Assert-Kinds 'text does not borrow another page plot' @($withoutPlot) @('ChartText')

$crossingGraphics = Invoke-Classifier $gfxTool 'gfx-crossing-only' @(
    (New-Graphic 'Stroke' 100 100 100 300),
    (New-Graphic 'Stroke' 100 275 300 275)
)
if (@($crossingGraphics | Where-Object Kind -like '*PlotBoxCandidate').Kind -cne 'CrossingAxisPlotBoxCandidate') {
    throw 'Crossing-axis control must produce only the crossing plot bounds.'
}
$crossingTick = Invoke-Classifier $textTool 'text-crossing-only' @((New-Text 70 150 '10')) `
    @('-ChartStructures', (Join-Path $scratch 'gfx-crossing-only-out.json'))
Assert-Kinds 'crossing-axis tick' @($crossingTick) @('ValueAxisTickLabel')
if ($crossingTick.RegionIndex -ne 0) { throw 'The crossing-axis tick must belong to the graphics region.' }

$left = @{ PageNumber = 1; Kind = 'AxisPairPlotBoxCandidate'; RegionIndex = 0; MinX = 100; MinY = 100; MaxX = 300; MaxY = 300 }
$right = @{ PageNumber = 1; Kind = 'PolarPlotBoxCandidate'; RegionIndex = 1; MinX = 350; MinY = 100; MaxX = 550; MaxY = 300 }
$swatch = @{ PageNumber = 1; Kind = 'MarkerCandidate'; RegionIndex = 1; MinX = 328; MinY = 156; MaxX = 332; MaxY = 164 }
$mixedStructures = Join-Path $scratch 'mixed-regions.json'
ConvertTo-Json -InputObject @($left, $right, $swatch) | Set-Content -LiteralPath $mixedStructures -Encoding UTF8
$mixedLabels = @(Invoke-Classifier $textTool 'text-mixed-regions' @(
    (New-Text 70 150 '10'), (New-Text 315 160 '20'), (New-Text 325 270 'Ambiguous'),
    (New-Text 337 160 'North'), (New-Text 460 150 '40%')
) @('-ChartStructures', $mixedStructures) | Sort-Object MinX)
$mixedKinds = @($mixedLabels | ForEach-Object { "$($_.Kind):$($_.RegionIndex)" }) -join ','
if ($mixedKinds -cne 'ValueAxisTickLabel:0,RightSideText:0,ChartText:,LegendText:1,DataLabelText:1') {
    throw "Mixed chart regions borrowed neighboring bounds or label cues: $mixedKinds"
}
Write-Host 'PASS: mixed chart regions retain their own axes, legend and numeric labels'

$bar = $right.Clone()
$bar.Kind = 'AxisPairPlotBoxCandidate'
$gridlines = $right.Clone()
$gridlines.Kind = 'VerticalGridlineGroupCandidate'
$cartesianStructures = Join-Path $scratch 'cartesian-regions.json'
ConvertTo-Json -InputObject @($left, $bar, $gridlines) | Set-Content -LiteralPath $cartesianStructures -Encoding UTF8
$cartesianLabels = @(Invoke-Classifier $textTool 'text-cartesian-regions' @((New-Text 70 150 '10'), (New-Text 340 150 'West')) `
    @('-ChartStructures', $cartesianStructures) | Sort-Object MinX)
$cartesianKinds = @($cartesianLabels | ForEach-Object { "$($_.Kind):$($_.RegionIndex)" }) -join ','
if ($cartesianKinds -cne 'ValueAxisTickLabel:0,CategoryAxisTickLabel:1') {
    throw "Cartesian chart regions borrowed a neighboring axis direction: $cartesianKinds"
}
Write-Host 'PASS: Cartesian chart regions retain their own axis directions'

$moved = @($gfxList) + @(On-Page $gfxList 2 600)
Invoke-Classifier $gfxTool 'gfx-moved-page-two' $moved | Out-Null
$compareLog = Join-Path $scratch 'moved-page-two-comparison.log'
& pwsh -NoProfile -File (Join-Path $repoRoot 'tools/ComparePdfGraphicsOperations.ps1') `
    -Reference (Join-Path $scratch 'gfx-joined-out.json') -Candidate (Join-Path $scratch 'gfx-moved-page-two-out.json') `
    -Kinds AxisPairPlotBoxCandidate -MatchByBounds *> $compareLog
$comparisonExit = $LASTEXITCODE
if ($comparisonExit -ne 1) { throw 'Moved page-2 chart must fail the classified axis-pair comparison' }

[ordered]@{Invocations=@($invocations.ToArray());MovedPageComparisonExit=$comparisonExit;Failures=@($failures)} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $scratch 'checks.json') -Encoding utf8
if ($failures.Count -ne 0) {
    throw ("Chart classifier adversarial checks failed: " + ($failures -join "; "))
}
Write-Host ("Chart classifier checks passed ({0} invocations and moved-page refusal)." -f $invocations.Count)
