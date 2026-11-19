# Exercise the markup gate's production geometry functions without starting Office,
# converting documents, or reading/writing the reference cache.
$ErrorActionPreference = 'Stop'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'CompareCachedDocxMarkupReference.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Markup comparison script did not parse.' }
foreach ($function in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    . ([scriptblock]::Create($function.Extent.Text))
}
. (Join-Path $PSScriptRoot 'PdfTextGeometry.ps1')
function New-Op([double] $X, [double] $Y, [string] $Text, [double] $Advance, [double] $Scale = 1d) {
    [pscustomobject]@{ PageNumber=1; X=$X; Y=$Y; EffectiveX=$X; EffectiveY=$Y; EffectiveA=$Scale; EffectiveB=0d; EffectiveC=0d; EffectiveD=$Scale; DecodedText=$Text; Payload=$Text; EmittedAdvancePoints=$Advance; NaturalWidthPoints=$Advance; AdjustmentTotalPoints=0; NetSpacingGapTotalPoints=0; CharacterSpacingGapTotalPoints=0; CharacterSpacing=0; FontSize=12; DecodedRuneCount=$Text.Length; TextChunkCount=1 }
}
function Assert-Equal($Expected, $Actual, [string] $Name) {
    if ($Expected -ne $Actual) { throw "$Name expected $Expected; actual $Actual" }
    Write-Host "PASS $Name"
}
$reference = @((New-Op 72 700 'Hello' 25), (New-Op 97 700 ' ' 3), (New-Op 100 700 'world' 25), (New-Op 72 680 'Second' 30), (New-Op 72 660 ' ' 3))
$candidate = @((New-Op 72 700 'Hello world' 53), (New-Op 72 680 'Second' 30))
$baseline = Get-TextBaselinePageSummary $reference 1
Assert-Equal 680 $baseline.LastBaselineY 'Whitespace-only operators do not set the last painted-text baseline'
Assert-Equal 20 $baseline.BodyFrameHeightUsedPoints 'Painted-text span ignores trailing blank lines'
Assert-Equal 5 $baseline.TextOperationCount 'Raw operator count is retained'
$delta = New-TextGateDeltaSummary $reference $candidate
Assert-Equal 0 $delta.OperationCountDelta 'Split words and a trailing blank line align'
Assert-Equal 0 $delta.MaxBaselineDelta 'Painted line baselines align'
Assert-Equal 0 $delta.MaxEmittedAdvanceDelta 'Inline whitespace contributes to advance'
$split = @((New-Op 100 700.02 'world' 25), (New-Op 97 700.04 ' ' 3), (New-Op 72 700.06 'Hello' 25), $reference[3])
$near = @((New-Op 72 700.04 'Hello world' 53), $candidate[1])
$rounding = New-TextGateDeltaSummary $split $near
Assert-Equal 0 $rounding.OperationCountDelta 'Baseline rounding does not split a line'
Assert-Equal 0 $rounding.MaxXDelta 'The first painted run anchors the horizontal position'
Assert-Equal 0 $rounding.MaxEmittedAdvanceDelta 'The merged line retains all advance'
$chain = @((New-Op 72 700 'A' 5), (New-Op 77 700.08 'B' 5), (New-Op 82 700.16 'C' 5))
$distinct = @(Merge-PdfSameLineOperations $chain -UseEffectiveMatrix)
Assert-Equal 2 $distinct.Count 'A baseline band cannot chain to another line'
Assert-Equal 'AB' $distinct[0].DecodedText 'Only operations within the anchored band merge'
$expanded = @((New-Op 72 700 'A' 5 2), (New-Op 82 700 'B' 5 2))
$expandedDelta = New-TextGateDeltaSummary $expanded @((New-Op 72 700 'AB' 10 2))
Assert-Equal 0 $expandedDelta.OperationCountDelta 'Effective advance joins expanded text'
Assert-Equal 0 $expandedDelta.MaxEmittedAdvanceDelta 'Merged metrics retain raw advances'
$contracted = @((New-Op 72 700 'A' 10 .5), (New-Op 82 700 'B' 10 .5))
Assert-Equal 2 @(Merge-PdfSameLineOperations $contracted -UseEffectiveMatrix).Count 'Effective gaps separate contracted text'
Assert-Equal 1 @(Merge-PdfSameLineOperations $contracted).Count 'Raw comparison measures raw gaps'
$rotated = New-Op 72 700 'Word' 20
$rotated.FontSize = 24
$rotated.EffectiveA = 0d
$rotated.EffectiveB = .5d
$rotated.EffectiveC = -.5d
$rotated.EffectiveD = 1.5d
$referenceWord = New-Op 72 700 'Word' 20
$physical = New-TextGateDeltaSummary @($referenceWord) @($rotated)
Assert-Equal 0 $physical.MaxFontSizeDelta 'Physical font size excludes rotation and shear'
Assert-Equal .5d $physical.MedianReferenceToCandidateFontSizeRatio 'Raw font size ratio remains available'
$rotated.FontSize = 24.5d
$drifted = New-TextGateDeltaSummary @($referenceWord) @($rotated)
Assert-Equal .25d $drifted.MaxFontSizeDelta 'Physical size still exposes an actual size change'
Assert-Equal 'failed' (New-GateResult 'font-size' 'text' $drifted.MaxFontSizeDelta .1d 'points').Status 'Physical size uses the existing threshold'
$collapsedSize = New-Op 72 700 'Word' 20
$collapsedSize.EffectiveA = 0d
$collapsed = New-TextGateDeltaSummary @($referenceWord, $reference[3]) @($collapsedSize, $reference[3])
Assert-Equal 'failed' (New-GateResult 'font-size' 'text' $collapsed.MaxFontSizeDelta .1d 'points').Status 'A zero-length text baseline fails even after a finite pair'
$rotated.PSObject.Properties.Remove('EffectiveD')
Assert-Equal 12.5d (New-TextGateDeltaSummary @($referenceWord) @($rotated)).MaxFontSizeDelta 'A partial matrix uses the raw font size'
$rotated.PSObject.Properties.Remove('FontSize')
Assert-Equal $null (New-TextGateDeltaSummary @($referenceWord) @($rotated)).MaxFontSizeDelta 'Missing font size stays unmeasured'
$missing = New-TextGateDeltaSummary $reference @($candidate[0])
Assert-Equal 1 $missing.OperationCountDelta 'A missing painted line still fails'
$moved = New-TextGateDeltaSummary $reference @($candidate[0], (New-Op 72 679 'Second' 30))
Assert-Equal 1 $moved.MaxBaselineDelta 'A moved painted line still has its actual delta'
$blank = Get-TextBaselinePageSummary @((New-Op 72 660 ' ' 3)) 1
Assert-Equal $null $blank.LastBaselineY 'Blank-only pages have no painted baseline'
$unknown = New-Op 72 660 'Encoded glyphs' 25
$unknown.PSObject.Properties.Remove('DecodedText')
Assert-Equal 660 (Get-TextBaselinePageSummary @($unknown) 1).LastBaselineY 'Undecoded operators remain measurable'
$unknown | Add-Member -NotePropertyName DecodedText -NotePropertyValue $null
$mixed = @(Merge-PdfSameLineOperations @($unknown, (New-Op 97 660 ' ' 3)) -UseEffectiveMatrix)
Assert-Equal 1 @(Select-NonWhitespacePdfTextOperations $mixed).Count 'Partly undecoded lines are retained'
Write-Host 'Markup text geometry checks passed.'
