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
function New-Op([double] $X, [double] $Y, [string] $Text, [double] $Advance) {
    [pscustomobject]@{ PageNumber=1; X=$X; Y=$Y; EffectiveX=$X; EffectiveY=$Y; DecodedText=$Text; Payload=$Text; EmittedAdvancePoints=$Advance; NaturalWidthPoints=$Advance; AdjustmentTotalPoints=0; NetSpacingGapTotalPoints=0; CharacterSpacingGapTotalPoints=0; CharacterSpacing=0; FontSize=12; DecodedRuneCount=$Text.Length; TextChunkCount=1 }
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
$mixed = @(Merge-SameLineTextOperations @($unknown, (New-Op 97 660 ' ' 3)))
Assert-Equal 1 @(Select-NonWhitespacePdfTextOperations $mixed).Count 'Partly undecoded lines are retained'
Write-Host 'Markup text geometry checks passed.'
