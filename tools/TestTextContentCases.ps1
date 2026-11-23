# Malformed expectations must fail before the gate builds or converts anything.
$ErrorActionPreference = 'Stop'
$gate = Join-Path $PSScriptRoot 'CheckPdfTextContent.ps1'
$root = Join-Path (Split-Path $PSScriptRoot) ('artifacts/text-content-cases-test-' + [guid]::NewGuid().ToString('N'))
$cases = Join-Path $root 'cases'
$output = Join-Path $root 'output'
New-Item -ItemType Directory -Path $cases | Out-Null

$valid = Get-Content -Raw (Join-Path $PSScriptRoot 'text-content-cases/docx-plain.json') | ConvertFrom-Json -AsHashtable
$valid | ConvertTo-Json | Set-Content (Join-Path $cases 'a-valid.json')
$invalid = Join-Path $cases 'z-invalid.json'
$expectations = [ordered]@{
    missing = @{}
    empty = @{ expected = @() }
    scalar = @{ expected = 'Hello' }
    number = @{ expected = @('Hello', 7) }
    nullEntry = @{ expected = @('Hello', $null) }
    blank = @{ expected = @('Hello', ' ') }
}
foreach ($name in $expectations.Keys) {
    (@{ input = $valid.input } + $expectations[$name]) | ConvertTo-Json | Set-Content -LiteralPath $invalid
    $failure = $null
    try { & $gate -CasesDir $cases -Out $output -SkipBuild }
    catch { $failure = $_.Exception.Message }
    if ($failure -ne "Case 'z-invalid.json' requires a nonempty expected array of text strings." -or
        (Test-Path -LiteralPath $output)) {
        throw "Expected manifest preflight to reject '$name' before producing output. Observed '$failure'."
    }
}
Write-Output "Text-content manifest checks passed for $($expectations.Count) invalid expectations."

. (Join-Path $PSScriptRoot 'PdfTextGeometry.ps1')
function New-TextShow([string] $Text, [double] $Y = 700d, [double] $X = 72d) {
    [pscustomobject]@{ PageNumber=1; DecodedText=$Text; EffectiveX=$X; EffectiveY=$Y;
        EffectiveA=1d; EffectiveB=0d; FontSize=12d }
}
function Assert-Missing($Operations, $Expected, $Missing, [string] $Label) {
    $actual = @(Get-MissingOrderedPdfTextPhrases $Operations $Expected)
    if ((ConvertTo-Json -InputObject $actual -Compress) -cne (ConvertTo-Json -InputObject $Missing -Compress)) {
        throw "Unexpected missing phrases for $Label. Observed [$($actual -join ', ')]."
    }
}
Assert-Missing @(New-TextShow 'Firsthand Secondhand') @('First','Second') @('First','Second') 'longer words'
Assert-Missing @(New-TextShow 'Second First') @('First','Second') @('Second') 'reversed order'
Assert-Missing @(New-TextShow 'First') @('First','First') @('First') 'missing repeated witness'
Assert-Missing @(New-TextShow 'First First') @('First','First') @() 'distinct repeated witnesses'
Assert-Missing @((New-TextShow 'First'), (New-TextShow 'Second' 682d)) @('First','Second') @() 'printed line reset'
Assert-Missing @((New-TextShow 'First'), (New-TextShow 'hand' 703d 92d)) @('First') @('First') 'small baseline change'
Assert-Missing @(New-TextShow 'Alpha  beta   gamma') @('Alpha  beta   gamma') @() 'exact interior spaces'
Assert-Missing @(New-TextShow 'Alpha beta gamma') @('Alpha  beta   gamma') @('Alpha  beta   gamma') 'collapsed spaces'

$owned = New-TextShow 'Owned witness'
$unassigned = New-TextShow 'Owned witness'
$unassigned.PageNumber = $null
$invalidPage = New-TextShow 'Owned witness'
$invalidPage.PageNumber = 0
Assert-Missing @($unassigned, $invalidPage) @('Owned witness') @('Owned witness') 'no owning page'
Assert-Missing @($unassigned, $owned) @('Owned witness','Owned witness') @('Owned witness') 'only one page-owned copy'
if (@(Get-MissingPdfTextPhrases @($unassigned, $owned) @('Owned witness')).Count -ne 0 -or
    @(Get-MissingPdfTextPhrases @($unassigned, $invalidPage) @('Owned witness')).Count -ne 1) {
    throw 'Unordered witnesses must come from page-owned text'
}
Write-Output 'Ordered source-text matching checks passed.'
