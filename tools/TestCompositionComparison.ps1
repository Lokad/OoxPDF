# Exercise comparison preflight and identity reuse without rendering or building.
[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [string] $ComparisonScript = (Join-Path $PSScriptRoot 'CompareCompositionPortfolio.ps1')
)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) {
    $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/comparison-test-' + [Guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo = Join-Path $OutputDirectory 'repository'
$tools = Join-Path $repo 'tools'
New-Item -ItemType Directory -Path $tools -Force | Out-Null
$script = Join-Path $tools 'CompareCompositionPortfolio.ps1'
Copy-Item -LiteralPath $ComparisonScript -Destination $script
@'
param($InputPdf, $OutputDirectory, $Dpi, [switch]$SkipBuild)
'Unexpected raster stage' | Set-Content -LiteralPath ($OutputDirectory + '.called')
throw 'Preflight allowed a rendering stage'
'@ | Set-Content -LiteralPath (Join-Path $tools 'RasterizePdf.ps1')
$thresholdNames = @('docx-math-normal-fraction-inline-arial-normal', 'pptx-ladder-11-composite-chart-port')
foreach ($name in $thresholdNames) {
    $relative = 'visual-cases/cases/' + $name + '/case.json'
    $target = Join-Path $repo $relative
    New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path (Split-Path $PSScriptRoot) $relative) -Destination $target
}
function New-Case([string] $Name, [string] $Fault) {
    $directory = Join-Path $OutputDirectory $Name
    $candidate = Join-Path $directory 'candidate'
    $reference = Join-Path $directory 'reference'
    $rows = @(foreach ($id in @('first', 'second')) {
        $folder = $id + '-final-preserve'
        $run = Join-Path $candidate $folder
        $source = Join-Path $reference $folder
        New-Item -ItemType Directory -Path $run, $source -Force | Out-Null
        $pdf = Join-Path $run 'candidate.pdf'
        ('Fake candidate ' + $id) | Set-Content -LiteralPath $pdf
        if ($Fault -ne 'missing-reference' -or $id -ne 'second') {
            ('Fake reference ' + $id) | Set-Content -LiteralPath (Join-Path $source 'reference.pdf')
        }
        [pscustomobject]@{Id=$id;Mode='final-preserve';HeldOut=$false;
            InputSha256=('0' * 64);PdfSha256=(Get-FileHash -LiteralPath $pdf).Hash;Severity='Test fixture'}
    })
    if ($Fault -eq 'candidate-hash') { $rows[1].PdfSha256 = '0' * 64 }
    if ($Fault -eq 'existing-comparison') {
        $diff = Join-Path $candidate 'second-final-preserve/office-diff'
        New-Item -ItemType Directory -Path $diff | Out-Null
        'Retain this evidence' | Set-Content -LiteralPath (Join-Path $diff 'sentinel.txt')
    }
    if ($Fault -eq 'duplicate') { $rows = @($rows[0], $rows[0]) }
    @{LibrarySha256=('1' * 64);Rows=$rows} | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath (Join-Path $candidate 'outcomes.json')
    return @{Directory=$directory;CandidateDirectory=$candidate;ReferenceDirectory=$reference;Rows=$rows}
}
function Snapshot([string] $Directory) {
    return @(Get-ChildItem -LiteralPath $Directory -Recurse -File | Sort-Object FullName |
        ForEach-Object { $_.FullName + ':' + (Get-FileHash -LiteralPath $_.FullName).Hash }) -join "`n"
}
$passed = @()
foreach ($fault in @('missing-reference', 'candidate-hash', 'existing-comparison', 'duplicate')) {
    $case = New-Case $fault $fault
    $before = Snapshot $case.Directory
    $rejected = $false
    try {
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
            -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools | Out-Null
    }
    catch {
        $expected = switch ($fault) {
            'missing-reference' { 'second-final-preserve' }
            'candidate-hash' { 'Candidate identity differs: second-final-preserve' }
            'existing-comparison' { 'Comparison already exists: second-final-preserve' }
            'duplicate' { 'Duplicate comparison outcome: first-final-preserve' }
        }
        if (!$_.Exception.Message.Contains($expected)) { throw }
        $rejected = $true
    }
    if (!$rejected -or (Snapshot $case.Directory) -ne $before) {
        throw ('Preflight must reject without stage mutations: ' + $fault)
    }
    $passed += $fault
}
$case = New-Case 'identity-reuse' ''
$cache = Join-Path $case.Directory 'prior'
New-Item -ItemType Directory -Path $cache | Out-Null
$priorRows = @($case.Rows | ForEach-Object {
    @{Id=$_.Id;Mode=$_.Mode;CandidatePdfSha256=$_.PdfSha256;
        ReferencePdfSha256=(Get-FileHash -LiteralPath (Join-Path $case.ReferenceDirectory ($_.Id+'-'+$_.Mode+'/reference.pdf'))).Hash;
        Failures=@('ForegroundRecall');Metrics=@(@{ForegroundRecall=.5});Thresholds=@{minForegroundRecall=.99}}
})
@{DocxThresholdSourceSha256=(Get-FileHash (Join-Path $repo ('visual-cases/cases/'+$thresholdNames[0]+'/case.json'))).Hash;
    PptxThresholdSourceSha256=(Get-FileHash (Join-Path $repo ('visual-cases/cases/'+$thresholdNames[1]+'/case.json'))).Hash;
    Rows=$priorRows} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $cache 'office-comparisons.json')
& $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
    -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
$report = Get-Content -LiteralPath (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
if ($report.Rows.Count -ne 2 -or @($report.Rows | Where-Object {
    !$_.ReusedComparison -or $_.Failures.Count -ne 1 -or $_.Failures[0] -ne 'ForegroundRecall'
}).Count -or @(Get-ChildItem -LiteralPath $case.Directory -Recurse -Filter '*.called').Count) {
    throw 'Identical comparisons must reuse retained failures without rendering'
}
$passed += 'identity-reuse'
@{Passed=$passed;LibraryBuilds=0;Conversions=0;Rasterizations=0;Scope='Fake PDF inputs; tool orchestration only'} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
Write-Output ($passed.Count.ToString() + ' comparison orchestration checks passed; no build or rendering.')
