# Check ownership, source identities and complete preflight without a build/conversion.
[CmdletBinding()]
param([string] $OutputDirectory, [string] $JoinScript = (Join-Path $PSScriptRoot 'JoinCompositionPortfolio.ps1'))
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) { $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/composition-join-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$root = (Resolve-Path -LiteralPath $OutputDirectory).Path
$manifest = Join-Path $root 'manifest.json'
@(@{Id='one';Input='one.docx';Sha256=('A'*64)},@{Id='two';Input='two.docx';Sha256=('B'*64)}) |
    ConvertTo-Json | Set-Content -LiteralPath $manifest
$sources = @()
foreach ($id in @('one','two')) {
    $source = Join-Path $root $id
    $run = Join-Path $source ($id+'-all-word-compatible')
    foreach ($folder in @('inspect','raster','office-diff')) { New-Item -ItemType Directory -Path (Join-Path $run $folder) -Force | Out-Null }
    [IO.File]::WriteAllText((Join-Path $run 'candidate.pdf'), 'fake PDF '+$id)
    [IO.File]::WriteAllText((Join-Path $run 'inspect.log'), 'one page')
    [IO.File]::WriteAllText((Join-Path $run 'inspect/text-operations.json'), '[]')
    [IO.File]::WriteAllText((Join-Path $run 'diagnostics.json'), '[]')
    [IO.File]::WriteAllText((Join-Path $run 'raster/page-001.png'), 'raster remains at its original path')
    [IO.File]::WriteAllText((Join-Path $run 'office-diff/metrics.json'), 'metrics remain at their original path')
    [ordered]@{LibrarySha256=('C'*64);CliSha256=('D'*64);InspectorSha256=('E'*64);Rows=@(
        @{Id=$id;Mode='all-word-compatible';InputSha256=$(if ($id -eq 'one') {'A'*64} else {'B'*64});
          PdfSha256=(Get-FileHash (Join-Path $run 'candidate.pdf')).Hash;MissingExpectedText=@('retained raw flag')}
    )} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $source 'outcomes.json')
    $sources += $source
}
function Join([string] $Name, [string[]] $Inputs=$sources, [string] $CaseManifest=$manifest) {
    & $JoinScript -SourceDirectory $Inputs -Manifest $CaseManifest -OutputDirectory (Join-Path $root $Name) -DocxPreset Review | Out-Null
}
function Reject([string] $Name, [string] $Expected, [string[]] $Inputs=$sources, [string] $CaseManifest=$manifest) {
    try { Join $Name $Inputs $CaseManifest }
    catch {
        if (!$_.Exception.Message.Contains($Expected)) { throw }
        if (Test-Path -LiteralPath (Join-Path $root $Name)) { throw 'Rejected join published a partial output' }
        return
    }
    throw ('Expected rejection: '+$Name)
}
function dotnet { throw 'A join must not build or convert' }
try {
    Join 'joined'
    $joined = Get-Content -LiteralPath (Join-Path $root 'joined/outcomes.json') -Raw | ConvertFrom-Json
    if ($joined.NewConversions -ne 0 -or $joined.Rows.Count -ne 2 -or $joined.SourceReports.Count -ne 2) { throw 'Joined outcomes differ' }
    foreach ($row in $joined.Rows) {
        $run = Join-Path $root ('joined/'+$row.Id+'-'+$row.Mode)
        if (Test-Path -LiteralPath (Join-Path $run 'raster')) { throw 'A raster tree was copied' }
        if (Test-Path -LiteralPath (Join-Path $run 'office-diff')) { throw 'A comparison tree was copied' }
        if (!$row.JoinedFrom.ReportSha256 -or $row.MissingExpectedText[0] -ne 'retained raw flag') { throw 'Provenance or original classification lost' }
    }
    $original = Join-Path $sources[0] 'one-all-word-compatible/candidate.pdf'
    $hash = (Get-FileHash -LiteralPath $original).Hash
    [IO.File]::WriteAllText((Join-Path $root 'joined/one-all-word-compatible/candidate.pdf'), 'changed owned copy')
    if ((Get-FileHash -LiteralPath $original).Hash -ne $hash) { throw 'Joined PDF aliases its source' }
    Reject 'missing' 'Measured outcome is missing' @($sources[0])
    $report = Join-Path $sources[1] 'outcomes.json'
    $bytes = [IO.File]::ReadAllBytes($report)
    $changed = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json
    $changed.LibrarySha256 = 'F'*64
    $changed | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $report
    Reject 'mixed-runtime' 'one frozen runtime'
    [IO.File]::WriteAllBytes($report, $bytes)
    $changed.Rows[0].InputSha256 = '0'*64
    $changed.LibrarySha256 = 'C'*64
    $changed | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $report
    Reject 'input-mismatch' 'Input identity differs'
    [IO.File]::WriteAllBytes($report, $bytes)
    $pdf = Join-Path $sources[1] 'two-all-word-compatible/candidate.pdf'
    $pdfBytes = [IO.File]::ReadAllBytes($pdf)
    [IO.File]::WriteAllText($pdf, 'bad bytes')
    Reject 'pdf-mismatch' 'Candidate identity differs'
    [IO.File]::WriteAllBytes($pdf, $pdfBytes)
    $badManifest = Join-Path $root 'bad-manifest.json'
    @(@{Id='../escape';Input='one.docx';Sha256=('A'*64)}) | ConvertTo-Json -AsArray | Set-Content -LiteralPath $badManifest
    Reject 'unsafe-id' 'Invalid or duplicate' $sources $badManifest
    $duplicate = Join-Path $root 'conflicting-source'
    Copy-Item -LiteralPath $sources[1] -Destination $duplicate -Recurse
    $duplicatePdf = Join-Path $duplicate 'two-all-word-compatible/candidate.pdf'
    [IO.File]::WriteAllText($duplicatePdf, 'different validly pinned fake PDF')
    $duplicateReport = Join-Path $duplicate 'outcomes.json'
    $duplicateData = Get-Content -LiteralPath $duplicateReport -Raw | ConvertFrom-Json
    $duplicateData.Rows[0].PdfSha256 = (Get-FileHash -LiteralPath $duplicatePdf).Hash
    $duplicateData | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $duplicateReport
    Reject 'conflicting' 'Conflicting duplicate outcome' @($sources[0],$sources[1],$duplicate)
    # An identical duplicate is reused once without repeating measurements.
    Join 'deduplicated' @($sources[0],$sources[0],$sources[1])
    $deduplicated = Get-Content -LiteralPath (Join-Path $root 'deduplicated/outcomes.json') -Raw | ConvertFrom-Json
    if ($deduplicated.Rows.Count -ne 2) { throw 'Duplicate outcome was retained twice' }
    # Final and review outcomes remain separate, while presentation modes are
    # retained independently of the DOCX preset.
    foreach ($source in $sources) {
        $dataPath = Join-Path $source 'outcomes.json'
        $data = Get-Content -LiteralPath $dataPath -Raw | ConvertFrom-Json -AsHashtable
        $final = $data.Rows[0].Clone()
        $final.Mode = 'final-preserve'
        Copy-Item -LiteralPath (Join-Path $source ($final.Id+'-all-word-compatible')) `
            -Destination (Join-Path $source ($final.Id+'-final-preserve')) -Recurse
        $data.Rows = @($data.Rows) + @($final)
        $data | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $dataPath
    }
    & $JoinScript -SourceDirectory $sources -Manifest $manifest -OutputDirectory (Join-Path $root 'both') | Out-Null
    $both = Get-Content -LiteralPath (Join-Path $root 'both/outcomes.json') -Raw | ConvertFrom-Json
    if ($both.Rows.Count -ne 4) { throw 'DOCX presets were conflated' }
    & $JoinScript -SourceDirectory $sources -Manifest $manifest -OutputDirectory (Join-Path $root 'final') -DocxPreset Final | Out-Null
    $finalOnly = Get-Content -LiteralPath (Join-Path $root 'final/outcomes.json') -Raw | ConvertFrom-Json
    if ($finalOnly.Rows.Count -ne 2 -or @($finalOnly.Rows | Where-Object Mode -NE 'final-preserve').Count) { throw 'Final selection differs' }
    $presentationManifest = Join-Path $root 'presentation-manifest.json'
    @(@{Id='one';Input='one.pptx';Sha256=('A'*64)}) | ConvertTo-Json -AsArray | Set-Content -LiteralPath $presentationManifest
    $presentationSource = Join-Path $root 'presentation-source'
    Copy-Item -LiteralPath $sources[0] -Destination $presentationSource -Recurse
    $presentationData = Get-Content -LiteralPath (Join-Path $presentationSource 'outcomes.json') -Raw | ConvertFrom-Json -AsHashtable
    $presentationData.Rows = @($presentationData.Rows[0])
    $presentationData.Rows[0].Mode = 'presentation'
    Copy-Item -LiteralPath (Join-Path $presentationSource 'one-all-word-compatible') -Destination (Join-Path $presentationSource 'one-presentation') -Recurse
    $presentationData | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $presentationSource 'outcomes.json')
    Join 'presentation' @($presentationSource) $presentationManifest
    $presentation = Get-Content -LiteralPath (Join-Path $root 'presentation/outcomes.json') -Raw | ConvertFrom-Json
    if ($presentation.Rows.Count -ne 1 -or $presentation.Rows[0].Mode -ne 'presentation') { throw 'Presentation selection differs' }
}
finally { Remove-Item Function:\dotnet }
Write-Output 'Composition outcome join ownership, identity, preflight and copy-avoidance checks passed.'
