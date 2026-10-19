# Prove a late invalid input prevents all conversion work, without running .NET.
[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [string] $MeasurementScript = (Join-Path $PSScriptRoot 'MeasureCompositionPortfolio.ps1')
)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) {
    $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/measurement-test-' + [Guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo = Join-Path $OutputDirectory 'repository'
$tools = Join-Path $repo 'tools'
New-Item -ItemType Directory -Path $tools -Force | Out-Null
$script = Join-Path $tools 'MeasureCompositionPortfolio.ps1'
Copy-Item -LiteralPath $MeasurementScript -Destination $script
$cli = Join-Path $repo 'candidate.dll'
$inspector = Join-Path $repo 'inspector.dll'
foreach ($file in @($cli, $inspector, (Join-Path $repo 'Lokad.OoxPdf.dll'))) {
    'Sentinel runtime; never execute' | Set-Content -LiteralPath $file
}
$calls = [Collections.Generic.List[object]]::new()
$inspectionOnly = $false
function dotnet {
    $calls.Add($args)
    if ($inspectionOnly) {
        if ($args[0] -ne $inspector) { throw 'Reinspection reached the converter' }
        New-Item -ItemType Directory -Path $args[2] | Out-Null
        @(@{PageNumber=1;Font='Body';DecodedText='Retained source.'}) | ConvertTo-Json -AsArray |
            Set-Content -LiteralPath (Join-Path $args[2] 'text-operations.json')
        $global:LASTEXITCODE = 0
        Write-Output '1 0 obj: page'
        return
    }
    throw 'Sentinel conversion reached'
}
$first = Join-Path $repo 'first.docx'
$second = Join-Path $repo 'second.docx'
'First input' | Set-Content -LiteralPath $first
$firstHash = (Get-FileHash -LiteralPath $first).Hash
foreach ($fault in @('missing', 'hash', 'valid')) {
    $case = Join-Path $OutputDirectory $fault
    New-Item -ItemType Directory -Path $case | Out-Null
    $manifest = Join-Path $case 'manifest.json'
    if ($fault -ne 'missing') { 'Second input' | Set-Content -LiteralPath $second }
    $secondHash = if ($fault -eq 'valid') { (Get-FileHash -LiteralPath $second).Hash } else { '0' * 64 }
    @(@{Id='first';Input='first.docx';Sha256=$firstHash},
      @{Id='second';Input='second.docx';Sha256=$secondHash}) |
        ConvertTo-Json | Set-Content -LiteralPath $manifest
    $output = Join-Path $case 'candidate'
    $calls.Clear()
    $failure = $null
    try {
        & $script -CliPath $cli -InspectorPath $inspector -Manifest $manifest -OutputDirectory $output
    }
    catch { $failure = $_ }
    if ($fault -eq 'valid') {
        if ($calls.Count -ne 1 -or "$failure" -notlike '*Sentinel conversion reached*') {
            throw 'Valid inputs must reach the converter after preflight'
        }
    }
    else {
        if (!$failure -or $calls.Count -ne 0 -or (Test-Path -LiteralPath $output)) {
            throw ('Late invalid input started conversion or published output: ' + $fault)
        }
        if ($fault -eq 'hash' -and "$failure" -notlike '*Portfolio input differs: second*') {
            throw 'Hash mismatch must identify its actual input'
        }
    }
}
$inspectionOnly = $true
foreach ($fault in @('runtime', 'input', 'pdf', 'missing', 'duplicate', 'valid')) {
    $case = Join-Path $OutputDirectory ('reuse-' + $fault)
    $prior = Join-Path $case 'prior'
    New-Item -ItemType Directory -Path $prior -Force | Out-Null
    $manifest = Join-Path $case 'manifest.json'
    $documents = @(@{Id='first';Input='first.docx';Sha256=$firstHash;ExpectedText=@('Retained source.');ExpectedEquationCount=0},
        @{Id='second';Input='second.docx';Sha256=(Get-FileHash -LiteralPath $second).Hash;ExpectedText=@('Retained source.');ExpectedEquationCount=0})
    $documents | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifest
    $priorRows = @(foreach ($document in $documents) {
        foreach ($mode in @('final-preserve','all-word-compatible')) {
            $key = $document.Id + '-' + $mode
            $run = Join-Path $prior $key
            New-Item -ItemType Directory -Path $run | Out-Null
            ('Pinned PDF ' + $key) | Set-Content -LiteralPath (Join-Path $run 'candidate.pdf')
            '[]' | Set-Content -LiteralPath (Join-Path $run 'diagnostics.json')
            if ($fault -ne 'missing' -or $key -ne 'second-all-word-compatible') {
                'Retained original conversion log' | Set-Content -LiteralPath (Join-Path $run 'convert.log')
            }
            @{Id=$document.Id;Mode=$mode;InputSha256=$document.Sha256;
                PdfSha256=(Get-FileHash -LiteralPath (Join-Path $run 'candidate.pdf')).Hash;ColdProcessElapsedMs=123d}
        }
    })
    if ($fault -eq 'input') { $priorRows[-1].InputSha256 = '0' * 64 }
    if ($fault -eq 'pdf') { $priorRows[-1].PdfSha256 = '0' * 64 }
    if ($fault -eq 'duplicate') { $priorRows += $priorRows[0] }
    @{LibrarySha256=$(if ($fault -eq 'runtime') {'0' * 64} else {(Get-FileHash -LiteralPath (Join-Path $repo 'Lokad.OoxPdf.dll')).Hash});
        CliSha256=(Get-FileHash -LiteralPath $cli).Hash;Rows=$priorRows} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $prior 'outcomes.json')
    $priorHash = (Get-FileHash -LiteralPath (Join-Path $prior 'outcomes.json')).Hash
    $calls.Clear()
    $output = Join-Path $case 'candidate'
    $failure = $null
    try { & $script -CliPath $cli -InspectorPath $inspector -Manifest $manifest -OutputDirectory $output -ReuseCandidateDirectory $prior | Out-Null }
    catch { $failure = $_ }
    if ($fault -ne 'valid') {
        if (!$failure -or $calls.Count -ne 0 -or (Test-Path -LiteralPath $output)) {
            throw ('Invalid reuse started work or published output: ' + $fault)
        }
        continue
    }
    if ($failure) { throw $failure }
    $report = Get-Content -LiteralPath (Join-Path $output 'outcomes.json') -Raw | ConvertFrom-Json
    if ($calls.Count -ne 4 -or $report.Conversions -ne 0 -or $report.ReusedCandidates -ne 4 -or
        $report.ReusedOutcomesSha256 -ne $priorHash -or
        (Get-FileHash -LiteralPath (Join-Path $prior 'outcomes.json')).Hash -ne $priorHash) {
        throw 'Successful reinspection must preserve all candidates and original provenance without conversion'
    }
    foreach ($row in $report.Rows) {
        if ($row.ColdProcessElapsedMs -ne 123d -or $row.MissingExpectedText.Count -ne 0 -or
            $row.PdfSha256 -ne (Get-FileHash -LiteralPath (Join-Path $prior ($row.Id + '-' + $row.Mode + '/candidate.pdf'))).Hash) {
            throw 'Reinspection changed candidate identity/timing or lost content'
        }
    }
}
Write-Output 'Measurement/reinspection preflight, retained PDF/provenance and zero-conversion reuse pass; five invalid reuse inputs reject before publication. No .NET processes or library builds.'
