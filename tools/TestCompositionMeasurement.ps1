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
function dotnet {
    $calls.Add($args)
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
Write-Output 'Late missing/hash-mismatched inputs prevent all conversions and output publication; valid inputs reach one sentinel conversion. No .NET processes or library builds.'
