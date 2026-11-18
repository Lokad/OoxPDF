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
