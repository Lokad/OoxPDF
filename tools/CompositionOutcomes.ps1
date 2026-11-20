. (Join-Path $PSScriptRoot 'PdfTextGeometry.ps1')

function Get-PortfolioModes($Document, [string] $Preset) {
    if (!$Document.Input.EndsWith('.docx', [StringComparison]::OrdinalIgnoreCase)) { return 'presentation' }
    if ($Preset -eq 'Final') { return 'final-preserve' }
    if ($Preset -eq 'Review') { return 'all-word-compatible' }
    return @('final-preserve','all-word-compatible')
}

function Get-PortfolioContentExpectations($Document, [string] $Mode) {
    $override = if ($Document.ExpectedByMode) { $Document.ExpectedByMode[$Mode] }
    [pscustomobject]@{
        Text = if ($override -and $override.Contains('ExpectedText')) { $override.ExpectedText } else { $Document.ExpectedText }
        Equations = if ($override -and $override.Contains('ExpectedEquationCount')) { $override.ExpectedEquationCount } else { $Document.ExpectedEquationCount }
    }
}

function Read-PortfolioTextOperations([string] $Path, [string] $Id) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $operations = ConvertFrom-Json -NoEnumerate -InputObject ([Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF))
    if ($operations -isnot [Array]) { throw ('Invalid portfolio text operations: ' + $Id) }
    [pscustomobject]@{Operations=$operations;Hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))}
}

function Get-PortfolioIgnoredEquationCount($Diagnostics) {
    @($Diagnostics | Where-Object { $_.Id -eq 'DOCX_UNSUPPORTED_EQUATION' }).Count
}

function Get-PortfolioContentSeverity($MissingExpectedText, $IgnoredEquationDiagnostics) {
    $hasTextCheck = $MissingExpectedText -is [array] -and
        !@($MissingExpectedText | Where-Object { $_ -isnot [string] }).Count
    $hasEquationCheck = ($IgnoredEquationDiagnostics -is [int] -or $IgnoredEquationDiagnostics -is [long]) -and
        $IgnoredEquationDiagnostics -ge 0
    if (($hasTextCheck -and $MissingExpectedText.Count) -or ($hasEquationCheck -and $IgnoredEquationDiagnostics)) {
        return 'Content blocker'
    }
    if ($hasTextCheck -and $hasEquationCheck) { return 'Needs geometry and visual review' }
    return 'Content evidence incomplete'
}

function Get-VerifiedDiagnosticsHash([string] $Path, $Diagnostics, [string] $Id) {
    if ($Diagnostics -isnot [Array]) { throw ('Retained diagnostics are missing or not an array: ' + $Id) }
    # Compare semantic data and pin the same sidecar bytes. Formatting and property order can differ.
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = [Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
    try {
        $actual = [Text.Json.Nodes.JsonNode]::Parse($text)
        $expected = [Text.Json.Nodes.JsonNode]::Parse((ConvertTo-Json -InputObject $Diagnostics -Depth 12 -Compress))
    }
    catch { throw ('Reused diagnostics are invalid: ' + $Id) }
    if ($actual -isnot [Text.Json.Nodes.JsonArray] -or ![Text.Json.Nodes.JsonNode]::DeepEquals($actual, $expected)) {
        throw ('Reused diagnostics differ from the retained outcome: ' + $Id)
    }
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}
