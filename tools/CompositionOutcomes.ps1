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
        Pictures = if ($override -and $override.Contains('ExpectedInlineImageCount')) { $override.ExpectedInlineImageCount } else { $Document.ExpectedInlineImageCount }
        FieldResults = $Document.ExpectedFieldResults
    }
}

function Test-PortfolioFieldExpectations($Actual, $Expected) {
    if ($null -eq $Actual -or $null -eq $Expected) { return $null -eq $Actual -and $null -eq $Expected }
    $actualJson = [Text.Json.Nodes.JsonNode]::Parse((ConvertTo-Json -InputObject $Actual -Depth 4 -Compress))
    $expectedJson = [Text.Json.Nodes.JsonNode]::Parse((ConvertTo-Json -InputObject $Expected -Depth 4 -Compress))
    [Text.Json.Nodes.JsonNode]::DeepEquals($actualJson, $expectedJson)
}

function Get-MissingPortfolioTextPhrases($Operations, $ExpectedText, [string] $Mode) {
    # Slide labels and table cells can paint in an order unrelated to their baselines.
    Get-MissingPdfTextPhrases $Operations $ExpectedText -SeparateBaselines:($Mode -ceq 'presentation')
}

function Read-PortfolioTextOperations([string] $Path, [string] $Id) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $operations = ConvertFrom-Json -NoEnumerate -InputObject ([Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF))
    if ($operations -isnot [Array]) { throw ('Invalid portfolio text operations: ' + $Id) }
    [pscustomobject]@{Operations=$operations;Hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))}
}

function Get-LogicalPdfPageCount([string[]] $Lines, [string] $Source) {
    $markers = @($Lines | Where-Object { $_.StartsWith('Logical pages:', [StringComparison]::Ordinal) })
    [int] $count = 0
    if ($markers.Count -ne 1 -or $markers[0] -cnotmatch '^Logical pages: ([0-9]+)$' -or
        ![int]::TryParse($Matches[1], [ref]$count) -or $count -lt 1) {
        throw ('Missing or invalid logical PDF page count: ' + $Source)
    }
    return $count
}

function Read-PortfolioPageInspection([string] $Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = [Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
    [pscustomobject]@{
        PageCount = Get-LogicalPdfPageCount ($text -split '\r?\n') $Path
        Hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    }
}

function Get-VerifiedPortfolioPageInspection([string] $Path, $Row) {
    if ($null -eq $Row.InspectionLogSha256) { return }
    $inspection = Read-PortfolioPageInspection $Path
    if ($inspection.Hash -cne $Row.InspectionLogSha256 -or
        ($Row.InspectedPages -isnot [int] -and $Row.InspectedPages -isnot [long]) -or
        $inspection.PageCount -ne $Row.InspectedPages) {
        throw ('Candidate page inspection differs: ' + $Path)
    }
    return $inspection
}

function Get-PortfolioIgnoredEquationCount($Diagnostics) {
    @($Diagnostics | Where-Object { $_.Id -eq 'DOCX_UNSUPPORTED_EQUATION' }).Count
}

function Read-PortfolioContentAudit([string] $Path, [hashtable] $Loaded) {
    $resolved = (Resolve-Path -LiteralPath $Path).ProviderPath
    if ($Loaded.ContainsKey($resolved)) { return $Loaded[$resolved] }
    $bytes = [IO.File]::ReadAllBytes($resolved)
    $report = ConvertFrom-Json -AsHashtable -InputObject ([Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF))
    if ($report.Schema -cne 'ooxpdf-portfolio-content-audit/1') { throw 'Unsupported portfolio content audit' }
    foreach ($hash in @($report.ToolSha256,$report.LibrarySha256,$report.InspectorSha256,$report.ManifestSha256)) {
        if ($hash -notmatch '^[A-F0-9]{64}$') { throw 'Invalid portfolio content audit identity' }
    }
    if (!$report.Rows) { throw 'Portfolio content audit has no outcomes' }
    $rows = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    foreach ($row in $report.Rows) {
        if (!$rows.TryAdd($row.Id + '/' + $row.Mode, $row)) { throw 'Duplicate portfolio content audit outcome' }
    }
    $audit = [pscustomobject]@{
        Path=$resolved; Sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
        Report=$report; Rows=$rows
    }
    $Loaded[$resolved] = $audit
    return $audit
}

function Get-PortfolioContentWitnesses($Row, $Outcomes, $Audit, [hashtable] $Loaded, [string] $ReferenceSha256) {
    $equations = if ([object]::Equals($Row.ExpectedSourceEquations, 0L)) { 'No expectation' } else { 'Unverified' }
    $pictures = if ([object]::Equals($Row.ExpectedSourcePictures, 0L)) { 'No expectation' } else { 'Unverified' }
    $result = [ordered]@{Equations=$equations; Pictures=$pictures}
    if ($Row.ExpectedFieldResults) { $result.FieldResults = 'Unverified' }
    if (!$Audit -and $Row.ContentWitnesses.Audit) {
        $pin = $Row.ContentWitnesses.Audit
        $Audit = Read-PortfolioContentAudit $pin.Path $Loaded
        if ($Audit.Sha256 -cne $pin.Sha256) { throw 'Retained portfolio content audit changed' }
    }
    if (!$Audit) { return $result }

    $key = $Row.Id + '/' + $Row.Mode
    if (!$Audit.Rows.ContainsKey($key)) { throw ('Portfolio content audit outcome missing: ' + $key) }
    $observed = $Audit.Rows[$key]
    if ($Audit.Report.LibrarySha256 -cne $Outcomes.LibrarySha256 -or
        $Audit.Report.InspectorSha256 -cne $Outcomes.InspectorSha256 -or
        $observed.InputSha256 -cne $Row.InputSha256 -or $observed.CandidatePdfSha256 -cne $Row.PdfSha256 -or
        !$Row.TextOperationsSha256 -or $observed.CandidateOperationsSha256 -cne $Row.TextOperationsSha256 -or
        $observed.ExpectedSourceEquations -cne $Row.ExpectedSourceEquations -or
        $observed.ExpectedSourcePictures -cne $Row.ExpectedSourcePictures -or
        !(Test-PortfolioFieldExpectations $observed.FieldResults.Expected $Row.ExpectedFieldResults) -or
        $observed.IgnoredEquationDiagnostics -cne $Row.IgnoredEquationDiagnostics -or
        $observed.ReferencePdfSha256 -notmatch '^[A-F0-9]{64}$' -or
        ($ReferenceSha256 -and $observed.ReferencePdfSha256 -cne $ReferenceSha256)) {
        throw ('Portfolio content audit identity differs: ' + $key)
    }
    $result.Audit = [ordered]@{Path=$Audit.Path; Sha256=$Audit.Sha256; ToolSha256=$Audit.Report.ToolSha256;
        ReferencePdfSha256=$observed.ReferencePdfSha256}
    if ($Row.ExpectedSourceEquations -gt 0 -and $observed.ExpectedSimpleMath -is [Collections.IDictionary] -and
        $observed.ExpectedSimpleMath.get_Count() -gt 0 -and
        $observed.MissingSimpleMath -is [Collections.IDictionary] -and
        $observed.UnresolvedSourceFeatures -is [array]) {
        if (!$observed.UnresolvedSourceFeatures.Count) {
            $result.Equations = if ($observed.MissingSimpleMath.get_Count()) { 'Missing' } else { 'Witnesses present' }
        } elseif ($observed.MissingReferenceConfirmedSimpleMath -is [Collections.IDictionary] -and
            $observed.MissingReferenceConfirmedSimpleMath.get_Count()) {
            $result.Equations = 'Missing'
        }
    }
    if ($Row.ExpectedSourcePictures -gt 0 -and $observed.PictureInk.ExpectedSourcePictures -ceq $Row.ExpectedSourcePictures) {
        $result.Pictures = switch -CaseSensitive ($observed.PictureInk.Status) {
            'Missing picture ink' { 'Missing' }
            'Picture witnesses present' { 'Witnesses present' }
            default { 'Unverified' }
        }
    }
    if ($Row.ExpectedFieldResults -and $observed.FieldResults.Missing -is [array] -and
        $observed.FieldResults.Unverified -is [array]) {
        if ($observed.FieldResults.Missing.Count) { $result.FieldResults = 'Missing' }
        elseif (!$observed.FieldResults.Unverified.Count) { $result.FieldResults = 'Witnesses present' }
    }
    return $result
}

function Get-PortfolioContentSeverity($MissingExpectedText, $IgnoredEquationDiagnostics, $Witnesses) {
    $hasTextCheck = $MissingExpectedText -is [array] -and
        !@($MissingExpectedText | Where-Object { $_ -isnot [string] }).Count
    $hasEquationCheck = ($IgnoredEquationDiagnostics -is [int] -or $IgnoredEquationDiagnostics -is [long]) -and
        $IgnoredEquationDiagnostics -ge 0
    if (($hasTextCheck -and $MissingExpectedText.Count) -or ($hasEquationCheck -and $IgnoredEquationDiagnostics) -or
        $Witnesses.Equations -ceq 'Missing' -or $Witnesses.Pictures -ceq 'Missing' -or $Witnesses.FieldResults -ceq 'Missing') {
        return 'Content blocker'
    }
    if ($hasTextCheck -and $hasEquationCheck -and
        $Witnesses.Equations -cin @('No expectation','Witnesses present') -and
        $Witnesses.Pictures -cin @('No expectation','Witnesses present') -and
        $Witnesses.FieldResults -cne 'Unverified') { return 'Needs geometry and visual review' }
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
