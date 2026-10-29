# Decoded whitespace has advance, but no glyph ink. Filter after merging runs
# when measuring line geometry, so spaces between words retain their advance.
# Missing/null decoded text is unknown and must remain in comparisons.
function Select-NonWhitespacePdfTextOperations($Operations) {
    foreach ($operation in @($Operations)) {
        if ($null -eq $operation) { continue }
        $decoded = $operation.PSObject.Properties['DecodedText']
        if ($null -eq $decoded -or $null -eq $decoded.Value -or -not [string]::IsNullOrWhiteSpace([string]$decoded.Value)) {
            $operation
        }
    }
}

# Preserve emitted separators and operation order within each page. A font,
# kerning or punctuation split does not add a source space. Page boundaries do
# separate text; this is a presence screen, not a reading-order reconstruction.
function Join-DecodedPdfTextPages($Operations) {
    $pages = @($Operations | Group-Object { $_.PageNumber } | Sort-Object { [int]$_.Name })
    ($pages | ForEach-Object { ($_.Group | ForEach-Object { $_.DecodedText }) -join '' }) -join "`n"
}
