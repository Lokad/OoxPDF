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
