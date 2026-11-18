# Shared byte-preserving PDF text reader and conservative page count (T04).

function Read-PdfLatin1Text([string] $PdfPath) {
    return [Text.Encoding]::Latin1.GetString([IO.File]::ReadAllBytes($PdfPath))
}

function Get-PdfPageCount([string] $PdfText) {
    return [regex]::Matches($PdfText, '/Type\s*/Page\b').Count
}
