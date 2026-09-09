# Shared PDF information readers (T04): Get-PdfPageCount counts page objects
# via a conservative page-type pattern. Dot-source this file instead of
# copying the function into comparison and check tools.

function Get-PdfPageCount([string] $PdfPath) {
    $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
    $text = [System.Text.Encoding]::Latin1.GetString($bytes)
    return ([regex]::Matches($text, '/Type\s*/Page\b') | Where-Object { $_.Value -notmatch '/Pages' }).Count
}
