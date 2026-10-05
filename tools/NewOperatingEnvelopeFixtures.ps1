# Deterministic public DOCX inputs for allocation/peak probes: one Latin glyph
# and an explicit page-break-before paragraph per page, with Letter page size.
param(
    [ValidateRange(1, 10000)]
    [int[]] $PageCounts = @(400, 800, 1600),
    [string] $OutputDirectory = "artifacts/operating-envelope/inputs"
)
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$outputFull = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path $repoRoot $OutputDirectory }))
New-Item -ItemType Directory -Force -Path $outputFull | Out-Null
$utf8 = [Text.UTF8Encoding]::new($false)
$timestamp = [DateTimeOffset]::new(2000, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
function Write-ZipText($Archive, [string] $Name, [string] $Text) {
    $entry = $Archive.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
    $entry.LastWriteTime = $timestamp
    $entryStream = $entry.Open()
    try { $bytes = $utf8.GetBytes($Text); $entryStream.Write($bytes) }
    finally { $entryStream.Dispose() }
}
$manifest = foreach ($pageCount in ($PageCounts | Sort-Object -Unique)) {
    $document = [Text.StringBuilder]::new()
    [void]$document.Append('<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>')
    for ($index = 0; $index -lt $pageCount; $index++) {
        [void]$document.Append('<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>')
    }
    [void]$document.Append('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>')
    $path = Join-Path $outputFull "breaks-$pageCount.docx"
    $stream = [IO.File]::Open($path, [IO.FileMode]::Create, [IO.FileAccess]::Write)
    try {
        $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            Write-ZipText $archive '[Content_Types].xml' '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
            Write-ZipText $archive '_rels/.rels' '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
            Write-ZipText $archive 'word/document.xml' $document.ToString()
        }
        finally { $archive.Dispose() }
    }
    finally { $stream.Dispose() }
    [ordered]@{ File = [IO.Path]::GetFileName($path); Pages = $pageCount; ByteSize = (Get-Item -LiteralPath $path).Length; Sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() }
}
@($manifest) | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $outputFull 'manifest.json') -Encoding utf8
Write-Host "Wrote $(@($manifest).Count) operating-envelope fixtures to $outputFull."
