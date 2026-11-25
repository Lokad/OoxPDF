# Public wrapped-note controls for authored paragraph spacing.
param([string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if (!$OutputDirectory) { $OutputDirectory = $cases }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)

$word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$seeds = @(
    @{ Name = 'note-wrapped-calibri-mixed-run-early'; Variants = @(
        @{ Name = 'note-wrapped-calibri-after-four-early'; After = 80 }
        @{ Name = 'note-wrapped-calibri-before-four-early'; Before = 80; After = 0 }
    ) }
    @{ Name = 'note-wrapped-calibri-mixed-met-later'; Variants = @(
        @{ Name = 'note-wrapped-calibri-after-eight-later'; After = 160 }
        @{ Name = 'note-wrapped-calibri-both-eight-later'; Before = 160; After = 160 }
    ) }
)
foreach ($seed in $seeds) {
    $entries = @{}
    $archive = [IO.Compression.ZipFile]::OpenRead((Join-Path $cases ($seed.Name + '.docx')))
    try {
        foreach ($entry in $archive.Entries) {
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $entries[$entry.FullName] = $reader.ReadToEnd() }
            finally { $reader.Dispose() }
        }
    }
    finally { $archive.Dispose() }

    $sourceBody = $entries['word/document.xml']
    foreach ($variant in $seed.Variants) {
        [xml] $body = $sourceBody
        $ns = [Xml.XmlNamespaceManager]::new($body.NameTable)
        $ns.AddNamespace('w', $word)
        foreach ($paragraph in $body.SelectNodes('/w:document/w:body/w:p[w:r/w:footnoteReference]', $ns)) {
            $spacing = $paragraph.SelectSingleNode('w:pPr/w:spacing', $ns)
            if ($variant.ContainsKey('Before')) {
                [void] $spacing.SetAttribute('before', $word, [string] $variant.Before)
            }
            [void] $spacing.SetAttribute('after', $word, [string] $variant.After)
        }
        $entries['word/document.xml'] = $body.OuterXml
        New-ZipPackage -Path (Join-Path $OutputDirectory ($variant.Name + '.docx')) -Entries $entries
    }
}
