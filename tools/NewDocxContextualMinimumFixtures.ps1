# Author contextual minimum-leading controls from the ordinary and note sources.
param(
    [Parameter(Mandatory)] [string] $OutputDirectory,
    [switch] $ElevenPoint
)

$ErrorActionPreference = 'Stop'
$cases = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/Lokad.OoxPdf.Tests/Cases'
. (Join-Path $PSScriptRoot 'ZipPackage.ps1')

$word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$variant = if ($ElevenPoint) { 'eleven-point' } else { 'contextual' }
$specifications = @(
    @{ Name = "ordinary-$variant-minimum.docx"; Source = 'ordinary-minimum-line-20.docx'; Alternating = $false },
    @{ Name = "note-$variant-minimum-alternating.docx"; Source = 'note-marked-minimum-line-20.docx'; Alternating = $true }
)

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
foreach ($specification in $specifications) {
    $destination = Join-Path $OutputDirectory $specification.Name
    if (Test-Path -LiteralPath $destination) {
        throw "Fixture already exists at $destination. Choose a fresh output directory."
    }
}

$styleIds = @('PublicContextualA', 'PublicContextualB')
foreach ($specification in $specifications) {
    $entries = @{}
    $archive = [IO.Compression.ZipFile]::OpenRead((Join-Path $cases $specification.Source))
    try {
        foreach ($entry in $archive.Entries) {
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $entries[$entry.FullName] = $reader.ReadToEnd() }
            finally { $reader.Dispose() }
        }
    } finally { $archive.Dispose() }

    [xml] $document = $entries['word/document.xml']
    $namespaces = [Xml.XmlNamespaceManager]::new($document.NameTable)
    $namespaces.AddNamespace('w', $word)
    $paragraphs = @($document.SelectNodes('/w:document/w:body/w:p', $namespaces))
    for ($index = 0; $index -lt $paragraphs.Count; $index++) {
        $paragraph = $paragraphs[$index]
        if ($ElevenPoint) {
            foreach ($run in $paragraph.SelectNodes('w:r', $namespaces)) {
                $properties = $run.SelectSingleNode('w:rPr', $namespaces)
                if (!$properties) {
                    $properties = $document.CreateElement('w', 'rPr', $word)
                    [void] $run.PrependChild($properties)
                }
                $size = $properties.SelectSingleNode('w:sz', $namespaces)
                if (!$size) {
                    $size = $document.CreateElement('w', 'sz', $word)
                    [void] $properties.AppendChild($size)
                }
                [void] $size.SetAttribute('val', $word, '22')
            }
        }

        $properties = $paragraph.SelectSingleNode('w:pPr', $namespaces)
        $spacing = $properties.SelectSingleNode('w:spacing', $namespaces)
        [void] $spacing.SetAttribute('before', $word, '80')
        [void] $spacing.SetAttribute('after', $word, '80')
        [void] $properties.AppendChild($document.CreateElement('w', 'contextualSpacing', $word))
        if ($specification.Alternating) {
            $style = $document.CreateElement('w', 'pStyle', $word)
            [void] $style.SetAttribute('val', $word, $styleIds[$index % $styleIds.Count])
            [void] $properties.PrependChild($style)
        }
    }

    if ($specification.Alternating) {
        [xml] $styles = $entries['word/styles.xml']
        foreach ($id in $styleIds) {
            $style = $styles.CreateElement('w', 'style', $word)
            [void] $style.SetAttribute('type', $word, 'paragraph')
            [void] $style.SetAttribute('styleId', $word, $id)
            $name = $styles.CreateElement('w', 'name', $word)
            [void] $name.SetAttribute('val', $word, $id)
            [void] $style.AppendChild($name)
            [void] $styles.DocumentElement.AppendChild($style)
        }
        $entries['word/styles.xml'] = $styles.OuterXml
    }

    $entries['word/document.xml'] = $document.OuterXml
    New-ZipPackage -Path (Join-Path $OutputDirectory $specification.Name) -Entries $entries
}
