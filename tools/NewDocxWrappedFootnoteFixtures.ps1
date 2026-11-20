# Public wrapped-body advance and footnote numbering controls.
param([Parameter(Mandatory)] [string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$cases = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/Lokad.OoxPdf.Tests/Cases'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
. (Join-Path $PSScriptRoot 'ZipPackage.ps1')

$word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$filler = 'Public wrapped words retain authored digits 108. ' * 9
foreach ($spec in @(
    @{ Name = 'calibri-zero-early'; Font = 'Calibri'; Source = 'run-break'; FillerBefore = $false; NumberFormat = 'decimalZero'; Kerning = '0' }
    @{ Name = 'calibri-unmet-later'; Font = 'Calibri'; Source = 'paragraph-break'; FillerBefore = $true; NumberFormat = 'decimalZero'; Kerning = '26' }
    @{ Name = 'calibri-met-one-early'; Font = 'Calibri'; Source = 'run-break'; FillerBefore = $false; NumberFormat = 'decimal'; Kerning = '1' }
    @{ Name = 'calibri-met-limit-later'; Font = 'Calibri'; Source = 'paragraph-break'; FillerBefore = $true; NumberFormat = 'decimalZero'; Kerning = '24' }
    @{ Name = 'arial-later-numbering'; Font = 'Arial'; Source = 'paragraph-break'; FillerBefore = $true; NumberFormat = 'decimalZero' }
    @{ Name = 'times-early-numbering'; Font = 'Times New Roman'; Source = 'run-break'; FillerBefore = $false; NumberFormat = 'decimal' }
)) {
    $entries = @{}
    $archive = [IO.Compression.ZipFile]::OpenRead((Join-Path $cases "note-each-page-$($spec.Source).docx"))
    try {
        foreach ($entry in $archive.Entries) {
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $entries[$entry.FullName] = $reader.ReadToEnd() }
            finally { $reader.Dispose() }
        }
    }
    finally { $archive.Dispose() }

    [xml] $document = $entries['word/document.xml']
    $ns = [Xml.XmlNamespaceManager]::new($document.NameTable)
    $ns.AddNamespace('w', $word)
    $beforeNote = if ($spec.FillerBefore) { $filler } else { '' }
    $afterNote = if ($spec.FillerBefore) { '' } else { $filler }
    foreach ($paragraph in $document.SelectNodes('/w:document/w:body/w:p[w:r/w:footnoteReference]', $ns)) {
        $texts = @($paragraph.SelectNodes('.//w:t', $ns))
        $tag = $texts[0].InnerText.Split(' ')[0]
        $texts[0].InnerText = "$tag $beforeNote" + 'before note 105 '
        $texts[-1].InnerText = " and after note 108. $afterNote"
    }
    $format = $document.SelectSingleNode('/w:document/w:body/w:sectPr/w:footnotePr/w:numFmt', $ns)
    [void] $format.SetAttribute('val', $word, $spec.NumberFormat)
    $entries['word/document.xml'] = $document.OuterXml

    foreach ($partName in @('word/document.xml', 'word/footnotes.xml')) {
        [xml] $part = $entries[$partName]
        $ns = [Xml.XmlNamespaceManager]::new($part.NameTable)
        $ns.AddNamespace('w', $word)
        foreach ($properties in $part.SelectNodes('//w:rPr', $ns)) {
            $fonts = $properties.SelectSingleNode('w:rFonts', $ns)
            if (!$fonts) {
                $fonts = $part.CreateElement('w', 'rFonts', $word)
                [void] $properties.PrependChild($fonts)
            }
            foreach ($attribute in @('ascii', 'hAnsi', 'eastAsia', 'cs')) {
                [void] $fonts.SetAttribute($attribute, $word, $spec.Font)
            }
            if ($partName -eq 'word/document.xml' -and $null -ne $spec.Kerning) {
                $kern = $properties.SelectSingleNode('w:kern', $ns)
                if (!$kern) {
                    $kern = $part.CreateElement('w', 'kern', $word)
                    [void] $properties.AppendChild($kern)
                }
                [void] $kern.SetAttribute('val', $word, $spec.Kerning)
            }
        }
        $entries[$partName] = $part.OuterXml
    }
    New-ZipPackage -Path (Join-Path $OutputDirectory "note-wrapped-$($spec.Name).docx") -Entries $entries
}
