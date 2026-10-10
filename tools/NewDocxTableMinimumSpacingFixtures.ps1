# Generates public plain table-grid minimum-height probes from the note-spacing seed.
param([string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repo 'tests/Lokad.OoxPdf.Tests/Cases'
if (!$OutputDirectory) { $OutputDirectory = $cases }
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
. (Join-Path $PSScriptRoot 'ZipPackage.ps1')
$source = @{}
$archive = [IO.Compression.ZipFile]::OpenRead((Join-Path $cases 'note-table-minimum-line-spacing-footnote.docx'))
try {
    foreach ($entry in $archive.Entries) {
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $source[$entry.FullName] = $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
} finally { $archive.Dispose() }
$word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach ($spec in @(@('row', 1), @('grid', 2))) {
    $entries = @{} + $source
    [xml] $document = $entries['word/document.xml']
    $ns = [Xml.XmlNamespaceManager]::new($document.NameTable)
    $ns.AddNamespace('w', $word)
    foreach ($reference in @($document.SelectNodes('//w:footnoteReference|//w:endnoteReference', $ns))) {
        [void] $reference.ParentNode.RemoveChild($reference)
    }
    $table = $document.SelectSingleNode('/w:document/w:body/w:tbl', $ns)
    $template = $table.SelectSingleNode('w:tr/w:tc', $ns).CloneNode($true)
    foreach ($row in @($table.SelectNodes('w:tr', $ns))) { [void] $table.RemoveChild($row) }
    $grid = $table.SelectSingleNode('w:tblGrid', $ns)
    $grid.RemoveAll()
    foreach ($unused in 0..1) {
        $column = $document.CreateElement('w', 'gridCol', $word)
        [void] $column.SetAttribute('w', $word, '4680')
        [void] $grid.AppendChild($column)
    }
    $borders = $document.CreateElement('w', 'tblBorders', $word)
    foreach ($edge in 'top', 'left', 'bottom', 'right', 'insideH', 'insideV') {
        $border = $document.CreateElement('w', $edge, $word)
        foreach ($attribute in @(@('val', 'single'), @('sz', '8'), @('color', '000000'))) {
            [void] $border.SetAttribute($attribute[0], $word, $attribute[1])
        }
        [void] $borders.AppendChild($border)
    }
    [void] $table.SelectSingleNode('w:tblPr', $ns).AppendChild($borders)
    for ($ri = 1; $ri -le $spec[1]; $ri++) {
        $row = $document.CreateElement('w', 'tr', $word)
        [void] $table.AppendChild($row)
        for ($ci = 1; $ci -le 2; $ci++) {
            $cell = $template.CloneNode($true)
            [void] $row.AppendChild($cell)
            [void] $cell.SelectSingleNode('w:tcPr/w:tcW', $ns).SetAttribute('w', $word, '4680')
            $paragraphs = @($cell.SelectNodes('w:p', $ns))
            foreach ($index in 0..1) {
                $paragraph = $paragraphs[$index]
                $style = $paragraph.SelectSingleNode('w:r/w:rPr', $ns).CloneNode($true)
                foreach ($run in @($paragraph.SelectNodes('w:r', $ns))) { [void] $paragraph.RemoveChild($run) }
                $fonts = $style.SelectSingleNode('w:rFonts', $ns)
                if (!$fonts) { $fonts = $document.CreateElement('w', 'rFonts', $word); [void] $style.AppendChild($fonts) }
                [void] $fonts.SetAttribute('ascii', $word, 'Calibri')
                [void] $fonts.SetAttribute('hAnsi', $word, 'Calibri')
                $size = $style.SelectSingleNode('w:sz', $ns)
                if (!$size) { $size = $document.CreateElement('w', 'sz', $word); [void] $style.AppendChild($size) }
                [void] $size.SetAttribute('val', $word, '24')
                $run = $document.CreateElement('w', 'r', $word)
                [void] $run.AppendChild($style)
                $text = $document.CreateElement('w', 't', $word)
                $text.InnerText = if ($index -eq 0) { "R${ri}C${ci} public preceding." } else { "R${ri}C${ci} public grid minimum paragraph followed by public words " + ('minimum line ' * 15) }
                [void] $run.AppendChild($text)
                [void] $paragraph.AppendChild($run)
            }
        }
    }
    $entries['word/document.xml'] = $document.OuterXml
    New-ZipPackage -Path (Join-Path $OutputDirectory ('table-minimum-line-spacing-' + $spec[0] + '.docx')) -Entries $entries
}
