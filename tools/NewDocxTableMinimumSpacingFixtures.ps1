# Generates public plain table-grid minimum-height probes from the note-spacing seed.
param([string] $OutputDirectory, [switch] $ParagraphPositionProbes, [switch] $DeclaredRowHeightProbes, [switch] $MixedParagraphFontProbes, [switch] $AlignedCellSpacingProbes, [switch] $AlignedCellMarginProbes, [switch] $CellContentClipProbes, [switch] $AutomaticCellSpacingProbes)
$ErrorActionPreference = 'Stop'
if (@($ParagraphPositionProbes, $DeclaredRowHeightProbes, $MixedParagraphFontProbes, $AlignedCellSpacingProbes, $AlignedCellMarginProbes, $CellContentClipProbes, $AutomaticCellSpacingProbes | Where-Object { $_ }).Count -gt 1) { throw 'Choose one probe family' }
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
$marginProbes = $AlignedCellMarginProbes -or $CellContentClipProbes -or $AutomaticCellSpacingProbes
$alignedProbes = $AlignedCellSpacingProbes -or $marginProbes
$positionProbes = $ParagraphPositionProbes -or $DeclaredRowHeightProbes -or $MixedParagraphFontProbes -or $alignedProbes
$specs = if ($CellContentClipProbes -or $AutomaticCellSpacingProbes) { @(@('center', 1, 1), @('bottom', 1, 1)) } elseif ($alignedProbes) { @(@('center', 1, 1), @('bottom', 2, 2)) } elseif ($MixedParagraphFontProbes) { @(@('family', 2, 2), @('size', 1, 1)) } elseif ($DeclaredRowHeightProbes) { @(@('exact', 1, 1), @('minimum', 2, 2)) } elseif ($ParagraphPositionProbes) { @(@('first', 1, 1), @('middle', 2, 2), @('single', 2, 2)) } else { @(@('row', 1, 2), @('grid', 2, 2)) }
foreach ($spec in $specs) {
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
    for ($unused = 0; $unused -lt $spec[2]; $unused++) {
        $column = $document.CreateElement('w', 'gridCol', $word)
        [void] $column.SetAttribute('w', $word, [string](9360 / $spec[2]))
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
        if (($DeclaredRowHeightProbes -or $alignedProbes) -and !$AutomaticCellSpacingProbes) {
            $properties = $document.CreateElement('w', 'trPr', $word)
            $height = $document.CreateElement('w', 'trHeight', $word)
            [void] $height.SetAttribute('val', $word, $(if ($CellContentClipProbes) { '960' } elseif ($alignedProbes) { '3600' } else { '2880' }))
            [void] $height.SetAttribute('hRule', $word, $(if ($alignedProbes -or $spec[0] -eq 'exact') { 'exact' } else { 'atLeast' }))
            [void] $properties.AppendChild($height)
            [void] $row.AppendChild($properties)
        }
        for ($ci = 1; $ci -le $spec[2]; $ci++) {
            $cell = $template.CloneNode($true)
            [void] $row.AppendChild($cell)
            [void] $cell.SelectSingleNode('w:tcPr/w:tcW', $ns).SetAttribute('w', $word, [string](9360 / $spec[2]))
            if ($alignedProbes) {
                $vertical = $document.CreateElement('w', 'vAlign', $word)
                [void] $vertical.SetAttribute('val', $word, $spec[0])
                [void] $cell.SelectSingleNode('w:tcPr', $ns).AppendChild($vertical)
            }
            if ($marginProbes) {
                $properties = $cell.SelectSingleNode('w:tcPr', $ns)
                foreach ($existing in @($properties.SelectNodes('w:tcMar', $ns))) { [void] $properties.RemoveChild($existing) }
                $margins = $document.CreateElement('w', 'tcMar', $word)
                foreach ($edge in @(@('top', '120'), @('bottom', '240'))) {
                    $margin = $document.CreateElement('w', $edge[0], $word)
                    [void] $margin.SetAttribute('w', $word, $edge[1])
                    [void] $margin.SetAttribute('type', $word, 'dxa')
                    [void] $margins.AppendChild($margin)
                }
                [void] $properties.AppendChild($margins)
            }
            $paragraphs = @($cell.SelectNodes('w:p', $ns))
            if ($positionProbes) {
                foreach ($paragraph in $paragraphs) { [void] $cell.RemoveChild($paragraph) }
                $sequence = if ($spec[0] -in @('first', 'exact')) { @(@('atLeast', $true), @('auto', $false)) } elseif ($spec[0] -in @('single', 'minimum')) { @(@('auto', $true), @('atLeast', $false), @('auto', $false)) } else { @(@('auto', $true), @('atLeast', $true), @('auto', $false)) }
                if ($MixedParagraphFontProbes) { $sequence = @(@('auto', $false), @('atLeast', ($spec[0] -eq 'family')), @('auto', $false)) }
                if ($alignedProbes) { $sequence = @(@('auto', $false), @('atLeast', $true), @('auto', $false)) }
                if ($CellContentClipProbes) { $sequence = ,@('atLeast', $true) }
                $paragraphs = @()
                for ($index = 0; $index -lt $sequence.Count; $index++) {
                    $paragraph = $template.SelectSingleNode('w:p[last()]', $ns).CloneNode($true)
                    $properties = $paragraph.SelectSingleNode('w:pPr', $ns)
                    $spacing = $properties.SelectSingleNode('w:spacing', $ns)
                    [void] $spacing.SetAttribute('before', $word, $(if ($index -eq 0) { '0' } else { [string](120 * (1 + $index % 2)) }))
                    [void] $spacing.SetAttribute('after', $word, '120')
                    [void] $spacing.SetAttribute('lineRule', $word, $sequence[$index][0])
                    [void] $spacing.SetAttribute('line', $word, $(if ($sequence[$index][0] -eq 'auto') { '240' } else { '480' }))
                    $contextual = $properties.SelectSingleNode('w:contextualSpacing', $ns)
                    if (!$contextual) { $contextual = $document.CreateElement('w', 'contextualSpacing', $word); [void] $properties.AppendChild($contextual) }
                    [void] $contextual.SetAttribute('val', $word, '0')
                    [void] $cell.AppendChild($paragraph)
                    $paragraphs += $paragraph
                }
            }
            for ($index = 0; $index -lt $paragraphs.Count; $index++) {
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
                if ($positionProbes) {
                    $prefix = "R${ri}C${ci}P$($index + 1)"
                    $text.InnerText = if ($sequence[$index][1]) { $prefix + ' public minimum paragraph ' + ('minimum line ' * $(if ($spec[2] -eq 1) { 9 } else { 3 })) } else { $prefix + ' short public paragraph.' }
                } else {
                    $text.InnerText = if ($index -eq 0) { "R${ri}C${ci} public preceding." } else { "R${ri}C${ci} public grid minimum paragraph followed by public words " + ('minimum line ' * 15) }
                }
                [void] $run.AppendChild($text)
                [void] $paragraph.AppendChild($run)
                if ($MixedParagraphFontProbes -or $CellContentClipProbes -or $AutomaticCellSpacingProbes) {
                    foreach ($properties in @($paragraph.SelectNodes('.//w:rPr', $ns))) {
                        $fonts = $properties.SelectSingleNode('w:rFonts', $ns)
                        if (!$fonts) { $fonts = $document.CreateElement('w', 'rFonts', $word); [void] $properties.AppendChild($fonts) }
                        $family = if ($CellContentClipProbes -or $AutomaticCellSpacingProbes -or $index -eq 1) { 'Arial' } else { 'Calibri' }
                        [void] $fonts.SetAttribute('ascii', $word, $family)
                        [void] $fonts.SetAttribute('hAnsi', $word, $family)
                        $halfPoints = if ($spec[0] -eq 'size') { @(20, 32, 24)[$index] } else { 24 }
                        foreach ($name in 'sz', 'szCs') {
                            $size = $properties.SelectSingleNode('w:' + $name, $ns)
                            if (!$size) { $size = $document.CreateElement('w', $name, $word); [void] $properties.AppendChild($size) }
                            [void] $size.SetAttribute('val', $word, [string]$halfPoints)
                        }
                    }
                }
            }
        }
    }
    $entries['word/document.xml'] = $document.OuterXml
    $name = if ($AutomaticCellSpacingProbes) { 'table-automatic-margin-minimum-line-spacing-' + $spec[0] + '.docx' } elseif ($CellContentClipProbes) { 'table-margin-content-clip-' + $spec[0] + '.docx' } elseif ($AlignedCellMarginProbes) { 'table-aligned-margin-minimum-line-spacing-' + $spec[0] + '.docx' } elseif ($alignedProbes) { 'table-aligned-minimum-line-spacing-' + $spec[0] + '.docx' } elseif ($MixedParagraphFontProbes) { 'table-mixed-minimum-line-spacing-' + $spec[0] + '.docx' } elseif ($DeclaredRowHeightProbes) { 'table-declared-minimum-line-spacing-' + $spec[0] + '.docx' } elseif ($ParagraphPositionProbes) { 'table-paragraph-minimum-line-spacing-' + $spec[0] + '.docx' } else { 'table-minimum-line-spacing-' + $spec[0] + '.docx' }
    New-ZipPackage -Path (Join-Path $OutputDirectory $name) -Entries $entries
}
