# Generates public two-section footnote/endnote numbering probes.
param([string] $OutputDirectory, [switch] $PaginationProbes, [switch] $CustomMarkProbes, [switch] $HitAreaSpacingProbes, [switch] $BeforeSpacingProbes, [switch] $MultilineBeforeSpacingProbes, [switch] $ContextualSpacingProbes, [switch] $MultilineContextualBeforeSpacingProbes, [switch] $CurrentContextualMultilineBeforeSpacingProbes)

$ErrorActionPreference = 'Stop'
if (@(@($PaginationProbes, $CustomMarkProbes, $HitAreaSpacingProbes, $BeforeSpacingProbes, $MultilineBeforeSpacingProbes, $ContextualSpacingProbes, $MultilineContextualBeforeSpacingProbes, $CurrentContextualMultilineBeforeSpacingProbes) | Where-Object { $_ }).Count -gt 1) { throw 'Select one probe family.' }
$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot 'tests/Lokad.OoxPdf.Tests/Cases'
if (!$OutputDirectory) { $OutputDirectory = $cases }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
. (Join-Path $PSScriptRoot 'ZipPackage.ps1')

$seed = Join-Path $cases 'note-nav-numbered-ids.docx'
$archive = [IO.Compression.ZipFile]::OpenRead($seed)
$source = @{}
try {
    foreach ($entry in $archive.Entries) {
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $source[$entry.FullName] = $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
}
finally { $archive.Dispose() }

$word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$kinds = if ($PaginationProbes) { @('footnote') } else { @('footnote', 'endnote') }
foreach ($kind in $kinds) {
    $entries = @{} + $source
    [xml] $document = $entries['word/document.xml']
    $ns = [Xml.XmlNamespaceManager]::new($document.NameTable)
    $ns.AddNamespace('w', $word)
    $body = $document.SelectSingleNode('/w:document/w:body', $ns)
    $sectionTemplate = $body.SelectSingleNode('w:sectPr', $ns).CloneNode($true)
    $paragraphTemplate = $body.SelectSingleNode('w:p', $ns).CloneNode($true)
    $body.RemoveAll()
    [xml] $notes = $entries['word/footnotes.xml']
    $noteNs = [Xml.XmlNamespaceManager]::new($notes.NameTable)
    $noteNs.AddNamespace('w', $word)
    $noteTemplate = $notes.SelectSingleNode('/w:footnotes/w:footnote[@w:id="37"]', $noteNs).CloneNode($true)
    foreach ($note in @($notes.SelectNodes('/w:footnotes/w:footnote[not(@w:type)]', $noteNs))) {
        [void] $notes.DocumentElement.RemoveChild($note)
    }
    $labels = @('Alpha', 'Bravo', 'Charlie', 'Delta')
    $ids = @(37, 4, 73, 8)
    foreach ($sectionIndex in 0..1) {
        $section = $sectionTemplate.CloneNode($true)
        foreach ($node in @($section.SelectNodes('w:footnotePr|w:endnotePr|w:type', $ns))) {
            [void] $section.RemoveChild($node)
        }
        $format = @('lowerRoman', 'decimalZero')[$sectionIndex]
        $start = @(4, 9)[$sectionIndex]
        $properties = $document.CreateElement('w', ($kind + 'Pr'), $word)
        $position = if ($kind -eq 'footnote') { 'pageBottom' } else { 'docEnd' }
        $properties.InnerXml = "<w:pos xmlns:w='$word' w:val='$position'/><w:numFmt xmlns:w='$word' w:val='$format'/><w:numStart xmlns:w='$word' w:val='$start'/><w:numRestart xmlns:w='$word' w:val='eachSect'/>"
        [void] $section.PrependChild($properties)
        $type = $document.CreateElement('w', 'type', $word)
        $breakType = if ($PaginationProbes) { 'continuous' } else { 'nextPage' }
        [void] $type.SetAttribute('val', $word, $breakType)
        [void] $section.AppendChild($type)
        foreach ($within in 0..1) {
            $index = 2 * $sectionIndex + $within
            $paragraph = $paragraphTemplate.CloneNode($true)
            $texts = @($paragraph.SelectNodes('.//w:t', $ns))
            $texts[0].InnerText = $labels[$index] + ' public body before note '
            $texts[-1].InnerText = ' and after note.'
            [void] $paragraph.SelectSingleNode('.//w:footnoteReference', $ns).SetAttribute('id', $word, [string]$ids[$index])
            [void] $body.AppendChild($paragraph)
            $note = $noteTemplate.CloneNode($true)
            [void] $note.SetAttribute('id', $word, [string]$ids[$index])
            $note.SelectSingleNode('.//w:t', $noteNs).InnerText = ' Public destination ' + $labels[$index] + '.'
            [void] $notes.DocumentElement.AppendChild($note)
        }
        if ($sectionIndex -eq 0) {
            [void] $paragraph.SelectSingleNode('w:pPr', $ns).AppendChild($section)
        }
        else { [void] $body.AppendChild($section) }
    }
    $entries['word/footnotes.xml'] = $notes.OuterXml
    $entries['word/document.xml'] = $document.OuterXml
    if ($kind -eq 'endnote') {
        $entries['word/document.xml'] = $entries['word/document.xml'].Replace('footnoteReference', 'endnoteReference').Replace('FootnoteReference', 'EndnoteReference')
        $entries['word/endnotes.xml'] = $entries['word/footnotes.xml'].Replace('footnote', 'endnote').Replace('Footnote', 'Endnote')
        $entries.Remove('word/footnotes.xml')
        foreach ($part in @('[Content_Types].xml', 'word/_rels/document.xml.rels')) {
            $entries[$part] = $entries[$part].Replace('footnotes', 'endnotes')
        }
    }
    $entries['word/settings.xml'] = "<w:settings xmlns:w='$word'><w:${kind}Pr><w:numFmt w:val='decimal'/><w:numStart w:val='105'/></w:${kind}Pr></w:settings>"
    if ($CustomMarkProbes -or (($HitAreaSpacingProbes -or $BeforeSpacingProbes -or $MultilineBeforeSpacingProbes -or $ContextualSpacingProbes -or $MultilineContextualBeforeSpacingProbes -or $CurrentContextualMultilineBeforeSpacingProbes) -and $kind -eq 'footnote')) {
        [xml] $customDocument = $entries['word/document.xml']
        $customNs = [Xml.XmlNamespaceManager]::new($customDocument.NameTable)
        $customNs.AddNamespace('w', $word)
        $reference = $customDocument.SelectSingleNode("//w:${kind}Reference", $customNs)
        [void] $reference.SetAttribute('customMarkFollows', $word, '1')
        $mark = $customDocument.CreateElement('w', 't', $word)
        $mark.InnerText = '*'
        [void] $reference.ParentNode.InsertAfter($mark, $reference)
        [xml] $customNotes = $entries["word/${kind}s.xml"]
        $customNoteNs = [Xml.XmlNamespaceManager]::new($customNotes.NameTable)
        $customNoteNs.AddNamespace('w', $word)
        $noteMark = $customNotes.SelectSingleNode("//w:$kind[@w:id='37']//w:${kind}Ref", $customNoteNs)
        $noteText = $customNotes.CreateElement('w', 't', $word)
        $noteText.InnerText = '*'
        [void] $noteMark.ParentNode.ReplaceChild($noteText, $noteMark)
        $entries['word/document.xml'] = $customDocument.OuterXml
        $entries["word/${kind}s.xml"] = $customNotes.OuterXml
    }
    if ($HitAreaSpacingProbes -or $BeforeSpacingProbes -or $MultilineBeforeSpacingProbes -or $ContextualSpacingProbes -or $MultilineContextualBeforeSpacingProbes -or $CurrentContextualMultilineBeforeSpacingProbes) {
        [xml] $spacingDocument = $entries['word/document.xml']
        $spacingNs = [Xml.XmlNamespaceManager]::new($spacingDocument.NameTable)
        $spacingNs.AddNamespace('w', $word)
        $paragraphProperties = $spacingDocument.SelectSingleNode('/w:document/w:body/w:p/w:pPr', $spacingNs)
        $spacing = $paragraphProperties.SelectSingleNode('w:spacing', $spacingNs)
        if (!$spacing) { $spacing = $spacingDocument.CreateElement('w', 'spacing', $word); [void] $paragraphProperties.AppendChild($spacing) }
        [void] $spacing.SetAttribute('after', $word, '480')
        if ($BeforeSpacingProbes -or $MultilineBeforeSpacingProbes -or $ContextualSpacingProbes -or $MultilineContextualBeforeSpacingProbes -or $CurrentContextualMultilineBeforeSpacingProbes) { [void] $spacing.SetAttribute('before', $word, '240') }
        if ($CurrentContextualMultilineBeforeSpacingProbes) { [void] $paragraphProperties.AppendChild($spacingDocument.CreateElement('w', 'contextualSpacing', $word)) }
        if ($ContextualSpacingProbes -or $MultilineContextualBeforeSpacingProbes) {
            if ($ContextualSpacingProbes) { [void] $paragraphProperties.AppendChild($spacingDocument.CreateElement('w', 'contextualSpacing', $word)) }
            $previous = $paragraphProperties.ParentNode.CloneNode($true)
            foreach ($run in @($previous.SelectNodes('w:r', $spacingNs))) { [void] $previous.RemoveChild($run) }
            $previousProperties = $previous.SelectSingleNode('w:pPr', $spacingNs)
            if ($MultilineContextualBeforeSpacingProbes) { [void] $previousProperties.AppendChild($spacingDocument.CreateElement('w', 'contextualSpacing', $word)) }
            $previousSpacing = $previousProperties.SelectSingleNode('w:spacing', $spacingNs)
            [void] $previousSpacing.SetAttribute('before', $word, '0')
            [void] $previousSpacing.SetAttribute('after', $word, '120')
            $run = $spacingDocument.CreateElement('w', 'r', $word)
            [void] $run.AppendChild($paragraphProperties.ParentNode.SelectSingleNode('w:r/w:rPr', $spacingNs).CloneNode($true))
            $text = $spacingDocument.CreateElement('w', 't', $word)
            $text.InnerText = 'Public preceding paragraph.'
            [void] $run.AppendChild($text)
            [void] $previous.AppendChild($run)
            [void] $paragraphProperties.ParentNode.ParentNode.InsertBefore($previous, $paragraphProperties.ParentNode)
        }
        if ($MultilineBeforeSpacingProbes -or $MultilineContextualBeforeSpacingProbes -or $CurrentContextualMultilineBeforeSpacingProbes) {
            $paragraph = $paragraphProperties.ParentNode
            @($paragraph.SelectNodes('.//w:t', $spacingNs))[-1].InnerText = ' and public following text ' + ('after note ' * 35)
        }
        $entries['word/document.xml'] = $spacingDocument.OuterXml
    }
    $name = if ($CurrentContextualMultilineBeforeSpacingProbes) { "note-current-contextual-multiline-before-spacing-$kind.docx" } elseif ($MultilineContextualBeforeSpacingProbes) { "note-multiline-contextual-before-spacing-$kind.docx" } elseif ($ContextualSpacingProbes) { "note-contextual-spacing-$kind.docx" } elseif ($MultilineBeforeSpacingProbes) { "note-multiline-before-spacing-$kind.docx" } elseif ($BeforeSpacingProbes) { "note-before-spacing-$kind.docx" } elseif ($HitAreaSpacingProbes) { "note-hit-area-$kind.docx" } elseif ($CustomMarkProbes) { "note-custom-mark-$kind.docx" } elseif ($PaginationProbes) { 'note-continuous-footnote.docx' } else { "note-sections-$kind.docx" }
    New-ZipPackage -Path (Join-Path $OutputDirectory $name) -Entries $entries
    if ($PaginationProbes) {
        [xml] $boundaryDocument = $entries['word/document.xml']
        $boundaryNs = [Xml.XmlNamespaceManager]::new($boundaryDocument.NameTable)
        $boundaryNs.AddNamespace('w', $word)
        foreach ($reference in @($boundaryDocument.SelectNodes('//w:footnoteReference', $boundaryNs))) {
            [void] $reference.ParentNode.ParentNode.RemoveChild($reference.ParentNode)
        }
        $boundaries = @($boundaryDocument.SelectNodes('//w:sectPr', $boundaryNs))
        [void] $boundaries[0].SelectSingleNode('w:type', $boundaryNs).SetAttribute('val', $word, 'nextPage')
        [xml] $emptyNotes = $entries['word/footnotes.xml']
        $emptyNs = [Xml.XmlNamespaceManager]::new($emptyNotes.NameTable)
        $emptyNs.AddNamespace('w', $word)
        foreach ($note in @($emptyNotes.SelectNodes('/w:footnotes/w:footnote[not(@w:type)]', $emptyNs))) {
            [void] $emptyNotes.DocumentElement.RemoveChild($note)
        }
        $entries['word/document.xml'] = $boundaryDocument.OuterXml
        $entries['word/footnotes.xml'] = $emptyNotes.OuterXml
        New-ZipPackage -Path (Join-Path $OutputDirectory 'section-following-start-type.docx') -Entries $entries
    }
}
