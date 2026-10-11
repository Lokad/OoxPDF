# Public mixed source-threshold Calibri controls; package payloads reproduce exactly.
param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if (!$OutputDirectory) { $OutputDirectory = $cases }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$specs = @(
    @{ Boundary='run-break'; Position='early'; Form='run'; Values=@(0,24); Name='note-wrapped-calibri-mixed-run-early' },
    @{ Boundary='paragraph-break'; Position='later'; Form='paragraph'; Values=@(1,22); Name='note-wrapped-calibri-mixed-met-later' }
)
function Set-KerningThreshold($run, $value, $xml, $ns) {
    $rp = $run.SelectSingleNode('w:rPr', $ns)
    $kern = $rp.SelectSingleNode('w:kern', $ns)
    if (!$kern) { $kern = $xml.CreateElement('w','kern',$word); [void]$rp.AppendChild($kern) }
    [void]$kern.SetAttribute('val', $word, [string]$value)
}
foreach ($spec in $specs) {
    $entries = @{}
    $archive = [IO.Compression.ZipFile]::OpenRead((Join-Path $cases ('note-each-page-'+$spec.Boundary+'.docx')))
    try {
        foreach ($entry in $archive.Entries) {
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $entries[$entry.FullName] = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
    } finally { $archive.Dispose() }
    foreach ($part in @('word/document.xml','word/footnotes.xml')) {
        [xml]$xml = $entries[$part]
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('w',$word)
        foreach ($rp in $xml.SelectNodes('//w:rPr',$ns)) {
            $fonts = $rp.SelectSingleNode('w:rFonts',$ns)
            if (!$fonts) { $fonts = $xml.CreateElement('w','rFonts',$word); [void]$rp.PrependChild($fonts) }
            foreach ($attribute in @('ascii','hAnsi','eastAsia','cs')) { [void]$fonts.SetAttribute($attribute,$word,'Calibri') }
        }
        $entries[$part] = $xml.OuterXml
    }
    [xml]$body = $entries['word/document.xml']
    $ns = [Xml.XmlNamespaceManager]::new($body.NameTable); $ns.AddNamespace('w',$word)
    [void]$body.SelectSingleNode('/w:document/w:body/w:sectPr/w:footnotePr/w:numFmt',$ns).SetAttribute('val',$word,'decimalZero')
    $filler = 'Public wrapped words retain authored digits 108. ' * 9
    $paragraphIndex = 0
    foreach ($p in $body.SelectNodes('/w:document/w:body/w:p[w:r/w:footnoteReference]',$ns)) {
        $texts = @($p.SelectNodes('.//w:t',$ns)); $tag = $texts[0].InnerText.Split(' ')[0]
        $texts[0].InnerText = $tag+' '+$(if ($spec.Position -eq 'later') { $filler } else { '' })+'before note 105 '
        $texts[-1].InnerText = ' and after note 108. '+$(if ($spec.Position -eq 'early') { $filler } else { '' })
        $authored = $p.InnerText; $counter = 0
        foreach ($run in @($p.SelectNodes('w:r',$ns))) {
            $text = $run.SelectSingleNode('w:t',$ns)
            if ($spec.Form -eq 'paragraph' -or !$text -or !$text.InnerText) {
                Set-KerningThreshold $run $spec.Values[$paragraphIndex % 2] $body $ns
                continue
            }
            $original = $text.InnerText
            $spaces = @(for ($i=1; $i -lt $original.Length-1; $i++) { if ($original[$i] -eq ' ') { $i+1 } })
            $cut = $spaces | Sort-Object { [Math]::Abs($_ - $original.Length/2.0) } | Select-Object -First 1
            if (!$cut) { throw 'A baseline text run must have a reproducible word boundary.' }
            foreach ($chunk in @($original.Substring(0,$cut), $original.Substring($cut))) {
                $clone = $run.CloneNode($true)
                $clone.SelectSingleNode('w:t',$ns).InnerText = $chunk
                Set-KerningThreshold $clone $spec.Values[$counter % 2] $body $ns
                $counter++
                [void]$p.InsertBefore($clone,$run)
            }
            [void]$p.RemoveChild($run)
        }
        if ($p.InnerText -cne $authored) { throw 'Splitting runs changed authored text.' }
        $paragraphIndex++
    }
    $entries['word/document.xml'] = $body.OuterXml
    New-ZipPackage -Path (Join-Path $OutputDirectory ($spec.Name+'.docx')) -Entries $entries
}
