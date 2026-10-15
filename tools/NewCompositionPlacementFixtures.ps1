# Development probes separate from the fixed six-document portfolio.
[CmdletBinding()]
param([string] $OutputDirectory = (Join-Path $PSScriptRoot '../tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'))
$ErrorActionPreference = 'Stop'
$cases = Join-Path $PSScriptRoot '../tests/Lokad.OoxPdf.Tests/Cases'
$source = Join-Path $cases 'Composition/report.docx'
$pin = @(Get-Content (Join-Path $cases 'Composition/portfolio.json') -Raw | ConvertFrom-Json | Where-Object Id -EQ report)[0]
if ((Get-FileHash $source).Hash -ne $pin.Sha256) { throw 'Pinned portfolio source differs' }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$m = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
$r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
function Read-XmlPart($Zip, [string] $Name) {
    $reader = [IO.StreamReader]::new($Zip.GetEntry($Name).Open())
    try { [xml] $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Write-XmlPart($Zip, [string] $Name, [xml] $Document) {
    if ($Zip.GetEntry($Name)) { $Zip.GetEntry($Name).Delete() }
    $writer = [IO.StreamWriter]::new($Zip.CreateEntry($Name).Open(), [Text.UTF8Encoding]::new($false))
    try { $Document.Save($writer) } finally { $writer.Dispose() }
}
$manifestPath = Join-Path $OutputDirectory 'probes.json'
$rows = if (Test-Path -LiteralPath $manifestPath) { @(Get-Content $manifestPath -Raw | ConvertFrom-Json) } else { @() }
foreach ($placement in @('body','cell','header','footer','body-review','cell-review','header-review','footer-review')) {
    $owner = $placement -replace '-review$',''
    $review = $placement.EndsWith('-review',[StringComparison]::Ordinal)
    $path = Join-Path $OutputDirectory ($placement+'.docx')
    if (Test-Path -LiteralPath $path) {
        $existing = @($rows | Where-Object Id -EQ $placement)
        if ($existing.Count -ne 1 -or $existing[0].SourceSha256 -ne $pin.Sha256 -or
            (Get-FileHash $path).Hash -ne $existing[0].Sha256) { throw ('Existing probe differs: '+$placement) }
        continue
    }
    Copy-Item -LiteralPath $source -Destination $path -Force
    $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $main = Read-XmlPart $zip 'word/document.xml'
        $ns = [Xml.XmlNamespaceManager]::new($main.NameTable)
        $ns.AddNamespace('w',$w); $ns.AddNamespace('m',$m)
        $paragraph = $main.CreateElement('w','p',$w)
        $start = $main.CreateElement('w','bookmarkStart',$w)
        $null = $start.SetAttribute('id',$w,'51'); $null = $start.SetAttribute('name',$w,'MixedObjects')
        $null = $paragraph.AppendChild($start)
        $equations = $main.SelectNodes('//m:oMath',$ns)
        $null = $paragraph.AppendChild($equations[0].CloneNode($true))
        $null = $paragraph.AppendChild($main.SelectSingleNode('//w:drawing',$ns).ParentNode.CloneNode($true))
        $null = $paragraph.AppendChild($equations[1].CloneNode($true))
        $end = $main.CreateElement('w','bookmarkEnd',$w); $null = $end.SetAttribute('id',$w,'51')
        $null = $paragraph.AppendChild($end)
        $section = $main.SelectSingleNode('//w:sectPr',$ns).CloneNode($true)
        $body = $main.SelectSingleNode('/w:document/w:body',$ns)
        $body.RemoveAll()
        if ($owner -in @('header','footer')) {
            $tag = if ($owner -eq 'header') {'hdr'} else {'ftr'}
            [xml]$story = "<w:$tag xmlns:w=`"$w`" xmlns:m=`"$m`" xmlns:r=`"$r`"/>"
            $null = $story.DocumentElement.AppendChild($story.ImportNode($paragraph,$true))
            Write-XmlPart $zip ('word/'+$owner+'1.xml') $story
            [xml]$relationships = "<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"image`" Type=`"$r/image`" Target=`"media/trend.png`"/></Relationships>"
            Write-XmlPart $zip ('word/_rels/'+$owner+'1.xml.rels') $relationships
            [xml]$prose = "<w:p xmlns:w=`"$w`"><w:r><w:t>Container source prose.</w:t></w:r></w:p>"
            if ($review) {
                $prose.DocumentElement.InnerXml = '<w:commentRangeStart w:id="0"/><w:r><w:t>Container source prose.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
            }
            $null = $body.AppendChild($main.ImportNode($prose.DocumentElement,$true))
        }
        elseif ($owner -eq 'cell') {
            [xml]$table = "<w:tbl xmlns:w=`"$w`"><w:tblPr><w:tblW w:w=`"9000`" w:type=`"dxa`"/></w:tblPr><w:tblGrid><w:gridCol w:w=`"9000`"/></w:tblGrid><w:tr><w:tc/></w:tr></w:tbl>"
            $tableNs = [Xml.XmlNamespaceManager]::new($table.NameTable); $tableNs.AddNamespace('w',$w)
            $null = $table.SelectSingleNode('//w:tc',$tableNs).AppendChild($table.ImportNode($paragraph,$true))
            $null = $body.AppendChild($main.ImportNode($table.DocumentElement,$true))
        }
        else { $null = $body.AppendChild($paragraph) }
        if ($review -and $owner -in @('body','cell')) {
            [xml]$anchor = "<w:p xmlns:w=`"$w`"><w:commentRangeStart w:id=`"0`"/><w:r><w:t>Review anchor.</w:t></w:r><w:commentRangeEnd w:id=`"0`"/><w:r><w:commentReference w:id=`"0`"/></w:r></w:p>"
            $null = $body.AppendChild($main.ImportNode($anchor.DocumentElement,$true))
        }
        $null = $body.AppendChild($section)
        Write-XmlPart $zip 'word/document.xml' $main
    }
    finally { $zip.Dispose() }
    $rows += [ordered]@{Id=$placement;Input=('tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/'+$placement+'.docx');Sha256=(Get-FileHash $path).Hash;SourceSha256=$pin.Sha256;MixedParagraphEquations=2;MixedParagraphImages=1;SurroundingProse=$false}
}
ConvertTo-Json -InputObject $rows -Depth 4 | Set-Content (Join-Path $OutputDirectory 'probes.json') -Encoding utf8
Write-Output 'Eight placement probes retained, including four review compositions; fixed portfolio unchanged.'
