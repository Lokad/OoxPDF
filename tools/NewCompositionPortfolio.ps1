# Fixed synthetic document portfolio for composition reliability. No Office required.
[CmdletBinding()]
param([string] $OutputDirectory = (Join-Path $PSScriptRoot '../tests/Lokad.OoxPdf.Tests/Cases/Composition'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ZipPackage.ps1')
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$cases = Join-Path $PSScriptRoot '../tests/Lokad.OoxPdf.Tests/Cases'
$w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$m = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
$packageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships'
$source = [IO.Compression.ZipFile]::OpenRead((Join-Path $cases 'docx-images.docx'))
try {
    $entry = @($source.Entries | Where-Object FullName -Like '*.png')[0]
    if (!$entry) { throw 'Public PNG source is missing' }
    $buffer = [IO.MemoryStream]::new()
    $stream = $entry.Open()
    try { $stream.CopyTo($buffer); $imageBytes = $buffer.ToArray() }
    finally { $stream.Dispose(); $buffer.Dispose() }
}
finally { $source.Dispose() }
function TextRun([string] $Text) {
    '<w:r><w:t xml:space="preserve">' + [Security.SecurityElement]::Escape($Text) + '</w:t></w:r>'
}
function Paragraph([string] $Text) { '<w:p>' + (TextRun $Text) + '</w:p>' }
$fraction = '<m:oMath><m:f><m:num><m:r><m:t>x+2</m:t></m:r></m:num><m:den><m:r><m:t>y</m:t></m:r></m:den></m:f></m:oMath>'
$literal = '<m:oMath><m:r><m:t>a+b=c</m:t></m:r></m:oMath>'
$picture = @'
<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="457200" cy="228600"/><wp:docPr id="1" name="Public trend illustration"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Public trend illustration"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
'@
$documents = @(
    @{ Id='report'; Title='Quarterly operations report'; Topic='Inventory availability'; Sections=12; HeldOut=$false },
    @{ Id='letter'; Title='Engineering review letter'; Topic='Acceptance of the revised delivery schedule'; Sections=4; HeldOut=$false },
    @{ Id='specification'; Title='Replenishment calculation specification'; Topic='Calculation inputs and acceptance criteria'; Sections=20; HeldOut=$false },
    @{ Id='technical-memo'; Title='Technical decision memorandum'; Topic='Capacity planning assumptions'; Sections=8; HeldOut=$true }
)
$manifest = @()
foreach ($document in $documents) {
    $body = [Text.StringBuilder]::new()
    $null = $body.Append('<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>' + (TextRun $document.Title) + '</w:p>')
    $null = $body.Append((Paragraph $document.Topic))
    $null = $body.Append('<w:p>' + (TextRun 'Summary: ') + '<w:bookmarkStart w:id="1" w:name="Calculation"/>' + $fraction + '<w:bookmarkEnd w:id="1"/>' + (TextRun ' governs the planning ratio. See ') + '<w:hyperlink r:id="link">' + (TextRun 'public method notes') + '</w:hyperlink>' + (TextRun '.') + '</w:p>')
    $null = $body.Append('<w:p>' + (TextRun 'The two relationships ') + $literal + (TextRun ' and ') + $fraction + (TextRun ' share the same input series.') + '</w:p>')
    $null = $body.Append('<w:p>' + (TextRun 'The trend marker ') + $picture + (TextRun ' accompanies the calculation ') + $literal + (TextRun ' without replacing it.') + '</w:p>')
    $table = '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' + (Paragraph 'Planning input') + '</w:tc><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' + (Paragraph 'Accepted relationship') + '</w:tc></w:tr><w:tr><w:tc><w:p>' + (TextRun 'Observed series') + '</w:p></w:tc><w:tc><w:p>' + $fraction + '</w:p></w:tc></w:tr></w:tbl>'
    if (!$document.HeldOut) { $null = $body.Append($table) }
    foreach ($index in 1..3) {
        $null = $body.Append('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' + (TextRun "Acceptance item ${index}: retain the visible calculation, the source illustration and the supporting explanation.") + '</w:p>')
    }
    foreach ($index in 1..$document.Sections) {
        $null = $body.Append((Paragraph "Review section ${index}. The planning team compares the observed stock position with the agreed service objective. Tables, calculations and illustrations support this decision. The same content must remain readable when it moves to another page or container."))
    }
    if ($document.HeldOut) { $null = $body.Append($table) }
    if ($document.Id -eq 'letter') { $null = $body.Append((Paragraph 'Please confirm the calculation and the agreed acceptance items. Kind regards, Public engineering team.')) }
    $null = $body.Append('<w:p><w:commentRangeStart w:id="0"/>' + (TextRun 'Review conclusion: the selected assumptions require a documented check.') + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>' + (TextRun ' Supporting note') + '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>')
    $null = $body.Append('<w:sectPr><w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>')
    $entries = @{
        '_rels/.rels' = "<Relationships xmlns=`"$packageRelationships`"><Relationship Id=`"document`" Type=`"$r/officeDocument`" Target=`"word/document.xml`"/></Relationships>"
        'word/document.xml' = "<w:document xmlns:w=`"$w`" xmlns:r=`"$r`" xmlns:m=`"$m`"><w:body>$body</w:body></w:document>"
        'word/styles.xml' = "<w:styles xmlns:w=`"$w`"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=`"Calibri`" w:hAnsi=`"Calibri`"/><w:sz w:val=`"22`"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=`"120`" w:line=`"240`" w:lineRule=`"auto`"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=`"paragraph`" w:styleId=`"Title`"><w:rPr><w:b/><w:sz w:val=`"36`"/></w:rPr></w:style></w:styles>"
        'word/settings.xml' = "<w:settings xmlns:w=`"$w`" xmlns:m=`"$m`"><m:mathPr><m:mathFont m:val=`"Cambria Math`"/></m:mathPr></w:settings>"
        'word/header1.xml' = "<w:hdr xmlns:w=`"$w`" xmlns:m=`"$m`"><w:p>$(TextRun $document.Title)</w:p><w:p>$literal</w:p></w:hdr>"
        'word/footer1.xml' = "<w:ftr xmlns:w=`"$w`"><w:p>$(TextRun 'Public synthetic composition portfolio')</w:p></w:ftr>"
        'word/comments.xml' = "<w:comments xmlns:w=`"$w`"><w:comment w:id=`"0`" w:author=`"Public reviewer`" w:date=`"2026-10-15T10:00:00Z`"><w:p>$(TextRun 'Check the calculation alongside the illustration.')</w:p></w:comment></w:comments>"
        'word/footnotes.xml' = "<w:footnotes xmlns:w=`"$w`"><w:footnote w:type=`"separator`" w:id=`"-1`"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id=`"1`"><w:p>$(TextRun 'All quantities in this synthetic example are illustrative.')</w:p></w:footnote></w:footnotes>"
        'word/numbering.xml' = "<w:numbering xmlns:w=`"$w`"><w:abstractNum w:abstractNumId=`"0`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"decimal`"/><w:lvlText w:val=`"%1.`"/><w:pPr><w:ind w:left=`"360`" w:hanging=`"360`"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=`"1`"><w:abstractNumId w:val=`"0`"/></w:num></w:numbering>"
    }
    $relationships = '<Relationships xmlns="' + $packageRelationships + '">'
    $types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
    foreach ($part in @('document','styles','settings','numbering','comments','footnotes','header1','footer1')) {
        $kind = if ($part -eq 'header1') {'header'} elseif ($part -eq 'footer1') {'footer'} else {$part}
        $mime = if ($part -eq 'document') {'document.main'} else {$kind}
        $types += "<Override PartName=`"/word/$part.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.wordprocessingml.$mime+xml`"/>"
        if ($part -ne 'document') { $relationships += "<Relationship Id=`"$kind`" Type=`"$r/$kind`" Target=`"$part.xml`"/>" }
    }
    $entries['[Content_Types].xml'] = $types + '</Types>'
    $entries['word/_rels/document.xml.rels'] = $relationships + "<Relationship Id=`"image`" Type=`"$r/image`" Target=`"media/trend.png`"/><Relationship Id=`"link`" Type=`"$r/hyperlink`" Target=`"https://example.org/method`" TargetMode=`"External`"/></Relationships>"
    $path = Join-Path $OutputDirectory ($document.Id + '.docx')
    New-ZipPackage -Path $path -Entries $entries
    $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Update)
    try { $stream = $zip.CreateEntry('word/media/trend.png').Open(); try { $stream.Write($imageBytes) } finally { $stream.Dispose() } }
    finally { $zip.Dispose() }
    $manifest += [ordered]@{Id=$document.Id;Input=('tests/Lokad.OoxPdf.Tests/Cases/Composition/'+$document.Id+'.docx');Sha256=(Get-FileHash $path).Hash;HeldOut=$document.HeldOut;ExpectedEquationCount=6;ExpectedInlineImageCount=1;ExpectedText=@($document.Title,$document.Topic,"Review section $($document.Sections).",'All quantities in this synthetic example are illustrative.');Features=@('tables','numbering','image','equations','bookmark','link','header','footer','footnote','comment')}
}
foreach ($deck in @(@{Id='briefing-deck';Name='pptx-ladder-11-composite-chart-port.pptx';HeldOut=$false},@{Id='held-out-deck';Name='pptx-ladder-10-composite-table-port.pptx';HeldOut=$true})) {
    $path = Join-Path $cases $deck.Name
    $manifest += [ordered]@{Id=$deck.Id;Input=('tests/Lokad.OoxPdf.Tests/Cases/'+$deck.Name);Sha256=(Get-FileHash $path).Hash;HeldOut=$deck.HeldOut;Features=@('text','shapes','table-or-chart')}
}
ConvertTo-Json -InputObject $manifest -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'portfolio.json') -Encoding utf8
Write-Output 'Four fixed DOCX composites generated; two existing PPTX composites pinned. Six documents total.'
