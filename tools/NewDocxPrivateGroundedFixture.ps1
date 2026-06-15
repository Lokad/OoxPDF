$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null

function New-ZipPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [hashtable] $Entries
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($name in $Entries.Keys) {
                $entry = $archive.CreateEntry($name)
                $entryStream = $entry.Open()
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Entries[$name])
                    $entryStream.Write($bytes, 0, $bytes.Length)
                }
                finally {
                    $entryStream.Dispose()
                }
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

New-ZipPackage -Path (Join-Path $cases "docx-private-grounded-review.docx") -Entries ([ordered]@{
    "[Content_Types].xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>
</Types>
'@
    "_rels/.rels" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@
    "word/_rels/document.xml.rels" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rIdCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
</Relationships>
'@
    "word/document.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="30"/></w:rPr><w:t>Private-grounded DOCX review fixture</w:t></w:r>
    </w:p>
    <w:p>
      <w:bookmarkStart w:id="11" w:name="SyntheticAnchor"/>
      <w:r><w:t>Complex cached field: </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> REF SyntheticAnchor </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>cached reference result</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:bookmarkEnd w:id="11"/>
    </w:p>
    <w:p>
      <w:r><w:t>Anchored text box follows this paragraph.</w:t></w:r>
      <w:r>
        <w:drawing>
          <wp:anchor simplePos="0" relativeHeight="251658241" behindDoc="0" layoutInCell="1" allowOverlap="1">
            <wp:extent cx="2743200" cy="914400"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
            <wp:wrapSquare wrapText="bothSides"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:t>Anchored review note</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="ReviewSpacing"/></w:pPr>
      <w:r><w:t>Styled spacing paragraph before a numbered block.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>
      <w:r><w:t>Numbered item with hanging indent and tab suffix.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr>
      <w:r><w:t>Nested numbered item resets under the parent item.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pPrChange w:id="31" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:pPr><w:spacing w:after="360"/></w:pPr></w:pPrChange>
      </w:pPr>
      <w:r><w:t>Formatting revision on paragraph and </w:t></w:r>
      <w:r>
        <w:rPr><w:rPrChange w:id="32" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>
        <w:t>run text</w:t>
      </w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p>
    <w:p>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>Visible comment anchor</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
    </w:p>
    <w:del w:id="41" w:author="Reviewer" w:date="2026-06-01T00:00:00Z">
      <w:p>
        <w:commentRangeStart w:id="2"/>
        <w:r><w:delText>Hidden deleted comment anchor</w:delText></w:r>
        <w:commentRangeEnd w:id="2"/>
        <w:r><w:commentReference w:id="2"/></w:r>
      </w:p>
    </w:del>
    <w:p>
      <w:commentRangeStart w:id="4"/>
      <w:r><w:t>Threaded parent comment anchor</w:t></w:r>
      <w:commentRangeEnd w:id="4"/>
      <w:r><w:commentReference w:id="4"/></w:r>
    </w:p>
    <w:p><w:r><w:t>Dense revisions:</w:t></w:r></w:p>
    <w:p><w:r><w:t>Dense 01 </w:t></w:r><w:del w:id="101" w:author="A" w:date="2026-06-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="102" w:author="B" w:date="2026-06-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>
    <w:p><w:r><w:t>Dense 02 </w:t></w:r><w:del w:id="103" w:author="A" w:date="2026-06-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="104" w:author="B" w:date="2026-06-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>
    <w:p><w:r><w:t>Dense 03 </w:t></w:r><w:del w:id="105" w:author="A" w:date="2026-06-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="106" w:author="B" w:date="2026-06-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>
    <w:p><w:r><w:t>Dense 04 </w:t></w:r><w:del w:id="107" w:author="A" w:date="2026-06-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="108" w:author="B" w:date="2026-06-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="7200" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="double" w:sz="8" w:color="2F5597"/>
          <w:left w:val="dashed" w:sz="8" w:color="548235"/>
          <w:bottom w:val="dotted" w:sz="8" w:color="C55A11"/>
          <w:right w:val="single" w:sz="8" w:color="7030A0"/>
          <w:insideH w:val="dashed" w:sz="6" w:color="808080"/>
          <w:insideV w:val="dotted" w:sz="6" w:color="808080"/>
        </w:tblBorders>
        <w:tblPrChange w:id="51" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:tblPrChange>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>
      <w:tr>
        <w:trPr><w:trPrChange w:id="52" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:trPr><w:trHeight w:val="500"/></w:trPr></w:trPrChange></w:trPr>
        <w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/><w:tcPrChange w:id="53" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Border style cell A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:del w:id="54" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:r><w:delText>old cell text</w:delText></w:r></w:del><w:ins w:id="55" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:r><w:t>new cell text</w:t></w:r></w:ins></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
'@
    "word/styles.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="ReviewSpacing">
    <w:name w:val="Review Spacing"/>
    <w:pPr><w:spacing w:before="240" w:after="180" w:line="276" w:lineRule="auto"/></w:pPr>
  </w:style>
</w:styles>
'@
    "word/numbering.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="7">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="tab"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:suff w:val="space"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>
</w:numbering>
'@
    "word/comments.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="1" w:author="Reviewer" w:initials="RV" w:date="2026-06-01T00:00:00Z">
    <w:p><w:r><w:t>Visible synthetic comment.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Reviewer" w:initials="RV" w:date="2026-06-01T00:00:00Z">
    <w:p><w:r><w:t>Hidden-anchor synthetic comment.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-01T00:00:00Z">
    <w:p><w:r><w:t>Orphaned synthetic comment.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="4" w:author="Reviewer" w:initials="RV" w:date="2026-06-01T00:00:00Z">
    <w:p w14:paraId="AAAABBBB"><w:r><w:t>Thread parent synthetic comment.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="5" w:author="Reviewer" w:initials="RV" w:date="2026-06-01T00:01:00Z">
    <w:p w14:paraId="CCCCDDDD"><w:r><w:t>Thread reply synthetic comment.</w:t></w:r></w:p>
  </w:comment>
</w:comments>
'@
    "word/commentsExtended.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:commentEx w15:paraId="AAAABBBB" w15:done="0"/>
  <w15:commentEx w15:paraId="CCCCDDDD" w15:paraIdParent="AAAABBBB" w15:done="0"/>
</w15:commentsEx>
'@
})
