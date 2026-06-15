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

function Join-Xml([string[]] $Items) {
    return ($Items -join "`n")
}

function New-ContentTypes {
    param(
        [switch] $CommentsExtended,
        [switch] $Header,
        [switch] $Footnotes
    )

    $overrides = @(
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
        '  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    )
    if ($CommentsExtended) {
        $overrides += '  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>'
    }
    if ($Header) {
        $overrides += '  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    }
    if ($Footnotes) {
        $overrides += '  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
    }

@"
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
$(Join-Xml $overrides)
</Types>
"@
}

function New-DocumentRelationships {
    param(
        [switch] $CommentsExtended,
        [switch] $Header,
        [switch] $Footnotes
    )

    $relationships = @(
        '  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
    )
    if ($CommentsExtended) {
        $relationships += '  <Relationship Id="rIdCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>'
    }
    if ($Header) {
        $relationships += '  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    }
    if ($Footnotes) {
        $relationships += '  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
    }

@"
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
$(Join-Xml $relationships)
</Relationships>
"@
}

function New-Comment {
    param(
        [int] $Id,
        [string] $Author,
        [string] $Initials,
        [string] $Date,
        [string] $ParagraphId,
        [string[]] $Paragraphs
    )

    $body = @()
    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        $paragraphIdAttribute = if ($i -eq 0 -and -not [string]::IsNullOrWhiteSpace($ParagraphId)) { " w14:paraId=`"$ParagraphId`"" } else { "" }
        $body += "    <w:p$paragraphIdAttribute><w:r><w:t>$($Paragraphs[$i])</w:t></w:r></w:p>"
    }

@"
  <w:comment w:id="$Id" w:author="$Author" w:initials="$Initials" w:date="$Date">
$(Join-Xml $body)
  </w:comment>
"@
}

function New-CommentsPart([string[]] $Comments) {
@"
<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
$(Join-Xml $Comments)
</w:comments>
"@
}

function New-CommentsExtendedPart([string[]] $Items) {
@"
<?xml version="1.0" encoding="UTF-8"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
$(Join-Xml $Items)
</w15:commentsEx>
"@
}

function New-DocumentXml {
    param(
        [string] $BodyXml,
        [string] $SectionXml
    )

@"
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
$BodyXml
$SectionXml
  </w:body>
</w:document>
"@
}

function New-StandardSection {
@'
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
'@
}

function New-CommentFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FileName,

        [Parameter(Mandatory = $true)]
        [string] $BodyXml,

        [Parameter(Mandatory = $true)]
        [string[]] $Comments,

        [string] $SectionXml = (New-StandardSection),

        [string] $CommentsExtendedXml = "",

        [string] $HeaderXml = "",

        [string] $FootnotesXml = ""
    )

    $entries = @{
        "[Content_Types].xml" = New-ContentTypes `
            -CommentsExtended: (-not [string]::IsNullOrWhiteSpace($CommentsExtendedXml)) `
            -Header: (-not [string]::IsNullOrWhiteSpace($HeaderXml)) `
            -Footnotes: (-not [string]::IsNullOrWhiteSpace($FootnotesXml))
        "_rels/.rels" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@
        "word/_rels/document.xml.rels" = New-DocumentRelationships `
            -CommentsExtended: (-not [string]::IsNullOrWhiteSpace($CommentsExtendedXml)) `
            -Header: (-not [string]::IsNullOrWhiteSpace($HeaderXml)) `
            -Footnotes: (-not [string]::IsNullOrWhiteSpace($FootnotesXml))
        "word/document.xml" = New-DocumentXml -BodyXml $BodyXml -SectionXml $SectionXml
        "word/comments.xml" = New-CommentsPart $Comments
    }

    if (-not [string]::IsNullOrWhiteSpace($CommentsExtendedXml)) {
        $entries["word/commentsExtended.xml"] = $CommentsExtendedXml
    }
    if (-not [string]::IsNullOrWhiteSpace($HeaderXml)) {
        $entries["word/header1.xml"] = $HeaderXml
    }
    if (-not [string]::IsNullOrWhiteSpace($FootnotesXml)) {
        $entries["word/footnotes.xml"] = $FootnotesXml
    }

    New-ZipPackage -Path (Join-Path $cases $FileName) -Entries $entries
}

New-CommentFixture `
    -FileName "docx-markup-comment-unresolved.docx" `
    -BodyXml @'
    <w:p>
      <w:r><w:t>Public comment fixture starts before </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>the reviewed phrase</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
      <w:r><w:t> and continues after the marker.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second body paragraph contains </w:t></w:r>
      <w:commentRangeStart w:id="2"/>
      <w:r><w:t>another public anchor</w:t></w:r>
      <w:commentRangeEnd w:id="2"/>
      <w:r><w:commentReference w:id="2"/></w:r>
      <w:r><w:t> to test ordering.</w:t></w:r>
    </w:p>
'@ `
    -Comments @(
        (New-Comment 1 "Reviewer One" "R1" "2026-06-01T08:00:00Z" "10101010" @("Open parent comment for a body anchor.")),
        (New-Comment 2 "Reviewer Two" "R2" "2026-06-01T08:05:00Z" "20202020" @("Open second comment for ordering."))
    )

New-CommentFixture `
    -FileName "docx-markup-comment-threaded-resolved.docx" `
    -BodyXml @'
    <w:p>
      <w:r><w:t>Threaded comment fixture has </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>a resolved discussion anchor</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
      <w:r><w:t> in the body text.</w:t></w:r>
    </w:p>
'@ `
    -Comments @(
        (New-Comment 1 "Reviewer One" "R1" "2026-06-01T09:00:00Z" "30303030" @("Resolved parent comment.")),
        (New-Comment 2 "Reviewer Two" "R2" "2026-06-01T09:05:00Z" "40404040" @("Reply comment remains open."))
    ) `
    -CommentsExtendedXml (New-CommentsExtendedPart @(
        '  <w15:commentEx w15:paraId="30303030" w15:done="1"/>',
        '  <w15:commentEx w15:paraId="40404040" w15:paraIdParent="30303030" w15:done="0"/>'
    ))

New-CommentFixture `
    -FileName "docx-markup-comment-hidden-anchors.docx" `
    -BodyXml @'
    <w:p>
      <w:r><w:t>Visible final-view comment </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>remains anchored</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p>
    <w:del w:id="20" w:author="Reviewer" w:date="2026-06-01T10:00:00Z">
      <w:p>
        <w:r><w:delText>Deleted text before hidden comment </w:delText></w:r>
        <w:commentRangeStart w:id="2"/>
        <w:r><w:delText>hidden anchor</w:delText></w:r>
        <w:commentRangeEnd w:id="2"/>
        <w:r><w:commentReference w:id="2"/></w:r>
      </w:p>
    </w:del>
'@ `
    -Comments @(
        (New-Comment 1 "Reviewer One" "R1" "2026-06-01T10:05:00Z" "50505050" @("Visible comment body.")),
        (New-Comment 2 "Reviewer Two" "R2" "2026-06-01T10:10:00Z" "60606060" @("Comment anchored only in deleted content.")),
        (New-Comment 3 "Reviewer Three" "R3" "2026-06-01T10:15:00Z" "70707070" @("Orphan public comment body."))
    )

New-CommentFixture `
    -FileName "docx-markup-comment-long.docx" `
    -BodyXml @'
    <w:p>
      <w:r><w:t>Long comment fixture anchors </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>a paragraph with a deliberately long review note</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
      <w:r><w:t> near the top of the page.</w:t></w:r>
    </w:p>
'@ `
    -Comments @(
        (New-Comment 1 "Reviewer One" "R1" "2026-06-01T11:00:00Z" "80808080" @(
            "This public comment has enough words to exercise wrapping in a narrow review balloon and to make preview truncation visible.",
            "A second paragraph checks separator spacing and multi-paragraph preview extraction without using private document text.",
            "A third paragraph pushes the balloon toward the maximum height behavior."
        ))
    )

New-CommentFixture `
    -FileName "docx-markup-comment-table.docx" `
    -BodyXml @'
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="9000" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="6" w:color="808080"/>
          <w:left w:val="single" w:sz="6" w:color="808080"/>
          <w:bottom w:val="single" w:sz="6" w:color="808080"/>
          <w:right w:val="single" w:sz="6" w:color="808080"/>
          <w:insideH w:val="single" w:sz="6" w:color="808080"/>
          <w:insideV w:val="single" w:sz="6" w:color="808080"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Table left</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Commented table cell</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Table right</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Second row left</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Second row center</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Second row right</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
'@ `
    -Comments @(
        (New-Comment 1 "Reviewer One" "R1" "2026-06-01T12:00:00Z" "90909090" @("Table-cell comment body."))
    )

$staticHeader = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Header text before </w:t></w:r>
    <w:commentRangeStart w:id="1"/>
    <w:r><w:t>header comment anchor</w:t></w:r>
    <w:commentRangeEnd w:id="1"/>
    <w:r><w:commentReference w:id="1"/></w:r>
  </w:p>
</w:hdr>
'@
$staticFootnotes = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2">
    <w:p>
      <w:r><w:t>Footnote text before </w:t></w:r>
      <w:commentRangeStart w:id="2"/>
      <w:r><w:t>footnote comment anchor</w:t></w:r>
      <w:commentRangeEnd w:id="2"/>
      <w:r><w:commentReference w:id="2"/></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
'@
New-CommentFixture `
    -FileName "docx-markup-comment-static-stories.docx" `
    -BodyXml @'
    <w:p>
      <w:r><w:t>Body paragraph references a footnote</w:t></w:r>
      <w:r><w:footnoteReference w:id="2"/></w:r>
      <w:r><w:t> while header and footnote stories carry comment anchors.</w:t></w:r>
    </w:p>
'@ `
    -SectionXml @'
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
'@ `
    -Comments @(
        (New-Comment 1 "Reviewer One" "R1" "2026-06-01T13:00:00Z" "A0A0A0A0" @("Header story comment body.")),
        (New-Comment 2 "Reviewer Two" "R2" "2026-06-01T13:05:00Z" "B0B0B0B0" @("Footnote story comment body."))
    ) `
    -HeaderXml $staticHeader `
    -FootnotesXml $staticFootnotes

New-CommentFixture `
    -FileName "docx-markup-comment-text-box.docx" `
    -BodyXml @'
    <w:p>
      <w:r><w:t>Body paragraph before floating text box.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:anchor simplePos="0" relativeHeight="251658241" behindDoc="0" layoutInCell="1" allowOverlap="1">
            <wp:extent cx="3200400" cy="1097280"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
            <wp:wrapNone/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:r><w:t>Text box before </w:t></w:r>
                        <w:commentRangeStart w:id="1"/>
                        <w:r><w:t>commented drawing text</w:t></w:r>
                        <w:commentRangeEnd w:id="1"/>
                        <w:r><w:commentReference w:id="1"/></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>
'@ `
    -Comments @(
        (New-Comment 1 "Reviewer One" "R1" "2026-06-01T14:00:00Z" "C0C0C0C0" @("Text-box comment body."))
    )
