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

$output = Join-Path $cases "docx-markup-links-fields.docx"
New-ZipPackage -Path $output -Entries @{
    "[Content_Types].xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rIdMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/markup-main" TargetMode="External"/>
  <Relationship Id="rIdMovedOld" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/markup-moved-old" TargetMode="External"/>
  <Relationship Id="rIdMovedNew" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/markup-moved-new" TargetMode="External"/>
  <Relationship Id="rIdCommented" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/markup-commented" TargetMode="External"/>
  <Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/markup-table" TargetMode="External"/>
</Relationships>
'@
    "word/comments.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="Reviewer" w:initials="RV" w:date="2026-06-01T09:00:00Z">
    <w:p><w:r><w:t>Public comment for linked and fielded text.</w:t></w:r></w:p>
  </w:comment>
</w:comments>
'@
    "word/document.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="10" w:name="FieldTarget"/>
      <w:r><w:t>Target paragraph for fields. </w:t></w:r>
      <w:bookmarkEnd w:id="10"/>
    </w:p>
    <w:p>
      <w:r><w:t>Revised link and field: </w:t></w:r>
      <w:hyperlink r:id="rIdMain">
        <w:r><w:t>Link </w:t></w:r>
        <w:del w:id="101" w:author="Reviewer" w:date="2026-06-01T10:00:00Z"><w:r><w:delText>old </w:delText></w:r></w:del>
        <w:ins w:id="102" w:author="Reviewer" w:date="2026-06-01T10:05:00Z"><w:r><w:t>new </w:t></w:r></w:ins>
      </w:hyperlink>
      <w:fldSimple w:instr=" REF FieldTarget ">
        <w:del w:id="103" w:author="Reviewer" w:date="2026-06-01T10:10:00Z"><w:r><w:delText>field-old </w:delText></w:r></w:del>
        <w:ins w:id="104" w:author="Reviewer" w:date="2026-06-01T10:15:00Z"><w:r><w:t>field-new </w:t></w:r></w:ins>
      </w:fldSimple>
    </w:p>
    <w:p>
      <w:r><w:t>Moved link and field: </w:t></w:r>
      <w:moveFrom w:id="201" w:author="Reviewer" w:date="2026-06-01T11:00:00Z">
        <w:hyperlink r:id="rIdMovedOld"><w:r><w:t>moved old link </w:t></w:r></w:hyperlink>
        <w:fldSimple w:instr=" REF FieldTarget "><w:r><w:t>moved old field </w:t></w:r></w:fldSimple>
      </w:moveFrom>
      <w:moveTo w:id="202" w:author="Reviewer" w:date="2026-06-01T11:05:00Z">
        <w:hyperlink r:id="rIdMovedNew"><w:r><w:t>moved new link </w:t></w:r></w:hyperlink>
        <w:fldSimple w:instr=" REF FieldTarget "><w:r><w:t>moved new field </w:t></w:r></w:fldSimple>
      </w:moveTo>
    </w:p>
    <w:p>
      <w:r><w:t>Commented link and field: </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:hyperlink r:id="rIdCommented"><w:r><w:t>commented link </w:t></w:r></w:hyperlink>
      <w:fldSimple w:instr=" REF FieldTarget "><w:r><w:t>commented field </w:t></w:r></w:fldSimple>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
    </w:p>
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
      <w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:hyperlink r:id="rIdTable"><w:r><w:t>table link</w:t></w:r></w:hyperlink></w:p></w:tc>
        <w:tc><w:p><w:fldSimple w:instr=" REF FieldTarget "><w:r><w:t>table field</w:t></w:r></w:fldSimple></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
'@
}

$noteOutput = Join-Path $cases "docx-markup-note-links-fields.docx"
New-ZipPackage -Path $noteOutput -Entries @{
    "[Content_Types].xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
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
  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>
'@
    "word/_rels/footnotes.xml.rels" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFootnoteLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/markup-footnote" TargetMode="External"/>
</Relationships>
'@
    "word/_rels/endnotes.xml.rels" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdEndnoteLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/markup-endnote" TargetMode="External"/>
</Relationships>
'@
    "word/document.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="20" w:name="NoteFieldTarget"/>
      <w:r><w:t>Target paragraph for note links and fields.</w:t></w:r>
      <w:bookmarkEnd w:id="20"/>
    </w:p>
    <w:p>
      <w:r><w:t>Placed note probes: </w:t></w:r>
      <w:r><w:footnoteReference w:id="5"/></w:r>
      <w:r><w:t> and </w:t></w:r>
      <w:r><w:endnoteReference w:id="7"/></w:r>
    </w:p>
    <w:p>
      <w:r><w:br w:type="page"/></w:r>
      <w:r><w:t>Second page body forces dynamic note page fields.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
'@
    "word/footnotes.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:footnote w:id="5">
    <w:p>
      <w:r><w:t>Footnote </w:t></w:r>
      <w:hyperlink r:id="rIdFootnoteLink">
        <w:r><w:t>external link</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t> page </w:t></w:r>
      <w:fldSimple w:instr=" PAGE "/>
      <w:r><w:t> of </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> NUMPAGES </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>9</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
'@
    "word/endnotes.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:endnote w:id="7">
    <w:p>
      <w:r><w:t>Endnote </w:t></w:r>
      <w:hyperlink r:id="rIdEndnoteLink">
        <w:r><w:t>external link</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t> page </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>3</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t> of </w:t></w:r>
      <w:fldSimple w:instr=" NUMPAGES "/>
    </w:p>
  </w:endnote>
</w:endnotes>
'@
}
