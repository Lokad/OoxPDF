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

function New-ReviewParagraph([int] $Index, [int] $CommentId) {
    $delId = $Index * 10 + 1
    $insId = $Index * 10 + 2
@"
    <w:p>
      <w:r><w:t>Markup margin probe $Index keeps public body text before </w:t></w:r>
      <w:del w:id="$delId" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:r><w:delText>removed wording </w:delText></w:r></w:del>
      <w:ins w:id="$insId" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:r><w:t>inserted wording </w:t></w:r></w:ins>
      <w:r><w:t>and anchors </w:t></w:r>
      <w:commentRangeStart w:id="$CommentId"/>
      <w:r><w:t>a review note</w:t></w:r>
      <w:commentRangeEnd w:id="$CommentId"/>
      <w:r><w:commentReference w:id="$CommentId"/></w:r>
      <w:r><w:t> for geometry measurement.</w:t></w:r>
    </w:p>
"@
}

function New-TitleParagraph([string] $Title) {
@"
    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="30"/></w:rPr><w:t>$Title</w:t></w:r>
    </w:p>
"@
}

function New-PageBreakParagraph {
@'
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
'@
}

function New-SpacerParagraph([int] $LineTwips) {
@"
    <w:p>
      <w:pPr><w:spacing w:line="$LineTwips" w:lineRule="exact"/></w:pPr>
      <w:r><w:t>Public spacer for review lane bands.</w:t></w:r>
    </w:p>
"@
}

function New-CommentPart([int] $Count) {
    $comments = 1..$Count | ForEach-Object {
@"
  <w:comment w:id="$_" w:author="Reviewer" w:initials="RV" w:date="2026-06-01T00:00:00Z">
    <w:p><w:r><w:t>Public comment $_ for markup margin geometry.</w:t></w:r></w:p>
  </w:comment>
"@
    }

@"
<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
$(Join-Xml $comments)
</w:comments>
"@
}

function New-DocxMarkupMarginFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FileName,

        [Parameter(Mandatory = $true)]
        [string] $BodyXml,

        [Parameter(Mandatory = $true)]
        [string] $SectionXml,

        [int] $CommentCount = 12,

        [string] $SettingsXml = ""
    )

    $settingsOverride = if ([string]::IsNullOrWhiteSpace($SettingsXml)) { "" } else { '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' }
    $settingsRelationship = if ([string]::IsNullOrWhiteSpace($SettingsXml)) { "" } else { '  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' }
    $entries = @{
        "[Content_Types].xml" = @"
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
$settingsOverride
</Types>
"@
        "_rels/.rels" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@
        "word/_rels/document.xml.rels" = @"
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
$settingsRelationship
</Relationships>
"@
        "word/document.xml" = @"
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
$BodyXml
$SectionXml
  </w:body>
</w:document>
"@
        "word/comments.xml" = New-CommentPart $CommentCount
    }

    if (-not [string]::IsNullOrWhiteSpace($SettingsXml)) {
        $entries["word/settings.xml"] = $SettingsXml
    }

    New-ZipPackage -Path (Join-Path $cases $FileName) -Entries $entries
}

$letterSection = @'
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
'@
$landscapeSection = @'
    <w:sectPr>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="1080" w:right="1440" w:bottom="1080" w:left="1440"/>
    </w:sectPr>
'@
$mirroredSection = @'
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="2160"/>
    </w:sectPr>
'@
$multiColumnSection = @'
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:cols w:num="2" w:space="720"/>
    </w:sectPr>
'@
$mirroredSettings = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:mirrorMargins/>
</w:settings>
'@

New-DocxMarkupMarginFixture `
    -FileName "docx-markup-margin-one-page.docx" `
    -BodyXml (Join-Xml @(
        (New-TitleParagraph "One-page markup margin geometry"),
        (New-ReviewParagraph 1 1),
        (New-ReviewParagraph 2 2),
        (New-ReviewParagraph 3 3))) `
    -SectionXml $letterSection

New-DocxMarkupMarginFixture `
    -FileName "docx-markup-margin-multi-page.docx" `
    -BodyXml (Join-Xml @(
        (New-TitleParagraph "Multi-page markup margin geometry"),
        (New-ReviewParagraph 1 1),
        (New-ReviewParagraph 2 2),
        (New-PageBreakParagraph),
        (New-ReviewParagraph 3 3),
        (New-ReviewParagraph 4 4),
        (New-PageBreakParagraph),
        (New-ReviewParagraph 5 5),
        (New-ReviewParagraph 6 6))) `
    -SectionXml $letterSection

New-DocxMarkupMarginFixture `
    -FileName "docx-markup-margin-landscape.docx" `
    -BodyXml (Join-Xml @(
        (New-TitleParagraph "Landscape markup margin geometry"),
        (New-ReviewParagraph 1 1),
        (New-ReviewParagraph 2 2),
        (New-ReviewParagraph 3 3))) `
    -SectionXml $landscapeSection

New-DocxMarkupMarginFixture `
    -FileName "docx-markup-margin-mirrored.docx" `
    -BodyXml (Join-Xml @(
        (New-TitleParagraph "Mirrored-margin markup geometry"),
        (New-ReviewParagraph 1 1),
        (New-PageBreakParagraph),
        (New-ReviewParagraph 2 2),
        (New-ReviewParagraph 3 3))) `
    -SectionXml $mirroredSection `
    -SettingsXml $mirroredSettings

New-DocxMarkupMarginFixture `
    -FileName "docx-markup-margin-multi-column.docx" `
    -BodyXml (Join-Xml @(
        (New-TitleParagraph "Multi-column markup margin geometry"),
        (New-ReviewParagraph 1 1),
        (New-ReviewParagraph 2 2),
        (New-ReviewParagraph 3 3),
        (New-ReviewParagraph 4 4))) `
    -SectionXml $multiColumnSection

$tableHeavyBody = Join-Xml @(
    (New-TitleParagraph "Table-heavy markup margin geometry"),
    @'
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
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc><w:p><w:r><w:t>Left public cell</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Middle public cell</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Right public cell</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Commented table anchor</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p></w:tc>
        <w:tc><w:p><w:del w:id="51" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:r><w:delText>old table value</w:delText></w:r></w:del><w:ins w:id="52" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:r><w:t>new table value</w:t></w:r></w:ins></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Wide table content tests reserved margin shrink.</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
'@,
    (New-ReviewParagraph 6 2),
    (New-ReviewParagraph 7 3))
New-DocxMarkupMarginFixture `
    -FileName "docx-markup-margin-table-heavy.docx" `
    -BodyXml $tableHeavyBody `
    -SectionXml $letterSection

$denseBody = @(New-TitleParagraph "Dense revision markup margin geometry")
foreach ($i in 1..14) {
    $denseBody += New-ReviewParagraph $i $i
}
New-DocxMarkupMarginFixture `
    -FileName "docx-markup-margin-dense-revisions.docx" `
    -BodyXml (Join-Xml $denseBody) `
    -SectionXml $letterSection `
    -CommentCount 14

New-DocxMarkupMarginFixture `
    -FileName "docx-markup-balloon-lane-bands.docx" `
    -BodyXml (Join-Xml @(
        (New-TitleParagraph "Balloon lane-band placement"),
        (New-ReviewParagraph 1 1),
        (New-SpacerParagraph 1800),
        (New-ReviewParagraph 2 2),
        (New-SpacerParagraph 1440),
        (New-ReviewParagraph 3 3))) `
    -SectionXml $letterSection `
    -CommentCount 3
