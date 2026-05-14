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

$pptxContentTypes = @'
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>
'@

$pptxPackageRels = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>
'@

$pptxPresentationRels = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>
'@

$pptxPresentation = @'
<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>
'@

New-ZipPackage -Path (Join-Path $cases "pptx-blank.pptx") -Entries @{
    "[Content_Types].xml" = $pptxContentTypes
    "_rels/.rels" = $pptxPackageRels
    "ppt/_rels/presentation.xml.rels" = $pptxPresentationRels
    "ppt/presentation.xml" = $pptxPresentation
    "ppt/slides/slide1.xml" = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>'
}

New-ZipPackage -Path (Join-Path $cases "pptx-shapes.pptx") -Entries @{
    "[Content_Types].xml" = $pptxContentTypes
    "_rels/.rels" = $pptxPackageRels
    "ppt/_rels/presentation.xml.rels" = $pptxPresentationRels
    "ppt/presentation.xml" = $pptxPresentation
    "ppt/slides/slide1.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F8F8F8"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1371600"/></a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="2F80ED"/></a:solidFill>
          <a:ln w="25400"><a:solidFill><a:srgbClr val="1B4F9C"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm rot="900000" flipH="1"><a:off x="4572000" y="1371600"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>
          <a:prstGeom prst="ellipse"/>
          <a:solidFill><a:srgbClr val="27AE60"/></a:solidFill>
          <a:ln w="25400"><a:solidFill><a:srgbClr val="145A32"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="4114800"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="line"/>
          <a:ln w="38100"><a:solidFill><a:srgbClr val="EB5757"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
'@
}

New-ZipPackage -Path (Join-Path $cases "docx-blank.docx") -Entries @{
    "[Content_Types].xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
'@
    "_rels/.rels" = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@
    "word/document.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p/>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>
'@
}

Get-ChildItem -LiteralPath $cases -Filter "*.pptx"
Get-ChildItem -LiteralPath $cases -Filter "*.docx"
