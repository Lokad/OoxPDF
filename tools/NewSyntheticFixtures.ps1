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

function Use-OfficePptxContainerForSlide {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $BasePath
    )

    $pathFull = [System.IO.Path]::GetFullPath($Path)
    $baseFull = [System.IO.Path]::GetFullPath($BasePath)
    $temporaryPath = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($pathFull),
        ([System.IO.Path]::GetFileNameWithoutExtension($pathFull) + ".office-container.tmp.pptx"))

    $sourceZip = [System.IO.Compression.ZipFile]::OpenRead($pathFull)
    try {
        $slideEntry = $sourceZip.GetEntry("ppt/slides/slide1.xml")
        $reader = [System.IO.StreamReader]::new($slideEntry.Open())
        try {
            $slideXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $sourceZip.Dispose()
    }

    if (Test-Path -LiteralPath $temporaryPath) {
        Remove-Item -LiteralPath $temporaryPath -Force
    }

    $baseZip = [System.IO.Compression.ZipFile]::OpenRead($baseFull)
    $targetZip = [System.IO.Compression.ZipFile]::Open($temporaryPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $baseZip.Entries) {
            $targetEntry = $targetZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $targetStream = $targetEntry.Open()
            try {
                if ($entry.FullName -eq "ppt/slides/slide1.xml") {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($slideXml)
                    $targetStream.Write($bytes, 0, $bytes.Length)
                }
                else {
                    $sourceStream = $entry.Open()
                    try {
                        $sourceStream.CopyTo($targetStream)
                    }
                    finally {
                        $sourceStream.Dispose()
                    }
                }
            }
            finally {
                $targetStream.Dispose()
            }
        }
    }
    finally {
        $targetZip.Dispose()
        $baseZip.Dispose()
    }

    Move-Item -LiteralPath $temporaryPath -Destination $pathFull -Force
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

New-ZipPackage -Path (Join-Path $cases "pptx-ladder-06-explicit-arc-stealth.pptx") -Entries @{
    "[Content_Types].xml" = $pptxContentTypes
    "_rels/.rels" = $pptxPackageRels
    "ppt/_rels/presentation.xml.rels" = $pptxPresentationRels
    "ppt/presentation.xml" = $pptxPresentation
    "ppt/slides/slide1.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Explicit Arc 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm flipV="1"><a:off x="720000" y="720000"/><a:ext cx="1433572" cy="1190088"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 19695970"/></a:avLst></a:prstGeom>
          <a:noFill/>
          <a:ln><a:solidFill><a:srgbClr val="2F5597"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Explicit Arc 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm flipV="1"><a:off x="2574000" y="720000"/><a:ext cx="1854923" cy="1661001"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 20800487"/></a:avLst></a:prstGeom>
          <a:noFill/>
          <a:ln><a:solidFill><a:srgbClr val="548235"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Explicit Arc 3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm rot="10800000" flipV="1"><a:off x="5100000" y="720000"/><a:ext cx="1989782" cy="1590214"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16862351"/><a:gd name="adj2" fmla="val 20776681"/></a:avLst></a:prstGeom>
          <a:noFill/>
          <a:ln><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="Explicit Arc 4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm flipV="1"><a:off x="720000" y="3030000"/><a:ext cx="2526451" cy="1532299"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16579230"/><a:gd name="adj2" fmla="val 18561823"/></a:avLst></a:prstGeom>
          <a:noFill/>
          <a:ln><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="6" name="Explicit Arc 5"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="4200000" y="3150000"/><a:ext cx="1830187" cy="1190088"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 19336501"/></a:avLst></a:prstGeom>
          <a:noFill/>
          <a:ln><a:solidFill><a:srgbClr val="BF9000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="7" name="Explicit Arc 6"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm flipV="1"><a:off x="6500000" y="3300000"/><a:ext cx="312814" cy="312135"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16546552"/><a:gd name="adj2" fmla="val 5274365"/></a:avLst></a:prstGeom>
          <a:noFill/>
          <a:ln w="15875"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="8" name="Explicit Arc 7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm rot="10800000" flipV="1"><a:off x="7100000" y="3300000"/><a:ext cx="312814" cy="312135"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16546552"/><a:gd name="adj2" fmla="val 5189620"/></a:avLst></a:prstGeom>
          <a:noFill/>
          <a:ln w="15875"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
'@
}

Use-OfficePptxContainerForSlide `
    -Path (Join-Path $cases "pptx-ladder-06-explicit-arc-stealth.pptx") `
    -BasePath (Join-Path $cases "pptx-ladder-06-shape-adjust-port-a.pptx")

New-ZipPackage -Path (Join-Path $cases "pptx-ladder-06-custom-geometry-default-fill.pptx") -Entries @{
    "[Content_Types].xml" = $pptxContentTypes
    "_rels/.rels" = $pptxPackageRels
    "ppt/_rels/presentation.xml.rels" = $pptxPresentationRels
    "ppt/presentation.xml" = $pptxPresentation
    "ppt/slides/slide1.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Implicit Fill Custom Geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr bwMode="auto">
          <a:xfrm><a:off x="900000" y="1200000"/><a:ext cx="1800000" cy="1400000"/></a:xfrm>
          <a:custGeom>
            <a:avLst/>
            <a:pathLst>
              <a:path w="100" h="100">
                <a:moveTo><a:pt x="0" y="35"/></a:moveTo>
                <a:lnTo><a:pt x="45" y="0"/></a:lnTo>
                <a:lnTo><a:pt x="100" y="25"/></a:lnTo>
                <a:lnTo><a:pt x="80" y="100"/></a:lnTo>
                <a:lnTo><a:pt x="20" y="85"/></a:lnTo>
                <a:close/>
              </a:path>
            </a:pathLst>
          </a:custGeom>
          <a:ln w="3175"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Explicit No Fill Custom Geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr bwMode="auto">
          <a:xfrm><a:off x="3400000" y="1200000"/><a:ext cx="1800000" cy="1400000"/></a:xfrm>
          <a:custGeom>
            <a:avLst/>
            <a:pathLst>
              <a:path w="100" h="100">
                <a:moveTo><a:pt x="0" y="35"/></a:moveTo>
                <a:lnTo><a:pt x="45" y="0"/></a:lnTo>
                <a:lnTo><a:pt x="100" y="25"/></a:lnTo>
                <a:lnTo><a:pt x="80" y="100"/></a:lnTo>
                <a:lnTo><a:pt x="20" y="85"/></a:lnTo>
                <a:close/>
              </a:path>
            </a:pathLst>
          </a:custGeom>
          <a:noFill/>
          <a:ln w="3175"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Explicit Solid Fill Custom Geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr bwMode="auto">
          <a:xfrm><a:off x="5900000" y="1200000"/><a:ext cx="1800000" cy="1400000"/></a:xfrm>
          <a:custGeom>
            <a:avLst/>
            <a:pathLst>
              <a:path w="100" h="100">
                <a:moveTo><a:pt x="0" y="35"/></a:moveTo>
                <a:lnTo><a:pt x="45" y="0"/></a:lnTo>
                <a:lnTo><a:pt x="100" y="25"/></a:lnTo>
                <a:lnTo><a:pt x="80" y="100"/></a:lnTo>
                <a:lnTo><a:pt x="20" y="85"/></a:lnTo>
                <a:close/>
              </a:path>
            </a:pathLst>
          </a:custGeom>
          <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
          <a:ln w="3175"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="5" name="Inherited Fill Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr bwMode="auto">
          <a:xfrm><a:off x="900000" y="3600000"/><a:ext cx="4300000" cy="1500000"/><a:chOff x="900000" y="3600000"/><a:chExt cx="4300000" cy="1500000"/></a:xfrm>
          <a:solidFill><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:solidFill>
        </p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="6" name="Group Inherited Fill Custom Geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr bwMode="auto">
            <a:xfrm><a:off x="900000" y="3600000"/><a:ext cx="1800000" cy="1200000"/></a:xfrm>
            <a:custGeom>
              <a:avLst/>
              <a:pathLst>
                <a:path w="100" h="100">
                  <a:moveTo><a:pt x="0" y="35"/></a:moveTo>
                  <a:lnTo><a:pt x="45" y="0"/></a:lnTo>
                  <a:lnTo><a:pt x="100" y="25"/></a:lnTo>
                  <a:lnTo><a:pt x="80" y="100"/></a:lnTo>
                  <a:lnTo><a:pt x="20" y="85"/></a:lnTo>
                  <a:close/>
                </a:path>
              </a:pathLst>
            </a:custGeom>
            <a:grpFill/>
            <a:ln w="3175"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln>
          </p:spPr>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="7" name="Group Inherited Fill Explicit No Fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr bwMode="auto">
            <a:xfrm><a:off x="3400000" y="3600000"/><a:ext cx="1800000" cy="1200000"/></a:xfrm>
            <a:custGeom>
              <a:avLst/>
              <a:pathLst>
                <a:path w="100" h="100">
                  <a:moveTo><a:pt x="0" y="35"/></a:moveTo>
                  <a:lnTo><a:pt x="45" y="0"/></a:lnTo>
                  <a:lnTo><a:pt x="100" y="25"/></a:lnTo>
                  <a:lnTo><a:pt x="80" y="100"/></a:lnTo>
                  <a:lnTo><a:pt x="20" y="85"/></a:lnTo>
                  <a:close/>
                </a:path>
              </a:pathLst>
            </a:custGeom>
            <a:noFill/>
            <a:ln w="3175"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln>
          </p:spPr>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>
'@
}

Use-OfficePptxContainerForSlide `
    -Path (Join-Path $cases "pptx-ladder-06-custom-geometry-default-fill.pptx") `
    -BasePath (Join-Path $cases "pptx-ladder-06-shape-adjust-port-a.pptx")

New-ZipPackage -Path (Join-Path $cases "pptx-ladder-06-straight-stealth-connectors.pptx") -Entries @{
    "[Content_Types].xml" = $pptxContentTypes
    "_rels/.rels" = $pptxPackageRels
    "ppt/_rels/presentation.xml.rels" = $pptxPresentationRels
    "ppt/presentation.xml" = $pptxPresentation
    "ppt/slides/slide1.xml" = @'
<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="2" name="Straight Connector Stealth 1pt Horizontal"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1200000" y="1500000"/><a:ext cx="2400000" cy="0"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="2F856A"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:cxnSp>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="3" name="Straight Connector Stealth 1pt Vertical"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm flipV="1"><a:off x="4800000" y="1200000"/><a:ext cx="0" cy="1800000"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="2F856A"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:cxnSp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Line Preset Stealth 1pt Control"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1200000" y="3000000"/><a:ext cx="2400000" cy="0"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:sp>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="5" name="Straight Connector Stealth 2pt Horizontal"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1200000" y="4200000"/><a:ext cx="2400000" cy="0"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="25400"><a:solidFill><a:srgbClr val="2F856A"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
        </p:spPr>
      </p:cxnSp>
    </p:spTree>
  </p:cSld>
</p:sld>
'@
}

Use-OfficePptxContainerForSlide `
    -Path (Join-Path $cases "pptx-ladder-06-straight-stealth-connectors.pptx") `
    -BasePath (Join-Path $cases "pptx-ladder-06-shape-adjust-port-a.pptx")

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

New-ZipPackage -Path (Join-Path $cases "docx-ladder-02-table-explicit-font.docx") -Entries @{
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
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="7200" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:left w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:bottom w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:right w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:insideH w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:insideV w:val="single" w:sz="8" w:color="4F81BD"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>
      <w:tr>
        <w:trPr><w:trHeight w:val="720"/></w:trPr>
        <w:tc>
          <w:tcPr><w:tcW w:w="3600" w:type="dxa"/><w:shd w:val="clear" w:fill="D9EAF7"/></w:tcPr>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:color w:val="1F4E79"/></w:rPr><w:t>Alpha</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3600" w:type="dxa"/><w:shd w:val="clear" w:fill="EAF2DD"/></w:tcPr>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:color w:val="375623"/></w:rPr><w:t>Beta</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
'@
}

New-ZipPackage -Path (Join-Path $cases "docx-ladder-03-table-missing-grid-spans.docx") -Entries @{
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
    <w:p><w:r><w:t>Missing table grid with logical spans</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="4320" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:left w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:bottom w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:right w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:insideH w:val="single" w:sz="8" w:color="4F81BD"/>
          <w:insideV w:val="single" w:sz="8" w:color="4F81BD"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr>
        <w:trPr><w:trHeight w:val="720"/></w:trPr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="2880" w:type="dxa"/>
            <w:gridSpan w:val="2"/>
            <w:shd w:val="clear" w:fill="D9EAF7"/>
          </w:tcPr>
          <w:p><w:r><w:rPr><w:sz w:val="24"/><w:color w:val="1F4E79"/></w:rPr><w:t>Two columns</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="1440" w:type="dxa"/><w:shd w:val="clear" w:fill="EAF2DD"/></w:tcPr>
          <w:p><w:r><w:rPr><w:sz w:val="24"/><w:color w:val="375623"/></w:rPr><w:t>Tail</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:trPr><w:trHeight w:val="720"/></w:trPr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="2880" w:type="dxa"/>
            <w:gridSpan w:val="2"/>
            <w:shd w:val="clear" w:fill="FCE4D6"/>
          </w:tcPr>
          <w:p><w:r><w:rPr><w:sz w:val="24"/><w:color w:val="843C0C"/></w:rPr><w:t>Wide cell</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="1440" w:type="dxa"/><w:shd w:val="clear" w:fill="E2F0D9"/></w:tcPr>
          <w:p><w:r><w:rPr><w:sz w:val="24"/><w:color w:val="548235"/></w:rPr><w:t>End</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
'@
}

Get-ChildItem -LiteralPath $cases -Filter "*.pptx"
Get-ChildItem -LiteralPath $cases -Filter "*.docx"
