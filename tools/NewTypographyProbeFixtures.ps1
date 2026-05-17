$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

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

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>
'@

$packageRels = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>
'@

$presentationRels = @'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>
'@

$presentation = @'
<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>
'@

function New-TypographyProbe {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Id,

        [Parameter(Mandatory = $true)]
        [string] $TextBody,

        [string] $Transform = '<a:xfrm><a:off x="822960" y="685800"/><a:ext cx="7498080" cy="5486400"/></a:xfrm>',

        [string] $BodyPr = '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>'
    )

    $slide = @"
<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="$Id"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          $Transform
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:txBody>
          $BodyPr
          <a:lstStyle/>
$TextBody
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@

    $path = Join-Path $cases "$Id.pptx"
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }

    $basePackage = Join-Path $cases "pptx-ladder-04-all-caps.pptx"
    $source = [System.IO.Compression.ZipFile]::OpenRead($basePackage)
    try {
        $destinationStream = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew)
        try {
            $destination = [System.IO.Compression.ZipArchive]::new($destinationStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
            try {
                foreach ($entry in $source.Entries) {
                    if ($entry.FullName -eq "ppt/slides/slide1.xml") {
                        continue
                    }

                    $copy = $destination.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                    $sourceStream = $entry.Open()
                    try {
                        $copyStream = $copy.Open()
                        try {
                            $sourceStream.CopyTo($copyStream)
                        }
                        finally {
                            $copyStream.Dispose()
                        }
                    }
                    finally {
                        $sourceStream.Dispose()
                    }
                }

                $slideEntry = $destination.CreateEntry("ppt/slides/slide1.xml", [System.IO.Compression.CompressionLevel]::Optimal)
                $slideStream = $slideEntry.Open()
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($slide)
                    $slideStream.Write($bytes, 0, $bytes.Length)
                }
                finally {
                    $slideStream.Dispose()
                }
            }
            finally {
                $destination.Dispose()
            }
        }
        finally {
            $destinationStream.Dispose()
        }
    }
    finally {
        $source.Dispose()
    }
}

New-TypographyProbe -Id "pptx-ladder-04-typography-capital-spacing-probe" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>The scale and growth</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>Large Global Supply</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>Lokad en quelques mots</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>To Va Ta AV</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-accent-spacing-probe" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>Dépendance à l'offre</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Calibri"/></a:rPr><a:t>Dépendance élevée côté coût</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Cambria"/></a:rPr><a:t>Écart dépendance qualité</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>éèê ÉÈÊ ÀÂ Çç</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-boundary-invariance-probe" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>The scale and growth</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>The </a:t></a:r><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>scale </a:t></a:r><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>and </a:t></a:r><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>growth</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>Large </a:t></a:r><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/><a:highlight><a:srgbClr val="FFF200"/></a:highlight></a:rPr><a:t>Global </a:t></a:r><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>Supply</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/></a:rPr><a:t>Dépendance </a:t></a:r><a:r><a:rPr sz="2800"><a:latin typeface="Arial"/><a:highlight><a:srgbClr val="FFF200"/></a:highlight></a:rPr><a:t>élevée</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-whitespace-controls-probe" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha  beta   gamma</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha&#xA0;beta&#x202F;gamma</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>SKU-123 / A+B, C.D</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"><a:tabLst><a:tab pos="1828800"/></a:tabLst></a:pPr><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Left</a:t></a:r><a:tab/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Tab</a:t></a:r><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>  Spaces</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Calibri"/></a:rPr><a:t>Dépendance&#xA0;élevée - coût&#x202F;unitaire</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-repeated-spaces" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha  beta   gamma</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Calibri"/></a:rPr><a:t>Left  middle   right</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-nbsp-narrow-space" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha&#xA0;beta&#x202F;gamma</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Calibri"/></a:rPr><a:t>Stock&#xA0;available&#x202F;today</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-punctuation-boundaries" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>SKU-123 / A+B, C.D</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2600"><a:latin typeface="Calibri"/></a:rPr><a:t>Q1/Q2: A-B+C, D.E</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-tab-space" -TextBody @'
          <a:p><a:pPr algn="l"><a:tabLst><a:tab pos="1828800"/></a:tabLst></a:pPr><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Left</a:t></a:r><a:tab/><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Tab</a:t></a:r><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>  Spaces</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"><a:tabLst><a:tab pos="2286000"/></a:tabLst></a:pPr><a:r><a:rPr sz="2600"><a:latin typeface="Calibri"/></a:rPr><a:t>Name</a:t></a:r><a:tab/><a:r><a:rPr sz="2600"><a:latin typeface="Calibri"/></a:rPr><a:t>Value</a:t></a:r><a:r><a:rPr sz="2600"><a:latin typeface="Calibri"/></a:rPr><a:t>  Units</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-inventory-opti-probe" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="3000" kern="1200"><a:latin typeface="Calibri Light"/></a:rPr><a:t>Inventory Optimization</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="3000" kern="1200"><a:latin typeface="Calibri Light"/></a:rPr><a:t>Inventory </a:t></a:r><a:r><a:rPr sz="3000" kern="1200"><a:latin typeface="Calibri Light"/><a:highlight><a:srgbClr val="FFF200"/></a:highlight></a:rPr><a:t>Optimization</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="2200" kern="1200"><a:latin typeface="Calibri"/></a:rPr><a:t>In v e n t o r y Op t i should never appear from run splitting</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-dense-column-probe" -TextBody @'
          <a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1500"><a:latin typeface="Calibri"/></a:rPr><a:t>Lokad en quelques mots</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1300"><a:latin typeface="Calibri"/></a:rPr><a:t>Demand forecasting, replenishment and pricing depend on typography remaining legible in dense left-side columns.</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1300"><a:latin typeface="Calibri"/></a:rPr><a:t>Dépendance, qualité, coût, délai, écart: accented letters must not introduce phantom spacing.</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1300"><a:latin typeface="Calibri"/></a:rPr><a:t>Large Global Supply and The scale and growth exercise capital-letter pairs.</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-justify-port" -TextBody @'
          <a:p><a:pPr algn="just"/><a:r><a:rPr sz="1800"><a:latin typeface="Calibri"/></a:rPr><a:t>Paragraph one aligned justify. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</a:t></a:r></a:p>
          <a:p><a:pPr algn="just"/><a:r><a:rPr sz="1800"><a:latin typeface="Calibri"/></a:rPr><a:t>Paragraph two aligned justify. Demand planning text should distribute word spacing across wrapped lines while preserving the final visual line.</a:t></a:r></a:p>
          <a:p><a:pPr algn="just"/><a:r><a:rPr sz="1800"><a:latin typeface="Calibri"/></a:rPr><a:t>Paragraph three aligned justify. The last visual line should remain left aligned like PowerPoint even when earlier lines expand spaces.</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-alignment-values-probe" -TextBody @'
          <a:p><a:pPr algn="just"/><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Just alignment should stretch spaces between words but must never insert extra spacing between letters.</a:t></a:r></a:p>
          <a:p><a:pPr algn="dist"/><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Distributed alignment with Latin text should not produce parasite letter spacing inside words.</a:t></a:r></a:p>
          <a:p><a:pPr algn="justLow"/><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Low justify alignment with Latin text should preserve normal glyph advances inside each word.</a:t></a:r></a:p>
          <a:p><a:pPr algn="thaiDist"/><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Thai distributed alignment on Latin text should be isolated from regular word layout behavior.</a:t></a:r></a:p>
'@

New-TypographyProbe -Id "pptx-ladder-04-typography-cambria-math-run-boundaries-probe" -TextBody @'
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>The scale and growth of </a:t></a:r><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/><a:highlight><a:srgbClr val="FFF200"/></a:highlight></a:rPr><a:t>XXXXXX</a:t></a:r><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t> supply network induces inefficiencies that compound over time.</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1400" b="1"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Large Global Supply Network implies structural inefficiencies and external volatility</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1200" b="1"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Client operates production facilities across countries, producing units annually. The sheer scale of this activity leads to inefficiencies that compound over time.</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>The key to reducing internal inefficiencies is in optimizing decisions despite external uncertainties.</a:t></a:r></a:p>
'@

New-TypographyProbe `
    -Id "pptx-ladder-04-typography-slide3-narrow-cambria-probe" `
    -Transform '<a:xfrm><a:off x="775665" y="1466572"/><a:ext cx="3371526" cy="1923604"/></a:xfrm>' `
    -BodyPr '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>' `
    -TextBody @'
          <a:p><a:pPr algn="l"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr sz="1400" b="1" kern="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Large Global Supply Network means strategic constraints and operational volatility</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr sz="1200" kern="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Inventory Optimization and Planning depend on local decisions across locations. The </a:t></a:r><a:r><a:rPr sz="1200" kern="1200"><a:highlight><a:srgbClr val="FFFF00"/></a:highlight><a:latin typeface="Cambria Math"/></a:rPr><a:t>AI</a:t></a:r><a:r><a:rPr sz="1200" kern="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t> workflow must keep words readable without parasite letter spacing.</a:t></a:r></a:p>
          <a:p><a:pPr algn="l"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr sz="1200" kern="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>The key to reducing internal inefficiencies is in optimizing decisions despite external uncertainties.</a:t></a:r></a:p>
'@

Get-ChildItem -LiteralPath $cases -Filter "pptx-ladder-04-typography-*-probe.pptx"
