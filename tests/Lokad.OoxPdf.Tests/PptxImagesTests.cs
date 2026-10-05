using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxImagesTests
{
    public static void PptxImageRecolorPreservesRawOoxmlTokens()
    {
        XElement luminancePicture = XElement.Parse("""
            <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:blipFill>
                <a:blip>
                  <a:lum bright="25000" contrast="-15000"/>
                </a:blip>
              </p:blipFill>
            </p:pic>
            """);
        PptxSceneImageRecolor luminance = PptxSceneBuilder.ReadImageRecolor(luminancePicture, PptxTheme.Empty);
        TestAssert.Equal(PptxSceneImageRecolorKind.Luminance, luminance.Kind);
        TestAssert.Equal("lum", luminance.KindValue ?? string.Empty);
        TestAssert.Equal("25000", luminance.BrightnessValue ?? string.Empty);
        TestAssert.Equal("-15000", luminance.ContrastValue ?? string.Empty);
        TestAssert.Equal(0.25d, luminance.Brightness);
        TestAssert.Equal(-0.15d, luminance.Contrast);

        XElement biLevelPicture = XElement.Parse("""
            <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:blipFill>
                <a:blip>
                  <a:biLevel thresh="62500"/>
                </a:blip>
              </p:blipFill>
            </p:pic>
            """);
        PptxSceneImageRecolor biLevel = PptxSceneBuilder.ReadImageRecolor(biLevelPicture, PptxTheme.Empty);
        TestAssert.Equal(PptxSceneImageRecolorKind.BiLevel, biLevel.Kind);
        TestAssert.Equal("biLevel", biLevel.KindValue ?? string.Empty);
        TestAssert.Equal("62500", biLevel.ThresholdValue ?? string.Empty);
        TestAssert.Equal(0.625d, biLevel.Threshold);
    }

    public static void PptxImageTilePreservesRawOoxmlTokens()
    {
        XElement tiledPicture = XElement.Parse("""
            <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:blipFill>
                <a:blip/>
                <a:tile algn="tl" flip="xy" sx="75000" sy="125000" tx="12700" ty="-25400"/>
              </p:blipFill>
            </p:pic>
            """);
        PptxScenePictureTile tile = PptxSceneBuilder.ReadPictureTile(tiledPicture);
        TestAssert.True(tile.HasTile, "Expected scene model to preserve image tile mode.");
        TestAssert.Equal("tile", tile.TileValue ?? string.Empty);
        TestAssert.Equal("tl", tile.AlignmentValue ?? string.Empty);
        TestAssert.Equal("xy", tile.FlipValue ?? string.Empty);
        TestAssert.Equal("75000", tile.ScaleXValue ?? string.Empty);
        TestAssert.Equal("125000", tile.ScaleYValue ?? string.Empty);
        TestAssert.Equal("12700", tile.OffsetXValue ?? string.Empty);
        TestAssert.Equal("-25400", tile.OffsetYValue ?? string.Empty);
    }

    public static void PptxImageAlphaPreservesRawOoxmlTokens()
    {
        XElement slidePicture = XElement.Parse("""
            <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:blipFill>
                <a:blip><a:alphaModFix amt="37500"/></a:blip>
              </p:blipFill>
            </p:pic>
            """);
        TestAssert.Equal(0.375d, PptxSceneBuilder.ReadPictureAlpha(slidePicture));
        TestAssert.Equal("37500", PptxSceneBuilder.ReadPictureAlphaValue(slidePicture) ?? string.Empty);

        XElement shapePictureFill = XElement.Parse("""
            <p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:blipFill>
                <a:blip><a:alphaModFix amt="62500"/></a:blip>
              </a:blipFill>
            </p:spPr>
            """);
        TestAssert.Equal(0.625d, PptxSceneBuilder.ReadPictureAlpha(shapePictureFill));
        TestAssert.Equal("62500", PptxSceneBuilder.ReadPictureAlphaValue(shapePictureFill) ?? string.Empty);
    }

    public static void PptxSyntheticPngPictureRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/XObject", pdf);
        TestAssert.True(
            PptxTests.CountOccurrences(pdf, "0 0 720 540 re W* n") == 1,
            "Pictures should rely on the page clip and their own frame clip, not an extra slide-sized node clip.");
        TestAssert.Contains("72 396 144 72 re W* n", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
    }

    public static void PptxSyntheticPngPictureOuterShadowUsesRasterSoftMask()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:effectLst><a:outerShdw blurRad="45720" dist="91440" dir="0"><a:srgbClr val="000000"><a:alpha val="43000"/></a:srgbClr></a:outerShdw></a:effectLst>
                    </p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int shadowDraw = pdf.IndexOf("/Im1 Do", StringComparison.Ordinal);
        int pictureDraw = pdf.IndexOf("/Im2 Do", StringComparison.Ordinal);
        TestAssert.True(shadowDraw >= 0 && pictureDraw > shadowDraw, "Picture outer shadow should draw before the picture image.");
        TestAssert.Contains("/GS43000F100000S gs", pdf);
        TestAssert.Contains("/SMask", pdf);
        TestAssert.Contains("/Width 153 /Height 81", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_EFFECT"), "Supported picture outer shadow should not emit an unsupported-effect diagnostic.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TRANSPARENCY"), "Rendered picture shadow alpha should not emit an unsupported-transparency diagnostic.");
    }

    public static void PptxSyntheticPngPictureRendersPictureOutline()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:ln w="25400" cap="sq"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("2 w", pdf);
        TestAssert.Contains("2 J", pdf);
        TestAssert.Contains("71 395 146 74 re S", pdf);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxSceneNodeSnapshot picture = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes.Single();
        TestAssert.True(picture.HasPicture, "Expected scene inspection to identify the picture node.");
        TestAssert.True(picture.PictureHasLine, "Expected scene inspection to expose the picture outline.");
        TestAssert.Equal(2d, picture.PictureLineWidth);
        TestAssert.True(picture.PictureLineWidthSpecified, "Expected the authored picture outline width to stay visible.");
    }

    public static void PptxSyntheticPngPictureClipIntersectsSlideBounds()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="-45720"/><a:ext cx="914400" cy="274320"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 522 72 18 re W* n", pdf);
        TestAssert.DoesNotContain("72 522 72 21.6 re W* n", pdf);
    }

    public static void PptxSyntheticTransformedPngPictureUsesSlideClip()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm flipV="1"><a:off x="914400" y="-45720"/><a:ext cx="914400" cy="274320"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(
            PptxTests.CountOccurrences(pdf, "0 0 720 540 re W* n") == 2,
            "Transformed pictures should add a slide clip before applying the picture transform.");
        TestAssert.Contains("72 522 72 21.6 re W* n", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticSvgPictureRendersVectorPath()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8("""
                <svg viewBox="0 0 20 10" xmlns="http://www.w3.org/2000/svg">
                  <path d="M0 0 L20 0 L20 10 L0 10 Z" fill="#112233"/>
                </svg>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.067 0.133 0.2 rg", pdf);
        TestAssert.Contains(" m", pdf);
        TestAssert.Contains(" l", pdf);
        TestAssert.DoesNotContain("/Subtype /Image", pdf);
    }

    // RV07: unsupported SVG content must diagnose instead of silently vanishing.
    public static void PptxSyntheticSvgUnsupportedContentEmitsDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8("""
                <svg viewBox="0 0 20 10" xmlns="http://www.w3.org/2000/svg">
                  <path d="M0 0 L20 0 L20 10 L0 10 Z" fill="#112233"/>
                  <text x="10" y="5">Hello</text>
                  <path d="M0 0 Q10 10 20 0 Z" fill="#778899"/>
                </svg>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            DiagnosticSink = diagnostics.Add,
        });
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.067 0.133 0.2 rg", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("text", StringComparison.Ordinal)), "Unsupported SVG element must diagnose.");
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("Q", StringComparison.Ordinal)), "Unsupported SVG path command must diagnose.");
    }

    public static void PptxSyntheticSvgPictureAppliesSourceCropBeforeScaling()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill>
                      <a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip>
                      <a:srcRect b="50000"/>
                    </p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8("""
                <svg viewBox="0 0 20 10" xmlns="http://www.w3.org/2000/svg">
                  <path d="M0 0 L20 0 L20 10 L0 10 Z" fill="#112233"/>
                </svg>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("216 324 l", pdf);
        TestAssert.DoesNotContain("216 396 l", pdf);
    }

    public static void PptxSyntheticSvgPictureSamplesCompoundPathGradient()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8("""
                <svg viewBox="0 0 100 20" xmlns="http://www.w3.org/2000/svg">
                  <linearGradient id="g" x1="0" y1="0" x2="100" y2="0" gradientUnits="userSpaceOnUse">
                    <stop offset="0" stop-color="#FF0000"/>
                    <stop offset="1" stop-color="#0000FF"/>
                  </linearGradient>
                  <path d="M0 0 H20 V20 H0 Z M80 0 H100 V20 H80 Z" fill="url(#g)"/>
                </svg>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.988 0 0.012 rg", pdf);
        TestAssert.Contains("0.012 0 0.988 rg", pdf);
    }

    private static string WriteSvgGradientDeck(string svg)
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8(svg)
        });
    }

    // RV07-S1: stretching a picture scales the stroke normal independently
    // of the path tangent. Read the emitted PDF geometry, including its CTM.
    public static void PptxSvgViewportStretchPreservesDirectionalStrokeWidths()
    {
        foreach (double size in new[] { 0.1d, 100d, 1000d })
        {
            string input = WriteSvgGradientDeck(FormattableString.Invariant($"""
                <svg viewBox="0 0 {size} {size}" xmlns="http://www.w3.org/2000/svg">
                  <path d="M{size * .15} {size * .25}H{size * .85}" fill="none" stroke="#0000FF" stroke-width="{size * .06}"/>
                  <path d="M{size * .25} {size * .15}V{size * .85}" fill="none" stroke="#0000FF" stroke-width="{size * .06}"/>
                </svg>
                """));
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            OoxPdfConverter.Convert(input, output);
            List<double> widths = ReadAxisAlignedPdfStrokeWidths(File.ReadAllText(output, Encoding.ASCII));
            TestAssert.Equal(2, widths.Count);
            TestAssert.True(Math.Abs(widths[0] - 4.32d) < .004d, $"Horizontal stroke thickness was {widths[0]}pt for source size {size}.");
            TestAssert.True(Math.Abs(widths[1] - 8.64d) < .004d, $"Vertical stroke thickness was {widths[1]}pt for source size {size}.");
        }
    }

    public static void PptxSvgExtremeViewportStrokeKeepsFinitePaint()
    {
        foreach (var probe in new[]
        {
            (ViewBox: "0 0 1000000 100", Path: "M100000 20H900000"),
            (ViewBox: "0 0 100 1000000", Path: "M10 200000H90"),
        })
        {
            string input = WriteSvgGradientDeck($"""
                <svg viewBox="{probe.ViewBox}" xmlns="http://www.w3.org/2000/svg">
                  <path d="{probe.Path}" fill="none" stroke="#0000FF" stroke-width="6"/>
                </svg>
                """);
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
            List<double> widths = ReadAxisAlignedPdfStrokeWidths(File.ReadAllText(output, Encoding.ASCII));
            TestAssert.Equal(1, widths.Count);
            TestAssert.True(double.IsFinite(widths[0]) && widths[0] > 0d, "Unrepresentable viewport anisotropy must retain finite stroke paint.");
            TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Extreme aspect ratios must not discard the picture.");
        }
    }

    private static List<double> ReadAxisAlignedPdfStrokeWidths(string pdf)
    {
        var widths = new List<double>();
        var states = new Stack<(double X, double Y, double Width)>();
        (double X, double Y, double Width) state = (1d, 1d, 1d);
        double startX = 0d, startY = 0d, endX = 0d, endY = 0d;
        foreach (string line in pdf.Split('\n'))
        {
            string[] words = line.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) continue;
            double Number(int index) => double.Parse(words[index], CultureInfo.InvariantCulture);
            switch (words[^1])
            {
                case "q": states.Push(state); break;
                case "Q": state = states.Pop(); break;
                case "cm":
                    TestAssert.True(Number(1) == 0d && Number(2) == 0d, "This probe expects axis-aligned PDF transforms.");
                    state = (state.X * Number(0), state.Y * Number(3), state.Width);
                    break;
                case "w": state.Width = Number(0); break;
                case "m": startX = endX = Number(0); startY = endY = Number(1); break;
                case "l": endX = Number(0); endY = Number(1); break;
                case "S":
                    double dx = endX - startX, dy = endY - startY;
                    double sourceLength = Math.Sqrt(dx * dx + dy * dy);
                    double pageLength = Math.Sqrt(Math.Pow(state.X * dx, 2) + Math.Pow(state.Y * dy, 2));
                    widths.Add(state.Width * Math.Abs(state.X * state.Y) * sourceLength / pageLength);
                    break;
            }
        }
        return widths;
    }

    // RV07: default objectBoundingBox gradients normalize into path space.
    public static void PptxSyntheticSvgObjectBoundingBoxGradientVariesAcrossStrips()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 20" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="1" y2="0">
                <stop offset="0" stop-color="#FF0000"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M0 0 H100 V20 H0 Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.988 0 0.012 rg", pdf);
        TestAssert.Contains("0.012 0 0.988 rg", pdf);
    }

    // RV07: vertical gradients tessellate along the dominant axis.
    public static void PptxSyntheticSvgVerticalGradientVariesTopToBottom()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0" stop-color="#FF0000"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M0 0 H100 V100 H0 Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.988 0 0.012 rg", pdf);
        TestAssert.Contains("0.012 0 0.988 rg", pdf);
    }

    // RV07: diagonal gradients project onto the dominant axis.
    public static void PptxSyntheticSvgDiagonalGradientProjectsOntoDominantAxis()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
                <stop offset="0" stop-color="#FF0000"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M0 0 H100 V100 H0 Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.745 0 0.255 rg", pdf);
        TestAssert.Contains("0.255 0 0.745 rg", pdf);
    }

    // RV07: path transforms shift emitted coordinates.
    public static void PptxSyntheticSvgTranslateShiftsPath()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V100 H0 Z" fill="#FF0000" transform="translate(10,20)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("86.4 453.6 m", pdf);
        TestAssert.Contains("230.4 453.6 l", pdf);
    }

    // RV07: path scale transforms scale emitted coordinates.
    public static void PptxSyntheticSvgScaleScalesPath()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V100 H0 Z" fill="#FF0000" transform="scale(2)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 468 m", pdf);
        TestAssert.Contains("360 468 l", pdf);
    }

    // RV07: path rotation rotates emitted coordinates.
    public static void PptxSyntheticSvgRotateTurnsPath()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V100 H0 Z" fill="#FF0000" transform="rotate(90)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 468 m", pdf);
        TestAssert.Contains("72 396 l", pdf);
        TestAssert.Contains("-72 396 l", pdf);
    }

    // RV07: ancestor group transforms apply to nested paths.
    public static void PptxSyntheticSvgGroupTransformShiftsNestedPath()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <g transform="translate(0,20)"><path d="M0 0 H100 V100 H0 Z" fill="#FF0000"/></g>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 453.6 m", pdf);
    }

    // RV07: unsupported transforms skip the path with a diagnostic while siblings render.
    public static void PptxSyntheticSvgUnsupportedTransformSkipsPath()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V100 H0 Z" fill="#FF0000"/>
              <path d="M10 10 H20 V20 H10 Z" fill="#FF0000" transform="bogus(1)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unsupported transforms", StringComparison.Ordinal)), "Unsupported transforms must diagnose.");
        TestAssert.Contains("72 468 m", pdf);
        TestAssert.DoesNotContain("86.4 460.8 m", pdf);
    }

    // RV07: style-attribute fills paint like fill attributes.
    public static void PptxSyntheticSvgStyleFillPaintsLikeFillAttribute()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 20" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V20 H0 Z" style="fill:#00FF00"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 1 0 rg", pdf);
    }

    // RV07: style fills win over fill attributes per CSS priority.
    public static void PptxSyntheticSvgStyleFillBeatsFillAttribute()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 20" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V20 H0 Z" fill="#FF0000" style="fill:#0000FF"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.DoesNotContain("1 0 0 rg", pdf);
    }

    // RV07: fill opacity emits a matching transparency group.
    public static void PptxSyntheticSvgFillOpacityEmitsTransparency()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 20" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V20 H0 Z" fill="#FF0000" fill-opacity="0.5"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/GS50000F100000S gs", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    // RV07: style named strokes resolve while the default black fill renders (second falsification: the diagnose expectation pinned hex-only parsing and was never Office-measured).
    public static void PptxSyntheticSvgStyleNamedStrokeRenders()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 20" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V20 H0 Z" style="stroke:blue"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\n0 g\n", pdf.Replace("\r\n", "\n"));
        TestAssert.Contains("\n0 0 1 RG\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: stroked-only paths render their stroke instead of vanishing.
    public static void PptxSyntheticSvgStrokedOnlyPathRendersStroke()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("1.44 w", pdf);
        TestAssert.Contains("\nS\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: paths with both fill and stroke paint both.
    public static void PptxSyntheticSvgFillAndStrokePathPaintsBoth()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="#FF0000" stroke="#0000FF"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0 0 1 RG", pdf);
        TestAssert.Contains("\nB\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: stroke widths scale from user units to points with the viewBox mapping.
    public static void PptxSyntheticSvgStrokeWidthScalesWithViewBox()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-width="2"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("2.88 w", pdf);
    }

    // RV07: a missing fill attribute defaults to black instead of vanishing the path.
    public static void PptxSyntheticSvgMissingFillDefaultsToBlack()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" stroke="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\n0 g\n", pdf.Replace("\r\n", "\n"));
        TestAssert.Contains("1 0 0 RG", pdf);
    }

    // RV07: gradient strokes diagnose while the fill still renders.
    public static void PptxSyntheticSvgGradientStrokeDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="#FF0000" stroke="url(#missing)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("stroke", StringComparison.Ordinal)), "Gradient strokes must diagnose.");
    }

    // RV07: dashed strokes render their pattern (second falsification: the diagnose-and-solid expectation pinned unmapped dashes and was never Office-measured).
    public static void PptxSyntheticSvgDashedStrokeRendersDash()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="#FF0000" stroke="#FF0000" stroke-dasharray="4 2"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("[5.76 2.88 ] 0 d", pdf);
        TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT"), "Mapped dashes must render without diagnostics.");
    }

    // RV07: group fills inherit into paths without their own fill.
    public static void PptxSyntheticSvgGroupFillInheritsToPath()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <g fill="#00FF00"><path d="M0 0 H100 V50 H0 Z"/></g>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 1 0 rg", pdf);
    }

    // RV07: group strokes inherit into paths without their own stroke.
    public static void PptxSyntheticSvgGroupStrokeInheritsToPath()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <g stroke="#0000FF"><path d="M0 0 H100 V50 H0 Z" fill="none"/></g>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 1 RG", pdf);
    }

    // RV07: element paint beats inherited paint (guard: cascade order never regresses).
    public static void PptxSyntheticSvgSelfAttributeBeatsInheritedStyle()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <g style="fill:#00FF00"><path d="M0 0 H100 V50 H0 Z" fill="#FF0000"/></g>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    // RV07: stroke widths do not inherit from groups (SVG non-inherited property).
    public static void PptxSyntheticSvgGroupStrokeWidthDoesNotInherit()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <g stroke-width="3"><path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000"/></g>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1.44 w", pdf);
    }

    // RV07: CSS named fills resolve instead of diagnosing.
    public static void PptxSyntheticSvgNamedColorFillRenders()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="red"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    // RV07: CSS named strokes resolve instead of vanishing.
    public static void PptxSyntheticSvgNamedColorStrokeRenders()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="blue"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 1 RG", pdf);
    }

    // RV07: short hex fills expand per CSS instead of diagnosing.
    public static void PptxSyntheticSvgShortHexFillRenders()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="#F00"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    // RV07: functional color syntax stays diagnosed (guard: boundary never silently renders).
    public static void PptxSyntheticSvgRgbFunctionFillDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="rgb(255,0,0)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unparsable paint", StringComparison.Ordinal)), "Functional color must diagnose.");
    }

    // RV07: rect elements render their box instead of diagnosing.
    public static void PptxSyntheticSvgRectRendersBox()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <rect x="0" y="0" width="100" height="50" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("72 468 m", pdf);
        TestAssert.Contains("216 396 l", pdf);
    }

    // RV07: circle elements render beziers instead of diagnosing.
    public static void PptxSyntheticSvgCircleRendersBeziers()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <circle cx="50" cy="25" r="10" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("158.4 432 m", pdf);
        TestAssert.Contains(" c\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: ellipse elements render beziers instead of diagnosing.
    public static void PptxSyntheticSvgEllipseRendersBeziers()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <ellipse cx="50" cy="25" rx="20" ry="10" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("172.8 432 m", pdf);
    }

    // RV07: line elements stroke their segment instead of diagnosing.
    public static void PptxSyntheticSvgLineRendersSegment()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <line x1="10" y1="10" x2="90" y2="40" fill="none" stroke="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("86.4 453.6 m", pdf);
        TestAssert.Contains("201.6 410.4 l", pdf);
    }

    // RV07: polyline elements stroke their points instead of diagnosing.
    public static void PptxSyntheticSvgPolylineRendersPoints()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <polyline points="10,10 90,10 90,40" fill="none" stroke="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("201.6 453.6 l", pdf);
        TestAssert.Contains("201.6 410.4 l", pdf);
    }

    // RV07: polygon elements fill their closed points instead of diagnosing.
    public static void PptxSyntheticSvgPolygonRendersClosedPoints()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <polygon points="10,10 90,10 50,40" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("86.4 453.6 m", pdf);
        TestAssert.Contains("144 410.4 l", pdf);
    }

    // RV07: rounded rect corners render beziers instead of diagnosing.
    public static void PptxSyntheticSvgRoundedRectRendersCorners()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <rect x="10" y="10" width="80" height="30" rx="5" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("93.6 453.6 m", pdf);
        TestAssert.Contains(" c\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: evenodd fill rules punch holes instead of filling them.
    public static void PptxSyntheticSvgEvenOddFillRulePunchesHole()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0H100V50H0Z M25 10H75V40H25 10Z" fill="#FF0000" fill-rule="evenodd"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\nf*\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: explicit nonzero fill rules keep the default operator (guard).
    public static void PptxSyntheticSvgNonzeroFillRuleKeepsOperator()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0H100V50H0Z" fill="#FF0000" fill-rule="nonzero"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\nf\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07-E2: source coordinate units must not determine printed gradient quality.
    public static void PptxSvgRadialSamplingIsInvariantToSourceCoordinateScale()
    {
        foreach (bool userSpace in new[] { false, true })
        {
            byte[]? baseline = null;
            foreach (int width in new[] { 10, 100, 1000 })
            {
                string units = userSpace
                    ? FormattableString.Invariant($"gradientUnits=\"userSpaceOnUse\" cx=\"{width / 2d}\" cy=\"{width / 4d}\" r=\"{width / 5d}\"")
                    : "cx=\"0.5\" cy=\"0.5\" r=\"0.2\"";
                string input = WriteSvgGradientDeck($"""
                    <svg viewBox="0 0 {width} {width / 2d}" xmlns="http://www.w3.org/2000/svg">
                      <defs><radialGradient id="g" {units}><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs>
                      <path d="M0 0H{width}V{width / 2d}H0Z" fill="url(#g)"/>
                    </svg>
                    """);
                string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
                var diagnostics = new List<OoxPdfDiagnostic>();
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
                byte[] pdf = File.ReadAllBytes(output);
                TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" || d.Id == "PPTX_NODE_RENDER_FAILED"),
                    "Equivalent supported radial gradients should remain renderable.");
                if (baseline is null) { baseline = pdf; }
                else { TestAssert.True(baseline.AsSpan().SequenceEqual(pdf),
                    "Changing SVG coordinate units without changing printed geometry must preserve radial sampling and PDF bytes."); }
            }
        }
    }

    public static void PptxSyntheticSvgRadialGradientPaintsSmoothShading()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <defs><radialGradient id="g" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs>
              <path d="M0 0H100V50H0Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/ShadingType 3", pdf);
        TestAssert.Contains("/C0 [1 0 0]", pdf);
        TestAssert.Contains("/C1 [0 0 1]", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable", StringComparison.Ordinal)), "Resolvable radial gradients must render.");
    }

    public static void PptxSvgIgnoredContainerAndStopOpacityDiagnosesOnlyUsedPaint()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg" opacity="0.5">
              <defs>
                <g opacity="0.5"><path d="M0 0H5V5H0Z"/></g>
                <radialGradient id="unused"><stop offset="0" stop-color="#FF0000" stop-opacity="0.5"/><stop offset="1" stop-color="#0000FF"/></radialGradient>
                <radialGradient id="g"><stop offset="0" stop-color="#FF0000" stop-opacity="50%"/><stop offset="1" stop-color="#0000FF" style="stop-opacity:0.5"/></radialGradient>
              </defs>
              <g opacity="1" style="opacity:0.25"><path d="M0 0H100V50H0Z" fill="url(#g)"/><path d="M0 0H100V50H0Z" fill="url(#g)"/></g>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("ignores opacity on 2 containers", StringComparison.Ordinal)));
        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("ignores stop-opacity on 2 gradient stops", StringComparison.Ordinal)));
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Ignored opacity must preserve the picture.");

        string control = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg" opacity="0.5" style="opacity:1">
              <defs><radialGradient id="unused"><stop offset="0" stop-color="#FF0000" stop-opacity="0.5"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs>
              <path d="M0 0H100V50H0Z" fill="#0000FF" opacity="0.5"/>
            </svg>
            """);
        diagnostics.Clear();
        OoxPdfConverter.Convert(control, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT"), "Unused definitions and effective unit container opacity must not warn.");
    }

    public static void PptxSvgRadialSpreadIsInvariantToSourceCoordinateScale()
    {
        foreach (string spread in new[] { "repeat", "reflect" })
        {
            byte[]? baseline = null;
            foreach (int width in new[] { 10, 100, 1000 })
            {
                string input = WriteSvgGradientDeck(FormattableString.Invariant($"""
                    <svg viewBox="0 0 {width} {width / 2d}" xmlns="http://www.w3.org/2000/svg">
                      <defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="{width / 2d}" cy="{width / 4d}" r="{width / 5d}" spreadMethod="{spread}"><stop offset="0" stop-color="#FF0000"/><stop offset="0.4" stop-color="#00FF00"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs>
                      <path d="M0 0H{width}V{width / 2d}H0Z" fill="url(#g)"/>
                    </svg>
                    """));
                string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
                var diagnostics = new List<OoxPdfDiagnostic>();
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
                byte[] pdf = File.ReadAllBytes(output);
                TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Repeated radial fills must remain renderable.");
                if (baseline is null) { baseline = pdf; }
                else { TestAssert.True(baseline.AsSpan().SequenceEqual(pdf),
                    "Source coordinate units must not alter the printed repeated gradient or PDF bytes."); }
            }
        }
    }

    public static void PptxSvgRadialUnrepresentableFunctionBoundsKeepPaintedFallback()
    {
        foreach (string offsets in new[] { "0.5,0.5", "0.5,0.500001", "0.000001,0.5", "0.5,0.999999" })
        {
            string[] pair = offsets.Split(',');
            string input = WriteSvgGradientDeck($"""
                <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
                  <defs><radialGradient id="g"><stop offset="{pair[0]}" stop-color="#FF0000"/><stop offset="{pair[1]}" stop-color="#0000FF"/></radialGradient></defs>
                  <path d="M0 0H100V50H0Z" fill="url(#g)"/>
                </svg>
                """);
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
            string pdf = File.ReadAllText(output, Encoding.ASCII);
            TestAssert.DoesNotContain("/ShadingType 3", pdf);
            TestAssert.Contains("0 0 1 rg", pdf);
            TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"),
                "Hard stops and sub-precision intervals must retain their painted fill.");
        }
    }

    public static void PptxSvgRadialSpreadExpansionRetainsBoundedFallback()
    {
        foreach ((double radius, int stopCount) in new[] { (0.1d, 2), (2d, 12), (1d, 4) })
        {
            string stops = string.Concat(Enumerable.Range(0, stopCount).Select(index =>
                FormattableString.Invariant($"<stop offset=\"{(stopCount == 4 ? 0.2d + 0.6d * index / (stopCount - 1) : (double)index / (stopCount - 1))}\" stop-color=\"#0000FF\"/>")));
            string input = WriteSvgGradientDeck(FormattableString.Invariant($"""
                <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
                  <defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="50" cy="25" r="{radius}" spreadMethod="repeat">{stops}</radialGradient></defs>
                  <path d="M0 0H100V50H0Z" fill="url(#g)"/>
                  <path d="M0 0H10V10H0Z" fill="#00FF00"/>
                </svg>
                """));
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
            byte[] bytes = File.ReadAllBytes(output);
            string pdf = Encoding.ASCII.GetString(bytes);
            TestAssert.True(bytes.Length < 50_000, "Dense repetition must not expand an unbounded PDF function graph.");
            TestAssert.DoesNotContain("/ShadingType 3", pdf);
            TestAssert.Contains("0 0 1 rg", pdf);
            TestAssert.Contains("0 1 0 rg", pdf);
            TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Bounded sampling must preserve neighboring paint.");
        }
    }

    // RV07-E1: pad extends the last stop beyond the outer circle, through the clipped shape.
    public static void PptxSyntheticSvgRadialPadPaintsCorners()
    {
        foreach (string units in new[] { "cx=\"50%\" cy=\"50%\" r=\"20%\"", "gradientUnits=\"userSpaceOnUse\" cx=\"50\" cy=\"25\" r=\"10\"" })
        {
            string input = WriteSvgGradientDeck($"""
                <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
                  <defs><radialGradient id="g" {units}><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs>
                  <path d="M0 0H100V50H0Z" fill="url(#g)"/>
                </svg>
                """);
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            OoxPdfConverter.Convert(input, output);
            string pdf = File.ReadAllText(output, Encoding.ASCII).Replace("\r\n", "\n");

            // The path clip and extended final stop cover corners beyond the outer circle.
            TestAssert.Contains("72 468 m\n216 468 l\n216 396 l\n72 396 l\nh\nW n\n", pdf);
            TestAssert.Contains("/Extend [true true]", pdf);
            TestAssert.Contains("/C0 [1 0 0]", pdf);
            TestAssert.Contains("/C1 [0 0 1]", pdf);
        }
    }

    // RV07-F1: PowerPoint 16 exports focal variants centered. Keep that Office
    // appearance while reporting the unsupported SVG focus to callers.
    public static void PptxSyntheticSvgRadialFocalFallbackMatchesOfficeCenter()
    {
        (byte[] Pdf, List<OoxPdfDiagnostic> Diagnostics) Render(string focalAttributes)
        {
            string input = WriteSvgGradientDeck($"""
                <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
                  <defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="50" cy="25" r="20" {focalAttributes}><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs>
                  <path d="M0 0H100V50H0Z" fill="url(#g)"/>
                </svg>
                """);
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
            return (File.ReadAllBytes(output), diagnostics);
        }

        var centered = Render("fx=\"50\" fy=\"25\"");
        TestAssert.True(!centered.Diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT"), "A centered supported gradient does not warn.");
        foreach (string focal in new[] { "fx=\"25\" fy=\"15\"", "fx=\"99\" fy=\"15\"", "fx=\"25\" fy=\"15\" fr=\"0\"" })
        {
            var candidate = Render(focal);
            TestAssert.True(centered.Pdf.AsSpan().SequenceEqual(candidate.Pdf), "Focal fallback preserves the Office-centered PDF bytes.");
            TestAssert.Equal(1, candidate.Diagnostics.Count(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("ignores focal points", StringComparison.Ordinal)));
            TestAssert.True(!candidate.Diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Focal fallback keeps the picture.");
        }
    }

    // RV07: garbage gradient vectors diagnose instead of killing conversion.
    public static void PptxSyntheticSvgGarbageGradientVectorDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <defs><linearGradient id="g" x1="bogus" x2="1"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></linearGradient></defs>
              <path d="M0 0H100V50H0Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable", StringComparison.Ordinal)), "Garbage gradient vectors must diagnose.");
    }

    // RV07: percentage gradient vectors resolve as fractions in bounding boxes.
    public static void PptxSyntheticSvgPercentGradientVectorRenders()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <defs><linearGradient id="g" x1="0%" y1="0" x2="100%" y2="0"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></linearGradient></defs>
              <path d="M0 0H100V50H0Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" rg", pdf);
        TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable", StringComparison.Ordinal)), "Percentage vectors must resolve.");
    }

    // RV07: percentage vectors in user space diagnose instead of killing conversion.
    public static void PptxSyntheticSvgUserSpacePercentVectorDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <defs><linearGradient id="g" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="50%" y2="0"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></linearGradient></defs>
              <path d="M0 0H100V50H0Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable", StringComparison.Ordinal)), "User-space percentages must diagnose.");
    }

    // RV07: garbage opacities fall back to opaque instead of killing conversion.
    public static void PptxSyntheticSvgGarbageOpacityFallsBack()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0H100V50H0Z" fill="#FF0000" opacity="bogus%"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    // RV07: garbage stop offsets drop the stop instead of killing conversion.
    public static void PptxSyntheticSvgGarbageStopOffsetDropsStop()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#FF0000"/><stop offset="bogus" stop-color="#0000FF"/></linearGradient></defs>
              <path d="M0 0H100V50H0Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    // RV07: dashed strokes render their pattern instead of solid fallback.
    public static void PptxSyntheticSvgDashedStrokeRendersDashPattern()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-dasharray="4 2"/>
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-dasharray="3"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("[5.76 2.88 ] 0 d", pdf);
        TestAssert.Contains("[4.32 4.32 ] 0 d", pdf);
    }

    // RV07: round line caps map to PDF caps.
    public static void PptxSyntheticSvgRoundLineCapRenders()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-linecap="round"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\n1 J\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: round line joins map to PDF joins.
    public static void PptxSyntheticSvgRoundLineJoinRenders()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-linejoin="round"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\n1 j\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: dash offsets shift the pattern phase in points.
    public static void PptxSyntheticSvgDashOffsetShiftsPhase()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-dasharray="4 2" stroke-dashoffset="1"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("[5.76 2.88 ] 1.44 d", pdf);
    }

    // RV07: invalid cap values diagnose and render butt instead of vanishing.
    public static void PptxSyntheticSvgInvalidStrokePresentationDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-linecap="bogus"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("default stroke effects", StringComparison.Ordinal)), "Invalid stroke presentation must diagnose.");
    }

    // RV07: scaled strokes without vector-effect keep the element-transform scale.
    public static void PptxSyntheticSvgScaledStrokeWithoutVectorEffectKeepsTransformScale()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H25 V12.5 H0 Z" fill="none" stroke="#FF0000" stroke-width="2" transform="scale(4)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("11.52 w", pdf);
    }

    // RV07: non-scaling-stroke keeps the untransformed stroke width.
    public static void PptxSyntheticSvgNonScalingStrokeIgnoresElementTransform()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H25 V12.5 H0 Z" fill="none" stroke="#FF0000" stroke-width="2" transform="scale(4)" vector-effect="non-scaling-stroke"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("2.88 w", pdf);
        TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("vector-effect", StringComparison.Ordinal)), "Honored non-scaling-stroke must not diagnose.");
    }

    // RV07: unknown vector-effect values diagnose while the scaled stroke still renders.
    public static void PptxSyntheticSvgUnknownVectorEffectDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" vector-effect="bogus-effect"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("vector-effect", StringComparison.Ordinal)), "Unknown vector effects must diagnose.");
    }

    // RV07: out-of-range path coordinates keep the sibling picture instead of dropping it.
    public static void PptxSyntheticSvgHugePathCoordinateKeepsSiblingPicture()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H1e999 V50 H0 Z" fill="#0000FF"/>
              <path d="M50 0 H100 V50 H50 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range coordinates must not drop the whole SVG picture.");
    }

    // RV07: out-of-range viewBox dimensions diagnose as unusable instead of dropping the picture.
    public static void PptxSyntheticSvgHugeViewBoxDiagnosesUnusable()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 1e999 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("no usable viewBox", StringComparison.Ordinal)), "Out-of-range viewBox must diagnose as unusable.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range viewBox must not drop the picture through node recovery.");
    }

    // RV07: subnormal viewBox dimensions overflow the picture scale and diagnose as unusable.
    public static void PptxSyntheticSvgTinyViewBoxDiagnosesUnusable()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 1e-320 1e-320" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("no usable viewBox", StringComparison.Ordinal)), "Overflowing picture scale must diagnose as unusable.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Overflowing picture scale must not drop the picture through node recovery.");
    }

    // RV07: out-of-range stroke widths diagnose as unparsable while the fill still paints.
    public static void PptxSyntheticSvgHugeStrokeWidthDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="#FF0000" stroke="#FF0000" stroke-width="1e999"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unparsable stroke paint", StringComparison.Ordinal)), "Out-of-range stroke width must diagnose.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range stroke width must not drop the whole SVG picture.");
    }

    // RV07: out-of-range dash lengths fall back to a solid stroke with a diagnosis.
    public static void PptxSyntheticSvgHugeDashPatternFallsBackToSolid()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-dasharray="1e999 2"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("default stroke effects", StringComparison.Ordinal)), "Out-of-range dash pattern must diagnose.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range dash pattern must not drop the whole SVG picture.");
    }

    // RV07: out-of-range miter limits diagnose and render the default.
    public static void PptxSyntheticSvgHugeMiterLimitDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-miterlimit="1e999"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("\n4 M\n", pdf.Replace("\r\n", "\n"));
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("default stroke effects", StringComparison.Ordinal)), "Out-of-range miter limit must diagnose.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range miter limit must not drop the whole SVG picture.");
    }

    // RV07: out-of-range dash offsets fall back to a solid stroke with a diagnosis.
    public static void PptxSyntheticSvgHugeDashOffsetDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-dasharray="4 2" stroke-dashoffset="1e999"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("default stroke effects", StringComparison.Ordinal)), "Out-of-range dash offset must diagnose.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range dash offset must not drop the whole SVG picture.");
    }

    // RV07: out-of-range gradient vectors skip the gradient so referencing paths diagnose.
    public static void PptxSyntheticSvgHugeGradientVectorDiagnosesUnresolvable()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="1e999" y2="0">
                <stop offset="0" stop-color="#0000FF"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M0 0 H100 V50 H0 Z" fill="url(#g)"/>
              <path d="M50 0 H100 V50 H50 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable gradient g", StringComparison.Ordinal)), "Out-of-range gradient vectors must diagnose as unresolvable.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range gradient vectors must not drop the whole SVG picture.");
    }

    // RV07: out-of-range radial radii skip the gradient so referencing paths diagnose.
    public static void PptxSyntheticSvgHugeRadialRadiusDiagnosesUnresolvable()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <radialGradient id="r" cx="0.5" cy="0.5" r="1e999">
                <stop offset="0" stop-color="#0000FF"/>
                <stop offset="1" stop-color="#0000FF"/>
              </radialGradient>
              <path d="M0 0 H100 V50 H0 Z" fill="url(#r)"/>
              <path d="M50 0 H100 V50 H50 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable gradient r", StringComparison.Ordinal)), "Out-of-range radial radii must diagnose as unresolvable.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range radial radii must not drop the whole SVG picture.");
    }

    // RV07: transforms composing past double range keep the sibling picture instead of dropping it.
    public static void PptxSyntheticSvgHugeTransformKeepsSiblingPicture()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H10 V5 H0 Z" fill="#0000FF" transform="scale(1e308)"/>
              <path d="M50 0 H100 V50 H50 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Overflowing transforms must not drop the whole SVG picture.");
    }

    // RV07: out-of-range shape geometry renders nothing instead of misdiagnosing a command.
    public static void PptxSyntheticSvgHugeShapeGeometrySkipsShape()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <rect x="10" y="10" width="1e999" height="30" fill="#0000FF"/>
              <path d="M50 0 H100 V50 H50 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unsupported I command", StringComparison.Ordinal)), "Out-of-range shape geometry must not misdiagnose a path command.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Out-of-range shape geometry must not drop the whole SVG picture.");
    }
    // RV07-N1: composed gradient overflow omits only the fill and keeps adjacent
    // shapes and the invalid fill's independently usable solid stroke.
    public static void PptxSyntheticSvgExtremeGradientGeometryKeepsOtherShapes()
    {
        foreach (var probe in new[]
        {
            (Gradient: "<linearGradient id=\"g\"><stop offset=\"0\" stop-color=\"#00FF00\"/><stop offset=\"1\" stop-color=\"#0000FF\"/></linearGradient>", Path: "M-1e308 0H1e308V20H-1e308Z", Transform: ""),
            (Gradient: "<radialGradient id=\"g\" cx=\"1e308\" cy=\"0.5\" r=\"1e308\"><stop offset=\"0\" stop-color=\"#00FF00\"/><stop offset=\"1\" stop-color=\"#0000FF\"/></radialGradient>", Path: "M0 0H10V5H0Z", Transform: ""),
            (Gradient: "<linearGradient id=\"g\" gradientUnits=\"userSpaceOnUse\" x1=\"1e308\" y1=\"0\" x2=\"1e308\" y2=\"1\"><stop offset=\"0\" stop-color=\"#00FF00\"/><stop offset=\"1\" stop-color=\"#0000FF\"/></linearGradient>", Path: "M0 0H10V5H0Z", Transform: "transform=\"scale(2)\""),
            (Gradient: "<linearGradient id=\"g\" gradientUnits=\"userSpaceOnUse\" x1=\"0\" y1=\"0\" x2=\"1e155\" y2=\"0\"><stop offset=\"0\" stop-color=\"#00FF00\"/><stop offset=\"1\" stop-color=\"#0000FF\"/></linearGradient>", Path: "M0 0H10V5H0Z", Transform: ""),
            (Gradient: "<linearGradient id=\"g\" gradientUnits=\"userSpaceOnUse\" x1=\"0\" y1=\"0\" x2=\"1e154\" y2=\"0\" spreadMethod=\"repeat\"><stop offset=\"0\" stop-color=\"#00FF00\"/><stop offset=\"1\" stop-color=\"#0000FF\"/></linearGradient>", Path: "M1e160 0H2e160V20H1e160Z", Transform: ""),
        })
        {
            string input = WriteSvgGradientDeck($"""
                <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
                  <defs>{probe.Gradient}</defs>
                  <path d="{probe.Path}" fill="url(#g)" stroke="#FF00FF" {probe.Transform}/>
                  <path d="M50 0H100V50H50Z" fill="#FF0000"/>
                </svg>
                """);
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
            string pdf = File.ReadAllText(output, Encoding.ASCII);
            TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Composed gradient overflow must not discard the entire SVG picture.");
            TestAssert.Contains("1 0 0 rg", pdf);
            TestAssert.Contains("1 0 1 RG", pdf);
            TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("non-finite gradient geometry", StringComparison.Ordinal)), "Composed gradient overflow must diagnose its omitted fill.");
        }
    }

    // RV07: stroke state overflowing past double range drops the stroke with a diagnosis, not the picture.
    public static void PptxSyntheticSvgHugeTransformStrokeStateDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H10 V5 H0 Z" fill="#FF0000" stroke="#FF0000" transform="scale(1e308)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unparsable stroke paint", StringComparison.Ordinal)), "Overflowing stroke state must diagnose.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Overflowing stroke state must not drop the whole SVG picture.");
    }
    // RV07: miter joins emit the SVG default miter limit instead of the PDF default.
    public static void PptxSyntheticSvgMiterJoinEmitsDefaultMiterLimit()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\n4 M\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: explicit miter limits map to PDF miter limits.
    public static void PptxSyntheticSvgExplicitMiterLimitMaps()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-miterlimit="2"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("\n2 M\n", pdf.Replace("\r\n", "\n"));
    }

    // RV07: invalid miter limits diagnose and render the default.
    public static void PptxSyntheticSvgInvalidMiterLimitDiagnoses()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0 H100 V50 H0 Z" fill="none" stroke="#FF0000" stroke-miterlimit="bogus"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("\n4 M\n", pdf.Replace("\r\n", "\n"));
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("miter", StringComparison.Ordinal)), "Invalid miter limits must diagnose.");
    }

    // RV07: repeating gradients wrap strip offsets instead of clamping.
    public static void PptxSyntheticSvgRepeatSpreadWrapsStripOffsets()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="50" y2="0" gradientUnits="userSpaceOnUse" spreadMethod="repeat">
                <stop offset="0" stop-color="#FF0000"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M60 0 H100 V20 H60 Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        // 1.3 wraps to 0.30000000000000004 in floating point, hence the trailing 2.
        TestAssert.Contains("0.698 0 0.302 rg", pdf);
    }

    // RV07: reflecting gradients mirror strip offsets instead of clamping.
    public static void PptxSyntheticSvgReflectSpreadMirrorsStripOffsets()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="50" y2="0" gradientUnits="userSpaceOnUse" spreadMethod="reflect">
                <stop offset="0" stop-color="#FF0000"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M60 0 H100 V20 H60 Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.298 0 0.698 rg", pdf);
    }

    // RV07: a missing stop-color defaults to black (PowerPoint normalizes black stops by
    // dropping the attribute).
    public static void PptxSyntheticSvgMissingStopColorDefaultsToBlack()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="1" y2="0">
                <stop offset="0" stop-color="#FFFFFF"/>
                <stop offset="1"/>
              </linearGradient>
              <path d="M0 0 H100 V100 H0 Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 g", pdf);
        TestAssert.Contains("0.012 g", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "SVG_UNSUPPORTED_CONTENT"), "Defaulted black stops must not diagnose.");
    }

    // RV07: gradient transforms shift the sampled vector.
    public static void PptxSyntheticSvgGradientTransformShiftsVector()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="50" y2="0" gradientUnits="userSpaceOnUse" gradientTransform="translate(50,0)">
                <stop offset="0" stop-color="#FF0000"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M60 0 H100 V20 H60 Z" fill="url(#g)"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.78 0 0.22 rg", pdf);
        TestAssert.Contains("0.02 0 0.98 rg", pdf);
    }

    // RV07: unparseable gradient transforms resolve to no gradient with a diagnostic.
    public static void PptxSyntheticSvgBadGradientTransformResolvesNoGradient()
    {
        string input = WriteSvgGradientDeck("""
            <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <linearGradient id="g" x1="0" y1="0" x2="50" y2="0" gradientUnits="userSpaceOnUse" gradientTransform="bogus(1)">
                <stop offset="0" stop-color="#FF0000"/>
                <stop offset="1" stop-color="#0000FF"/>
              </linearGradient>
              <path d="M60 0 H100 V20 H60 Z" fill="url(#g)"/>
              <path d="M0 0 H20 V20 H0 Z" fill="#FF0000"/>
            </svg>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable gradient g", StringComparison.Ordinal)), "Bad gradient transforms must diagnose.");
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    public static void PptxSyntheticPngPictureAppliesLuminanceRecolor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"><a:lum bright="70000" contrast="-70000"/></a:blip></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [32, 64, 96])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(PptxTests.CountOccurrences(pdf, "/Subtype /Image") >= 2, "The same image part with and without recolor must use distinct cached image XObjects.");
        List<byte[]> imageRgbStreams = PptxTests.ReadPdfDeviceRgbImageStreams(output, width: 1, height: 1);
        TestAssert.True(imageRgbStreams.Any(rgb => rgb.SequenceEqual(new byte[] { 32, 64, 96 })), "Expected the original image RGB stream to remain unchanged.");
        TestAssert.True(imageRgbStreams.Any(rgb => rgb.SequenceEqual(new byte[] { 188, 198, 207 })), "Expected OOXML luminance contrast to scale channels before brightness lifts them toward white.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported PNG luminance recolor should not emit unsupported diagnostics.");
    }

    public static void PptxSyntheticPngPictureAppliesDuotoneRecolor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"><a:duotone><a:srgbClr val="000000"/><a:prstClr val="white"/></a:duotone></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [128, 128, 128])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported PNG duotone recolor should not emit unsupported diagnostics.");
    }

    public static void PptxSyntheticJpegPictureDuotoneRecolorEmitsDiagnostic()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x03,
            0x00, 0x05,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="jpeg" ContentType="image/jpeg"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill>
                      <a:blip r:embed="rId1">
                        <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone>
                        <a:alphaModFix amt="50000"/>
                      </a:blip>
                    </p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.jpeg"] = jpegHeader
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/DCTDecode", pdf);
        TestAssert.True(diagnostics.Any(d =>
            d.Id == "PPTX_UNSUPPORTED_IMAGE_RECOLOR" &&
            d.Feature == "image recolor" &&
            d.Fallback == "Original image" &&
            d.Message.Contains("duotone", StringComparison.Ordinal) &&
            d.Message.Contains("image/jpeg", StringComparison.Ordinal) &&
            d.Message.Contains("baseline DCT", StringComparison.Ordinal)), "JPEG duotone recolor should keep a public diagnostic until a structural JPEG/PDF recolor path exists.");
    }

    public static void PptxSyntheticBaselineJpegPictureAppliesDuotoneRecolor()
    {
        byte[] jpeg = Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAIDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD+Rb4g/wDI++N/+xv8S/8Ap5vaKKK/7o/An/kyHg3/ANmq8PP/AFkcoPyHxn/5PD4r/wDZyuOv/WozQ//Z");
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="jpeg" ContentType="image/jpeg"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill>
                      <a:blip r:embed="rId1">
                        <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone>
                      </a:blip>
                    </p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.jpeg"] = jpeg
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.True(!pdf.Contains("/DCTDecode", StringComparison.Ordinal), "Recolored JPEG output should be materialized as decoded RGB image data.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported baseline JPEG duotone recolor should not emit unsupported diagnostics.");
    }

    public static void PptxSyntheticPngPictureAppliesGrayAndBilevelRecolor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"><a:grayscl/></a:blip></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"><a:biLevel thresh="60000"/></a:blip></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [128, 16, 240])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(PptxTests.CountOccurrences(pdf, "/Subtype /Image") >= 2, "Distinct grayscale and bi-level recolors should use distinct cached image XObjects.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported PNG gray/bi-level recolor should not emit unsupported diagnostics.");
    }

    public static void PptxSyntheticShapePictureFillRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:blipFill><a:blip r:embed="rId1"/></a:blipFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticRgbaPngPicturePreservesSoftMask()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbaPng(2, 1, [255, 0, 0, 255, 0, 0, 255, 64])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/SMask", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticEllipsePictureFillClipsImage()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="ellipse"/>
                      <a:blipFill><a:blip r:embed="rId1"/></a:blipFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("W n", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticBmpPictureRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="bmp" ContentType="image/bmp"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.bmp"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.bmp"] = TestFixtures.CreateRgbBmp(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "IMAGE_UNSUPPORTED_FORMAT"), "BMP should render instead of emitting unsupported image diagnostics.");
    }

    public static void PptxSyntheticBmpPictureAppliesDuotoneRecolor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="bmp" ContentType="image/bmp"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.bmp"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"><a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone></a:blip></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.bmp"] = TestFixtures.CreateRgbBmp(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(PptxTests.CountOccurrences(pdf, "/Subtype /Image") >= 2, "The same BMP image part with and without recolor must use distinct cached image XObjects.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported BMP duotone recolor should not emit unsupported diagnostics.");
        TestAssert.True(diagnostics.All(d => d.Id != "IMAGE_UNSUPPORTED_FORMAT"), "Supported BMP duotone recolor should render instead of emitting unsupported image diagnostics.");
    }

    public static void PptxUnsupportedPngImageEmitsDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateUnsupportedHighBitDepthPng()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        TestAssert.True(File.Exists(output), "Unsupported image should not fail the whole conversion.");
        TestAssert.True(collector.Diagnostics.Any(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT" && d.Severity == OoxPdfSeverity.Error), "Unsupported image should emit a release-blocking diagnostic.");
    }

    public static void PptxSyntheticCroppedPictureUsesClipping()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/><a:srcRect l="25000" r="25000"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 396 144 72 re W* n", pdf);
        TestAssert.DoesNotContain(" re W n", pdf);
        TestAssert.Contains("144 0 0 72 72 396 cm", pdf);
        TestAssert.DoesNotContain("288 0 0 72 0 396 cm", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    // RV18: numbers trailing a close used to hang the path walk because Z
    // consumed no tokens. The shared parser stops there instead, so the
    // conversion terminates and the supported prefix still paints.
    public static void PptxSyntheticSvgPathNumbersAfterCloseTerminate()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="svg" ContentType="image/svg+xml"/>
                <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8("""
                <svg viewBox="0 0 20 10" xmlns="http://www.w3.org/2000/svg">
                <path d="M0 0 L20 0 L20 10 L0 10 Z 5" fill="#112233"/>
                </svg>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.067 0.133 0.2 rg", pdf);
        TestAssert.Contains(" m", pdf);
        TestAssert.DoesNotContain("/Subtype /Image", pdf);
    }

    // RV08: a dangling image relationship errors visibly instead of failing the conversion.
    public static void PptxMissingImagePartDiagnosesErrorAndContinues()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        OoxPdfDiagnostic[] matches = diagnostics.Where(d => d.Id == "IMAGE_MISSING_PART").ToArray();
        TestAssert.Equal(1, matches.Length);
        TestAssert.Equal(OoxPdfSeverity.Error, matches[0].Severity);
        TestAssert.Equal("Ignored", matches[0].Fallback);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.DoesNotContain("/Subtype /Image", pdf);
    }
}
