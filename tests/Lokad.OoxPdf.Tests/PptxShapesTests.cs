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

internal static partial class PptxShapesTests
{
    public static void PptxSyntheticShapesProduceDrawingOperators()
    {
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="ellipse"/>
                          <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="line"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="4572000" y="3657600"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="roundRect"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="5943600" y="3657600"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:pattFill prst="dkDnDiag">
                            <a:fgClr><a:srgbClr val="2F856A"/></a:fgClr>
                            <a:bgClr><a:srgbClr val="EEEEEE"/></a:bgClr>
                          </a:pattFill>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string content = Regex.Replace(PptxTests.RenderSingleSlideContent(input), @"\s+", " ");
        TestAssert.Contains("1 g 0 0 720 540 re f* Q", content);
        TestAssert.Contains("1 0 0 rg 72 396 144 72 re f* 0 0 1 RG 2 w 72 396 144 72 re S Q", content);

        string ellipse = Regex.Match(content, @"(?:^| )0 1 0 rg (?<path>[^qQ]+) Q\b").Groups["path"].Value;
        TestAssert.True(ellipse.StartsWith("360 432 m ", StringComparison.Ordinal) &&
            ellipse.EndsWith("360 432 c h f*", StringComparison.Ordinal),
            "The green shape must close and fill an ellipse at its authored frame.");
        TestAssert.Equal(4, Regex.Matches(ellipse, " c ").Count);
        TestAssert.Contains("324 468 c ", ellipse);
        TestAssert.Contains("288 432 c ", ellipse);
        TestAssert.Contains("324 396 c ", ellipse);

        TestAssert.Contains("0 G 1 w 72 252 m 216 180 l S", content);
        string rounded = Regex.Match(content, @"(?:^| )0 G 1 w (?<path>360 240 m [^qQ]+) Q\b").Groups["path"].Value;
        TestAssert.True(rounded.EndsWith("360 192 c h S", StringComparison.Ordinal),
            "The rounded rectangle must close and stroke its own frame.");
        TestAssert.Equal(4, Regex.Matches(rounded, " c ").Count);
        TestAssert.Contains("0.933 g 468 180 72 72 re f* q 468 180 72 72 re W n 0.184 0.522 0.416 RG", content);
    }

    public static void PptxSyntheticFilledStrokedShapeRepeatsSlideClipBeforeFill()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII).Replace("\r\n", "\n", StringComparison.Ordinal);
        int shapeClip = pdf.IndexOf("0 0 720 540 re W* n\n1 0 0 rg", StringComparison.Ordinal);
        int fill = pdf.IndexOf("72 396 144 72 re f*", StringComparison.Ordinal);
        int strokeColor = pdf.IndexOf("0 0 1 RG", StringComparison.Ordinal);
        TestAssert.True(shapeClip >= 0, "Expected the repeated slide clip to precede the filled shape region.");
        TestAssert.True(fill > shapeClip, "Expected the shape fill after the repeated slide clip.");
        TestAssert.True(strokeColor > fill, "Expected the stroke state after the filled region.");
        TestAssert.True(pdf.IndexOf("re f*\n0 0 720 540 re W* n", fill, StringComparison.Ordinal) < 0, "Filled/stroked shapes should not repeat the slide clip between fill and stroke.");
    }

    public static void PptxSyntheticShapeExplicitLineColorInheritsStyleLineWidth()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="LineWidthTheme">
                  <a:themeElements>
                    <a:clrScheme name="LineWidthTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
                    </a:clrScheme>
                    <a:fontScheme name="LineWidthTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
                    <a:fmtScheme name="LineWidthTheme">
                      <a:fillStyleLst/>
                      <a:lnStyleLst>
                        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                      </a:lnStyleLst>
                      <a:effectStyleLst/>
                      <a:bgFillStyleLst/>
                    </a:fmtScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:style>
                          <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
                          <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
                          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
                          <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
                        </p:style>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:noFill/>
                          <a:ln><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode shapeNode = scene.Slides[0].SlideNodes[0];
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(shapeNode.Shape?.NoFill == true, "Expected explicit shape noFill provenance in the scene model.");
        TestAssert.True(shapeNode.Shape?.Fill.HasFill == false, "Expected explicit shape noFill to suppress scene solid fill resolution.");
        TestAssert.Equal(0, shapeNode.Shape?.FillReference.Index ?? -1);
        TestAssert.True(shapeNode.Shape?.FillReference.Reference is not null, "Expected fillRef provenance even when idx=0 does not resolve a theme style.");
        TestAssert.True(shapeNode.Shape?.FillReference.Style is null, "Expected idx=0 fillRef to remain unresolved against the theme fill style list.");
        TestAssert.Equal(2, shapeNode.Shape?.LineReference.Index ?? 0);
        TestAssert.True(shapeNode.Shape?.LineReference.Reference is not null, "Expected line reference XML ownership in the shape scene model.");
        TestAssert.True(shapeNode.Shape?.LineReference.Style is not null, "Expected line reference to resolve through the theme line style list.");
        TestAssert.Contains("1.5 w", pdf);
        TestAssert.Contains("0.753 0 0 RG", pdf);
        TestAssert.Contains("72 396 144 72 re S", pdf);
    }

    public static void PptxSyntheticShapeExplicitLineWithoutWidthUsesOfficeDefault()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:noFill/>
                          <a:ln><a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(scene.Slides[0].SlideNodes[0].Shape?.Line.HasLine == true, "Expected explicit <a:ln> to produce a scene line.");
        TestAssert.Equal(0.75d, scene.Slides[0].SlideNodes[0].Shape?.Line.Width ?? 0d);
        TestAssert.Contains("0.75 w", pdf);
        TestAssert.Contains("0.2 0.4 0.6 RG", pdf);
        TestAssert.Contains("72 396 144 72 re S", pdf);
    }

    public static void PptxSyntheticConnectorExplicitLineWidthInheritsStyleLineColor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="LineCascadeTheme">
                  <a:themeElements>
                    <a:clrScheme name="LineCascadeTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
                    </a:clrScheme>
                    <a:fontScheme name="LineCascadeTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
                    <a:fmtScheme name="LineCascadeTheme">
                      <a:fillStyleLst/>
                      <a:lnStyleLst>
                        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                      </a:lnStyleLst>
                      <a:effectStyleLst/>
                      <a:bgFillStyleLst/>
                    </a:fmtScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:style>
                          <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
                          <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
                          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
                          <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
                        </p:style>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="0" cy="1828800"/></a:xfrm>
                          <a:prstGeom prst="line"/>
                          <a:ln w="12700"/>
                        </p:spPr>
                      </p:cxnSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(scene.Slides[0].SlideNodes[0].Shape?.Line.HasLine == true, "Expected style line color to complete the explicit connector line.");
        TestAssert.Equal(1d, scene.Slides[0].SlideNodes[0].Shape?.Line.Width ?? 0d);
        TestAssert.Equal(new RgbColor(68, 114, 196), scene.Slides[0].SlideNodes[0].Shape?.Line.Color ?? default);
        TestAssert.Contains("1 w", pdf);
        TestAssert.Contains("0.267 0.447 0.769 RG", pdf);
        TestAssert.Contains("72 468 m 72 324 l S", pdf);
    }

    public static void PptxSyntheticArrowAndConnectorShapesRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="1828800" cy="0"/></a:xfrm>
                          <a:prstGeom prst="line"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm flipV="1"><a:off x="2743200" y="914400"/><a:ext cx="0" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="straightConnector1"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="2F856A"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="downArrow"/>
                          <a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.133 g", pdf);
        TestAssert.Contains("72 324.5 m", pdf);
        TestAssert.Contains("212.667 323.5 l", pdf);
        TestAssert.Contains("216 324 m", pdf);
        TestAssert.Contains("212 325.5 l", pdf);
        TestAssert.Contains("212 322.5 l", pdf);
        TestAssert.Contains("1 -0 -0 -1 0 864 cm", pdf);
        TestAssert.Contains("216.5 468 m", pdf);
        TestAssert.Contains("216.5 400 l", pdf);
        TestAssert.Contains("215.5 468 l", pdf);
        TestAssert.Contains("216 396 m", pdf);
        TestAssert.Contains("219 402 l", pdf);
        TestAssert.Contains("213 402 l", pdf);
        TestAssert.DoesNotContain("216 468 m 216 396 l S", pdf);
        TestAssert.Contains("0.184 0.522 0.416 rg", pdf);
        TestAssert.Contains("0.753 0 0 rg", pdf);
        TestAssert.Contains("90 468 m", pdf);
        TestAssert.Contains("126 432 l", pdf);
        TestAssert.Contains("108 396 l", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
    }

    public static void PptxSyntheticTriangleConnectorsPaintTheirOfficeMarkerWidths()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      {{Connector("line", "3657600")}}
                      {{Connector("straightConnector1", "5486400")}}
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(input, ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var markers = Regex.Matches(pdf,
            @"(?m)^(?<tip>288|432) 468 m\r?\n(?<first>[\d.]+) 456 l\r?\n(?<second>[\d.]+) 456 l\r?\nh\r?\nf\r?$");
        TestAssert.Equal(2, markers.Count);
        foreach ((string tip, string left, string right) in new[] { ("288", "283.5", "292.5"), ("432", "426", "438") })
        {
            var marker = markers.Single(match => match.Groups["tip"].Value == tip);
            string[] edges = [marker.Groups["first"].Value, marker.Groups["second"].Value];
            TestAssert.True(edges.Order(StringComparer.Ordinal).SequenceEqual([left, right]),
                "Each filled triangle must retain its connector's authored marker width.");
        }

        static string Connector(string preset, string x) => $$"""
            <p:cxnSp>
              <p:spPr>
                <a:xfrm><a:off x="{{x}}" y="914400"/><a:ext cx="0" cy="914400"/></a:xfrm>
                <a:prstGeom prst="{{preset}}"/>
                <a:ln w="38100"><a:solidFill><a:srgbClr val="156082"/></a:solidFill><a:headEnd type="triangle" w="med" len="med"/></a:ln>
              </p:spPr>
            </p:cxnSp>
            """;
    }

    public static void PptxStraightConnectorStrokeRetainsAuthoredEndMarkers()
    {
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    {{Connector(1, dash: true, tail: "triangle")}}
                    {{Connector(2, dash: true, tail: "arrow")}}
                    {{Connector(3, round: true, tail: "triangle")}}
                    {{Connector(4, head: "diamond", tail: "triangle")}}
                    {{Connector(5, head: "diamond", tail: "arrow")}}
                    {{Connector(6, dash: true, head: "arrow", tail: "triangle")}}
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        using var output = new MemoryStream();
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            DiagnosticSink = diagnostics.Add
        });

        TestAssert.Equal(0, diagnostics.Count);
        string pdf = Encoding.ASCII.GetString(output.ToArray()).ReplaceLineEndings("\n");
        foreach (int y in new[] { 468, 396, 324, 252, 180, 108 })
            TestAssert.Contains(FormattableString.Invariant($"72 {y} m 216 {y} l S"), pdf);

        foreach (int y in new[] { 468, 324, 252, 108 })
            TestAssert.Contains(FormattableString.Invariant($"216 {y} m\n200 {y + 8} l\n200 {y - 8} l\nh\nf\n"), pdf);

        foreach ((string x, int y) in new[] { ("216.01", 396), ("216.01", 180), ("71.99", 108) })
        {
            string marker = FormattableString.Invariant($@"(?m)^\S+ \S+ m\n{Regex.Escape(x)} {y} l\n(?:[^\n]+ [lc]\n)+h\nf$");
            TestAssert.True(Regex.IsMatch(pdf, marker), $"Expected a filled arrow at ({x}, {y}).");
        }
        foreach (int y in new[] { 252, 180 })
            TestAssert.Contains(FormattableString.Invariant($"72 {y} m\n80 {y - 6.4d} l\n88 {y} l\n80 {y + 6.4d} l\nh\nf\n"), pdf);

        TestAssert.Equal(3, Regex.Matches(pdf, @"(?m)^\[[\d .]+\] 0 d$").Count);
        TestAssert.Equal(3, Regex.Matches(pdf, @"(?m)^\[\] 0 d$").Count);
        TestAssert.Contains("1 J\n1 j\n72 324 m 216 324 l S", pdf);
        TestAssert.Contains("0 J\n0 j", pdf);

        static string Connector(int row, bool dash = false, bool round = false, string? head = null, string? tail = null)
        {
            string cap = round ? """ cap="rnd" """ : "";
            string dashElement = dash ? """<a:prstDash val="dash"/>""" : "";
            string headElement = head is null ? "" : $$"""<a:headEnd type="{{head}}"/>""";
            string tailElement = tail is null ? "" : $$"""<a:tailEnd type="{{tail}}"/>""";
            return FormattableString.Invariant($$"""
                <p:cxnSp><p:spPr>
                  <a:xfrm><a:off x="914400" y="{{row * 914400}}"/><a:ext cx="1828800" cy="0"/></a:xfrm>
                  <a:prstGeom prst="straightConnector1"/>
                  <a:ln w="50800"{{cap}}><a:solidFill><a:srgbClr val="CC1133"/></a:solidFill>{{dashElement}}{{headElement}}{{tailElement}}</a:ln>
                </p:spPr></p:cxnSp>
                """);
        }
    }

    public static void PptxSyntheticCurvedConnectorRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="666666"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.4 g", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
        TestAssert.DoesNotContain("144 360 180 324 216 324 c" + Environment.NewLine + "S", pdf);
    }

    public static void PptxSyntheticCurvedConnector2RendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector2"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.867 g", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
        TestAssert.DoesNotContain("111.765 468 144 435.765 144 396 c" + Environment.NewLine + "S", pdf);
    }

    public static void PptxSyntheticCurvedConnector2LoopUsesQuarterTurnTangents()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm rot="5400000" flipH="1" flipV="1"><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector2"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector2"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("-0 1 -1 -0", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
        TestAssert.DoesNotContain("111.765 468 144 435.765 144 396 c" + Environment.NewLine + "S", pdf);
    }

    public static void PptxSyntheticCurvedConnector2TriangleTailUsesOfficeSeparateMarkerSubpath()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector2"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.867 g", pdf);
        TestAssert.True(Regex.Matches(pdf, " m" + Environment.NewLine).Count >= 2, "Expected separate body and triangle marker subpaths.");
        TestAssert.True(Regex.Matches(pdf, " l" + Environment.NewLine).Count >= 98, "Expected Office-like dense connector body sampling.");
        TestAssert.True(Regex.Matches(pdf, "h" + Environment.NewLine).Count >= 2, "Expected both connector body and triangle marker to be closed before filling.");
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
        TestAssert.DoesNotContain(" c" + Environment.NewLine + "S", pdf);
    }

    public static void PptxSyntheticCurvedConnectorArrowTailUsesFilledOutline()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector3"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="606060"/></a:solidFill><a:tailEnd type="arrow"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.376 g", pdf);
        TestAssert.Contains(" c" + Environment.NewLine, pdf);
        TestAssert.True(Regex.Matches(pdf, " c" + Environment.NewLine).Count >= 4, "Expected Office-like arrow-tail filled connector paths to retain four cubic segments.");
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
        TestAssert.DoesNotContain(" c" + Environment.NewLine + "S", pdf);
    }

    public static void PptxSyntheticCustomGeometryCubicPathRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:custGeom>
                            <a:pathLst>
                              <a:path w="21600" h="10800">
                                <a:moveTo><a:pt x="0" y="5400"/></a:moveTo>
                                <a:cubicBezTo>
                                  <a:pt x="5400" y="0"/>
                                  <a:pt x="16200" y="10800"/>
                                  <a:pt x="21600" y="5400"/>
                                </a:cubicBezTo>
                              </a:path>
                            </a:pathLst>
                          </a:custGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="008000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.502 0 RG", pdf);
        TestAssert.Contains("72 432 m", pdf);
        TestAssert.Contains("108 468 180 396 216 432 c", pdf);
        TestAssert.Contains("S", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CUSTOM_GEOMETRY"), "Renderable custom cubic geometry should not emit the unsupported diagnostic.");
    }

    public static void PptxSyntheticCustomGeometryOpenCubicStealthTailUsesFilledOutline()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:custGeom>
                            <a:pathLst>
                              <a:path w="21600" h="10800" fill="none">
                                <a:moveTo><a:pt x="0" y="5400"/></a:moveTo>
                                <a:cubicBezTo>
                                  <a:pt x="5400" y="0"/>
                                  <a:pt x="16200" y="10800"/>
                                  <a:pt x="21600" y="5400"/>
                                </a:cubicBezTo>
                              </a:path>
                            </a:pathLst>
                          </a:custGeom>
                          <a:noFill/>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="008000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.502 0 rg", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
        TestAssert.True(Regex.Matches(pdf, " l" + Environment.NewLine).Count >= 120, "Expected Office-like dense stealth-tail connector body sampling.");
        TestAssert.DoesNotContain(" c" + Environment.NewLine + "S", pdf);
    }

    public static void PptxCustomGeometryStrokeRetainsAuthoredEndMarkers()
    {
        const string head = """<a:headEnd type="diamond"/>""";
        const string tail = """<a:tailEnd type="triangle"/>""";
        const string close = "<a:close/>";
        const string secondSubpath = """
            <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
            <a:lnTo><a:pt x="1000" y="1000"/></a:lnTo>
            """;
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    {{Shape(1, head)}}
                    {{Shape(2, head + tail)}}
                    {{Shape(3, tail, dashed: true)}}
                    {{Shape(4, head + tail, suffix: close)}}
                    {{Shape(5, head + tail, suffix: secondSubpath)}}
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        using var output = new MemoryStream();
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            DiagnosticSink = diagnostics.Add
        });

        TestAssert.Equal(0, diagnostics.Count);
        string pdf = Encoding.ASCII.GetString(output.ToArray()).ReplaceLineEndings("\n");
        foreach (int row in new[] { 1, 2, 3 })
        {
            int start = 522 - 72 * row, end = 486 - 72 * row;
            TestAssert.Contains(FormattableString.Invariant($"72 {start} m\n108 {start} 180 {end} 216 {end} c\nS\n"), pdf);
            if (row <= 2)
                TestAssert.Contains(FormattableString.Invariant($"72 {start} m\n76 {start - 3.2} l\n80 {start} l\n76 {start + 3.2} l\nh\nf\n"), pdf);
            if (row >= 2)
                TestAssert.Contains(FormattableString.Invariant($"216 {end} m\n209 {end + 3.15} l\n209 {end - 3.15} l\nh\nf\n"), pdf);
        }
        TestAssert.Contains("216 198 c\nh\nS\n", pdf);
        TestAssert.Contains("216 126 c\n72 180 m\n", pdf);
        TestAssert.Equal(4, Regex.Matches(pdf, @"(?m)^f$").Count);
        TestAssert.Equal(1, Regex.Matches(pdf, @"(?m)^\[[\d .]+\] 0 d$").Count);
        TestAssert.Contains("[] 0 d\n", pdf);

        static string Shape(int row, string markers, bool dashed = false, string suffix = "")
        {
            string dash = dashed ? """<a:prstDash val="dash"/>""" : "";
            return FormattableString.Invariant($$"""
                <p:sp><p:spPr>
                  <a:xfrm><a:off x="914400" y="{{row * 914400}}"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                  <a:custGeom><a:pathLst><a:path w="21600" h="10800" fill="none">
                    <a:moveTo><a:pt x="0" y="2700"/></a:moveTo>
                    <a:cubicBezTo><a:pt x="5400" y="2700"/><a:pt x="16200" y="8100"/><a:pt x="21600" y="8100"/></a:cubicBezTo>
                    {{suffix}}
                  </a:path></a:pathLst></a:custGeom>
                  <a:noFill/>
                  <a:ln w="25400"><a:solidFill><a:srgbClr val="008000"/></a:solidFill>{{dash}}{{markers}}</a:ln>
                </p:spPr></p:sp>
                """);
        }
    }

    public static void PptxSyntheticCustomGeometryArcPathRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:custGeom>
                            <a:gdLst>
                              <a:gd name="rx" fmla="val 10800"/>
                              <a:gd name="ry" fmla="val 5400"/>
                              <a:gd name="start" fmla="val 0"/>
                              <a:gd name="sweep" fmla="val 10800000"/>
                            </a:gdLst>
                            <a:pathLst>
                              <a:path w="21600" h="10800" fill="none">
                                <a:moveTo><a:pt x="21600" y="5400"/></a:moveTo>
                                <a:arcTo wR="rx" hR="ry" stAng="start" swAng="sweep"/>
                              </a:path>
                            </a:pathLst>
                          </a:custGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="008000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII).ReplaceLineEndings("\n");
        TestAssert.True(Regex.IsMatch(pdf, @"(?m)^0 0\.502 0 RG\n0 0\.502 0 rg\n2 w\n216 432 m\n[^\n]+ 144 396 c\n[^\n]+ 72 432 c\nS$"),
            "The green arc must sweep below its chord in source coordinates, through (144, 396) in PDF coordinates.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CUSTOM_GEOMETRY"), "Renderable custom arc geometry should not emit the unsupported diagnostic.");
    }

    // RV18: fully unsupported custom geometry diagnoses and falls back to a
    // filled bounding rectangle; neither the scene path nor the XML fallback
    // emits custom path operators for it.
    public static void PptxSyntheticFullyUnsupportedCustomGeometryFallsBackToRectangle()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:custGeom>
                            <a:pathLst>
                              <a:path w="21600" h="10800">
                                <a:unknownPathCommand/>
                              </a:path>
                            </a:pathLst>
                          </a:custGeom>
                          <a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_CUSTOM_GEOMETRY"), "Unrenderable custom geometry must stay diagnosed.");
        TestAssert.Contains("1 0 1 rg", pdf);
        TestAssert.Contains("72 396 144 72 re f*", pdf);
        TestAssert.Contains("72 396 144 72 re S", pdf);
    }

    public static void PptxSyntheticPresetArcRendersArcInsteadOfRectangle()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 5400000"/></a:avLst></a:prstGeom>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="444444"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.267 G", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.DoesNotContain("72 396 72 72 re S", pdf);
    }

    public static void PptxSyntheticPresetArcConvertsVisualAnglesOnEllipse()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 2700000"/></a:avLst></a:prstGeom>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("216 432 m", pdf);
        TestAssert.Contains("176.199 399.801 c", pdf);
        TestAssert.DoesNotContain("194.912", pdf);
    }

    public static void PptxSyntheticPresetArcStealthTailUsesFilledOutline() =>
        CheckPresetArcStealthTail(
            """<a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 19500000"/>""",
            flipVertical: true, expectedTip: "185.841 461.297");

    public static void PptxSyntheticPresetArcUsesOfficeDefaultAdjustments() =>
        CheckPresetArcStealthTail(guides: "", flipVertical: false, expectedTip: "216 432");

    private static void CheckPresetArcStealthTail(string guides, bool flipVertical, string expectedTip)
    {
        string flip = flipVertical ? " flipV=\"1\"" : "";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm{{flip}}><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="arc"><a:avLst>{{guides}}</a:avLst></a:prstGeom>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="444444"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string content = PptxTests.RenderSingleSlideContent(input);
        // The ellipse starts at its top midpoint, with half the 1.5pt line width on either side.
        // The second closed contour is the four-vertex stealth tail at the end angle.
        const string line = @"[-\d.]+ [-\d.]+ l\s+";
        string outline = $@"(?m)^0\.267 g\s+144 468\.75 m\s+(?:{line})+144 467\.25 l\s+144 468\.75 l\s+h\s+"
            + $@"{Regex.Escape(expectedTip)} m\s+(?:{line}){{3}}h\s+f\*?(?:\s|$)";
        TestAssert.True(Regex.IsMatch(content, outline),
            $"Expected the filled arc band and the stealth tail with tip at {expectedTip}.");
        TestAssert.True(!Regex.IsMatch(content, @"(?:^|\s)[Ss](?:\s|$)"), "The arc and tail share one filled outline.");
    }

    public static void PptxSyntheticShapeRoundRectHonorsAdjustment()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
                      <a:noFill/>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 460.8 m", pdf);
        TestAssert.DoesNotContain("72 456.48 m", pdf);
    }

    public static void PptxThemeOutlineMatchesAnEquivalentLocalOutline()
    {
        const string slidePart = "ppt/slides/slide1.xml";
        const string themePart = "ppt/theme/theme1.xml";
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..",
            "Cases", "pptx-ladder-06-round-join-rect.pptx"));
        using var archive = ZipFile.OpenRead(fixture);
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using Stream source = entry.Open();
            using var copy = new MemoryStream();
            source.CopyTo(copy);
            parts.Add(entry.FullName, copy.ToArray());
        }

        XNamespace p = OoxNamespaces.PresentationNamespace;
        XNamespace a = OoxNamespaces.DrawingNamespace;
        XDocument slide = XDocument.Parse(Encoding.UTF8.GetString(parts[slidePart]));
        XDocument theme = XDocument.Parse(Encoding.UTF8.GetString(parts[themePart]));
        XElement shape = slide.Descendants(p + "sp").Single();
        XElement shapeProperties = shape.Element(p + "spPr")!;
        XElement outline = shapeProperties.Element(a + "ln")!;
        outline.SetAttributeValue("cap", "rnd");
        outline.Element(a + "prstDash")!.SetAttributeValue("val", "dash");
        int lineIndex = (int)shape.Element(p + "style")!.Element(a + "lnRef")!.Attribute("idx")!;
        theme.Descendants(a + "lnStyleLst").Single().Elements(a + "ln").ElementAt(lineIndex - 1)
            .ReplaceWith(new XElement(outline));
        parts[themePart] = TestFixtures.Utf8(theme.ToString(SaveOptions.DisableFormatting));

        byte[] direct = ConvertSlide();
        outline.Remove();
        byte[] inherited = ConvertSlide();
        string pdf = Encoding.ASCII.GetString(inherited);
        TestAssert.Contains("[96 72 ] 0 d", pdf);
        TestAssert.Contains("1 J", pdf);
        TestAssert.Contains("1 j", pdf);
        TestAssert.True(inherited.AsSpan().SequenceEqual(direct), "The theme reference and the equivalent direct outline paint identically.");

        shapeProperties.Element(a + "noFill")!.AddAfterSelf(new XElement(a + "ln", new XElement(a + "noFill")));
        string suppressed = Encoding.ASCII.GetString(ConvertSlide());
        TestAssert.True(!suppressed.Contains("24 w", StringComparison.Ordinal), "A direct noFill suppresses the themed stroke.");

        byte[] ConvertSlide()
        {
            parts[slidePart] = TestFixtures.Utf8(slide.ToString(SaveOptions.DisableFormatting));
            using MemoryStream input = TestFixtures.CreateZipPackage(parts);
            using var output = new MemoryStream();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });
            return output.ToArray();
        }
    }

    public static void PptxSyntheticShapeStrokeDashCapAndJoinRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/></a:xfrm>
                        <a:prstGeom prst="line"/>
                        <a:ln w="25400" cap="rnd">
                          <a:solidFill><a:srgbClr val="222222"/></a:solidFill>
                          <a:prstDash val="dashDot"/>
                        </a:ln>
                      </p:spPr>
                    </p:sp>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="triangle"/>
                        <a:noFill/>
                        <a:ln w="25400">
                          <a:solidFill><a:srgbClr val="445566"/></a:solidFill>
                          <a:bevel/>
                        </a:ln>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("[8 6 2 6 ] 0 d", pdf);
        TestAssert.Contains("1 J", pdf);
        TestAssert.Contains("1 j", pdf);
        TestAssert.Contains("[] 0 d", pdf);
        TestAssert.Contains("2 j", pdf);
    }

    public static void PptxSyntheticShapeStrokeDashPresetVariantsRender()
    {
        string[] dashValues = ["dot", "sysDot", "dash", "sysDash", "lgDash", "dashDot", "sysDashDot", "lgDashDot", "lgDashDotDot", "sysDashDotDot", "solid"];
        var shapes = new StringBuilder();
        for (int i = 0; i < dashValues.Length; i++)
        {
            int y = 914400 + i * 274320;
            shapes.Append(CultureInfo.InvariantCulture, $"""
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="{y}"/><a:ext cx="1828800" cy="0"/></a:xfrm>
                        <a:prstGeom prst="line"/>
                        <a:ln w="12700"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:prstDash val="{dashValues[i]}"/></a:ln>
                      </p:spPr>
                    </p:sp>

                """);
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                {{shapes}}
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("[1 2 ] 0 d", pdf);
        TestAssert.Contains("[4 3 ] 0 d", pdf);
        TestAssert.Contains("[8 3 ] 0 d", pdf);
        TestAssert.Contains("[4 3 1 3 ] 0 d", pdf);
        TestAssert.Contains("[8 3 1 3 ] 0 d", pdf);
        TestAssert.Contains("[8 3 1 3 1 3 ] 0 d", pdf);
        TestAssert.True(pdf.AsSpan().Count("[] 0 d") == dashValues.Length - 1, "Each dashed line should reset the dash pattern; solid must not set a dash pattern.");
    }

    public static void PptxSyntheticOuterShadowRendersOffsetShape()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        <a:effectLst><a:outerShdw dist="91440" dir="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/GS50000F100000S gs", pdf);
        TestAssert.Contains("79.2 396 72 72 re f*", pdf);
        TestAssert.Contains("72 396 72 72 re f*", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_EFFECT"), "Supported outer shadow should not emit an unsupported-effect diagnostic.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TRANSPARENCY"), "Rendered outer shadow alpha should not emit an unsupported-transparency diagnostic.");
    }

    public static void PptxSyntheticShapeRectGlowUsesRasterSoftMask()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        <a:effectLst><a:glow rad="91440"><a:srgbClr val="0000FF"><a:alpha val="25000"/></a:srgbClr></a:glow></a:effectLst>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/GS25000F100000S gs", pdf);
        TestAssert.True(!pdf.Contains("64.8 388.8 86.4 86.4 re f*", StringComparison.Ordinal), "Glow must not emit an expanded solid rectangle.");
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/SMask ", pdf);
        TestAssert.Contains("Do", pdf);
        TestAssert.Contains("72 396 72 72 re f*", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_EFFECT"), "Rendered rectangular glow should not emit an unsupported-effect diagnostic.");
    }

    public static void PptxSyntheticLinearGradientShapeUsesPdfShading()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:gradFill>
                          <a:gsLst>
                            <a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs>
                            <a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr></a:gs>
                          </a:gsLst>
                          <a:lin ang="0"/>
                        </a:gradFill>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/ShadingType 2", pdf);
        TestAssert.Contains("/ca 0.5", pdf);
        TestAssert.Contains("72 396 144 72 re W n", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Supported linear gradient should not emit an unsupported-gradient diagnostic.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TRANSPARENCY"), "Uniform gradient alpha should not emit an unsupported-transparency diagnostic.");
    }

    public static void PptxSyntheticVariableAlphaGradientShapeEmitsDiagnosticUntilSoftMaskExists()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="Gradient"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:gradFill>
                          <a:gsLst>
                            <a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs>
                            <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
                          </a:gsLst>
                          <a:lin ang="0"/>
                        </a:gradFill>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/ShadingType 2", pdf);
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Variable gradient stop alpha needs PDF soft-mask support before it can be marked supported.");
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TRANSPARENCY"), "Variable gradient stop alpha should keep a transparency diagnostic until the PDF backend can express it exactly.");
    }

    public static void PptxSyntheticMultiStopLinearGradientShapeUsesStitchedPdfShading()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:gradFill>
                          <a:gsLst>
                            <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
                            <a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>
                            <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
                          </a:gsLst>
                          <a:lin ang="0"/>
                        </a:gradFill>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/FunctionType 3", pdf);
        TestAssert.Contains("/Bounds [0.5]", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Supported multi-stop linear gradient should not emit an unsupported-gradient diagnostic.");
    }

    public static void PptxSyntheticLineEndSizeVariantsAffectMarkerGeometry()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="diamond" w="sm" len="lg"/></a:ln></p:spPr></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("216 468 m", pdf);
        TestAssert.Contains("210 469.6 l", pdf);
        TestAssert.Contains("204 468 l", pdf);
        TestAssert.Contains("210 466.4 l", pdf);
    }

    public static void PptxSyntheticRotatedShapeProducesTransform()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm rot="2700000" flipH="1"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("q", pdf);
        TestAssert.Contains("-0.707107 0.707107 0.707107 0.707107", pdf);
        TestAssert.Contains(" cm", pdf);
        TestAssert.Contains("Q", pdf);
    }

    public static void PptxSyntheticCenteredShapeAutoFitCentersVisibleTextAndPreservesTrailingSpace()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2984500" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1600"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Forecasting models align compact labels with structural intent.</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFlowFrameSnapshot flowFrame = PptxRenderer.InspectTextFlow(document, package, 0).Frames.Single();
        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .Single()
            .Paragraphs
            .Single()
            .Lines
            .ToArray();

        TestAssert.True(lines.Length >= 2, $"Expected wrapped centered shape-autofit text, got {lines.Length} line.");
        double textRight = flowFrame.TextX + flowFrame.TextWidth;
        // PowerPoint keeps the fitting word and its trailing space on the first
        // line. The space may extend outside the frame; visible text stays centered.
        foreach (PptxTextLineLayoutSnapshot line in lines)
        {
            double visibleEndX = line.Spans.SelectMany(span => span.Atoms)
                .Where(atom => atom.Draw && atom.Kind is not ("Space" or "HiddenAdvance"))
                .Max(atom => atom.X + atom.Width);
            TestAssert.True(visibleEndX <= textRight + 0.01d,
                $"Centered shape-autofit visible text must fit the frame; visibleEnd={visibleEndX}, textRight={textRight}.");
            TestAssert.True(Math.Abs((line.StartX - flowFrame.TextX) - (textRight - visibleEndX)) < 0.01d,
                "Expected equal margins around the visible centered text, excluding trailing space.");
        }
        PptxTextAtomLayoutSnapshot trailingSpace = lines[0].Spans.SelectMany(span => span.Atoms)
            .Last(atom => atom.Kind == "Space");
        TestAssert.True(Math.Abs(lines[0].EndX - (trailingSpace.X + trailingSpace.Width)) < 0.01d,
            "Expected the wrapped line to preserve its trailing space for text emission.");
    }

    public static void PptxSyntheticVerticalShapeAutoFitPrefersSingleLine()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="365760" cy="2057400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="vert270" wrap="square" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t>Results reintegration for action</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.True(lines.Length == 1, "Vertical spAutoFit text should shrink before accepting an avoidable word wrap.");
        TestAssert.True(lines[0].EndX - lines[0].StartX <= 150d, "Fitted vertical text should stay inside the rotated text width.");
    }

    public static void PptxSyntheticRotatedGroupRotatesChildShape()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:grpSp>
                      <p:grpSpPr>
                        <a:xfrm rot="5400000">
                          <a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/>
                          <a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="1828800"/>
                        </a:xfrm>
                      </p:grpSpPr>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="685800" y="228600"/><a:ext cx="457200" cy="1371600"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 -1 1 0", pdf);
        TestAssert.Contains("36 108 re f", pdf);
    }

    public static void PptxSyntheticVerticalGradientKeepsFirstStopAtTop()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="2743200" y="2057400"/><a:ext cx="3657600" cy="2286000"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:gradFill flip="none" rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="2F80ED"/></a:gs><a:gs pos="100000"><a:srgbClr val="27AE60"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>
                        </p:spPr>
                      </p:sp>
                    </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match coords = Regex.Match(pdf, "/Coords \\[([^\\]]+)\\]");
        TestAssert.True(coords.Success, "Expected an axial shading coordinate array.");
        string[] parts = coords.Groups[1].Value.Split((char)32, StringSplitOptions.RemoveEmptyEntries);
        TestAssert.Equal(4, parts.Length);
        double y0 = double.Parse(parts[1], CultureInfo.InvariantCulture);
        double y1 = double.Parse(parts[3], CultureInfo.InvariantCulture);
        TestAssert.True(y0 > y1, "A 90-degree linear gradient must place the first stop at the top in y-up PDF space.");
        TestAssert.Contains("0.184 0.502 0.929", pdf);
    }

    public static void GroupTransformIdentityRoundTripsBounds()
    {
        // R17: the identity transform preserves bounds exactly (finite/range
        // contract baseline).
        var bounds = new PptxRenderer.ShapeBounds(10, 20, 300, 400, 0d, false, false);
        PptxRenderer.ShapeBounds applied = PptxRenderer.GroupTransform.Identity.Apply(bounds);
        TestAssert.Equal(bounds, applied);
    }

    public static void GroupTransformAppliesOffsetAndScale()
    {
        // R17: offsets and scales compose in the legacy order (local origin scaled
        // after subtracting the child offset, then shifted by the group origin).
        var transform = new PptxRenderer.GroupTransform(100, 200, 1000, 2000, 10, 20, 2d, 3d, 0d, false, false);
        var bounds = new PptxRenderer.ShapeBounds(20, 30, 30, 40, 0d, false, false);
        PptxRenderer.ShapeBounds applied = transform.Apply(bounds);
        TestAssert.Equal(120, applied.X);
        TestAssert.Equal(230, applied.Y);
        TestAssert.Equal(60, applied.Width);
        TestAssert.Equal(120, applied.Height);
    }

    public static void GroupTransformRejectsOverflowingCoordinates()
    {
        // R17: hostile EMU inputs fail as malformed geometry instead of overflowing
        // long math or wrapping offsets.
        var transform = new PptxRenderer.GroupTransform(long.MaxValue, 0, 1, 1, 0, 0, 1d, 1d, 0d, false, false);
        var bounds = new PptxRenderer.ShapeBounds(1, 0, 1, 1, 0d, false, false);
        TestAssert.Throws<InvalidDataException>(() => transform.Apply(bounds));
    }

    public static void GroupTransformRejectsNonFiniteInputs()
    {
        // R17: infinite scales and NaN rotations fail instead of carrying NaN into layout.
        var infinite = new PptxRenderer.GroupTransform(0, 0, 100, 100, 0, 0, double.PositiveInfinity, 1d, 0d, false, false);
        TestAssert.Throws<InvalidDataException>(() => infinite.Apply(new PptxRenderer.ShapeBounds(0, 0, 10, 10, 0d, false, false)));
        var nanRotation = new PptxRenderer.GroupTransform(0, 0, 100, 100, 0, 0, 1d, 1d, double.NaN, false, false);
        TestAssert.Throws<InvalidDataException>(() => nanRotation.Apply(new PptxRenderer.ShapeBounds(0, 0, 10, 10, 0d, false, false)));
    }

    public static void GroupTransformCombineRejectsExtremeChild()
    {
        // R17: nested composition inherits the same contract through Apply.
        var parent = new PptxRenderer.GroupTransform(0, 0, 100, 100, 0, 0, 1d, 1d, 0d, false, false);
        var child = new PptxRenderer.GroupTransform(long.MaxValue, long.MaxValue, long.MaxValue, long.MaxValue, 0, 0, 4d, 4d, 0d, false, false);
        TestAssert.Throws<InvalidDataException>(() => parent.Combine(child));
    }

    public static void RunTextAttributeReadersParseKnownAttributes()
    {
        // R14: the shared run-attribute readers interpret sz/spc/baseline/u/strike/cap
        // identically for scene building and rendering, with documented defaults.
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runProperties = new XElement(drawing + "rPr",
            new XAttribute("sz", "1200"),
            new XAttribute("spc", "50"),
            new XAttribute("baseline", "25000"),
            new XAttribute("u", "sng"),
            new XAttribute("strike", "sngStrike"),
            new XAttribute("cap", "all"));
        TestAssert.Equal(12d, PptxRunTextAttributeReaders.ReadFontSize(runProperties, null));
        TestAssert.Equal(0.5d, PptxRunTextAttributeReaders.ReadCharacterSpacing(runProperties, null));
        TestAssert.Equal(3d, PptxRunTextAttributeReaders.ReadBaselineOffset(runProperties, null, 12d));
        TestAssert.Equal("sng", PptxRunTextAttributeReaders.ReadUnderlineValue(runProperties, null));
        TestAssert.Equal("sngStrike", PptxRunTextAttributeReaders.ReadStrikeValue(runProperties, null));
        TestAssert.Equal("all", PptxRunTextAttributeReaders.ReadTextCapsValue(runProperties, null));
        TestAssert.True(PptxRunTextAttributeReaders.IsStrikeEnabled("sngStrike"), "Strike spellings other than noStrike enable strike.");
        TestAssert.True(!PptxRunTextAttributeReaders.IsStrikeEnabled("noStrike"), "noStrike disables strike.");
        TestAssert.True(!PptxRunTextAttributeReaders.IsStrikeEnabled(null), "Missing strike disables strike.");
        var defaultRunProperties = new XElement(drawing + "defRPr", new XAttribute("sz", "900"));
        TestAssert.Equal(18d, PptxRunTextAttributeReaders.ReadFontSize(null, null));
        TestAssert.Equal(9d, PptxRunTextAttributeReaders.ReadFontSize(null, defaultRunProperties));
        TestAssert.Equal(12d, PptxRunTextAttributeReaders.ReadFontSize(runProperties, defaultRunProperties));
        TestAssert.Equal(0d, PptxRunTextAttributeReaders.ReadCharacterSpacing(null, null));
        TestAssert.Equal(0d, PptxRunTextAttributeReaders.ReadBaselineOffset(null, null, 12d));
        TestAssert.True(PptxRunTextAttributeReaders.ReadUnderlineValue(null, null) is null, "Missing underline stays null.");
        TestAssert.True(PptxRunTextAttributeReaders.ReadStrikeValue(null, null) is null, "Missing strike stays null.");
        TestAssert.True(PptxRunTextAttributeReaders.ReadTextCapsValue(null, null) is null, "Missing caps stays null.");
    }
}
