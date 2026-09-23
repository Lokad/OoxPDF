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

internal static class PptxTypographyTests
{
    public static void PptxTextConversionUsesCustomFontResolver()
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
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:noFill/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>resolver probe</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var resolver = new CountingFontResolver();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        TestAssert.True(resolver.ResolveCalls > 0, "PPTX conversion should use the supplied font resolver for text layout and embedding.");
    }

    // Office drops out-of-frame clip lines whole (anchor-overflow reference:
    // 5 emitted operations, no Clip two), so a 24pt line that does not fit
    // the 12pt clip rectangle emits no text operation.
    public static void PptxSyntheticTextClipDropsLineWithBaselineOutsideClip()
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
                          <a:xfrm><a:off x="914400" y="0"/><a:ext cx="3657600" cy="152400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:noFill/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr vertOverflow="clip" lIns="0" tIns="0" rIns="0" bIns="0"/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>clipped cap</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = PptxTests.ReadPdfDecodedAscii(output);
        TestAssert.Contains(" W* n", pdf);
        TestAssert.Equal(0, PptxTests.CountTextMatrices(pdf));
    }

    public static void PptxSyntheticRotatedTextBoxProducesTransform()
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
                    <p:spPr><a:xfrm rot="1800000"><a:off x="2743200" y="1828800"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3200"><a:latin typeface="Arial"/></a:rPr><a:t>Rotated text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.866 -0.5 0.5 0.866", pdf);
        TestAssert.Contains("0052", pdf);
    }

    public static void PptxSyntheticTransparentTextUsesGlyphOutlinePaths()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="336699"><a:alpha val="45000"/></a:srgbClr></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>OO</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("GS45000F100000S", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f", pdf);
        TestAssert.True(!pdf.Contains(" TJ", StringComparison.Ordinal), "Transparent PPTX text should use glyph outline paths instead of positioned PDF text.");
        TestAssert.True(!pdf.Contains("> Tj", StringComparison.Ordinal), "Transparent PPTX text should use glyph outline paths instead of simple PDF text.");
    }

    public static void PptxSyntheticTransparentSyntheticBoldTextUsesGlyphOutlinePaths()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3600" b="1"><a:solidFill><a:srgbClr val="336699"><a:alpha val="45000"/></a:srgbClr></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><a:t>Bold</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("GS45000F100000S", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f", pdf);
        TestAssert.True(!pdf.Contains(" TJ", StringComparison.Ordinal), "Transparent synthetic-bold PPTX text should use glyph outline paths instead of positioned PDF text.");
        TestAssert.True(!pdf.Contains("> Tj", StringComparison.Ordinal), "Transparent synthetic-bold PPTX text should use glyph outline paths instead of simple PDF text.");
        TestAssert.True(!pdf.Contains(" Tr", StringComparison.Ordinal), "Transparent synthetic-bold PPTX text should not use PDF text rendering mode.");
    }

    public static void PptxSyntheticTransparentSyntheticItalicTextUsesGlyphOutlinePaths()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3600" i="1"><a:solidFill><a:srgbClr val="663399"><a:alpha val="45000"/></a:srgbClr></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><a:t>Italic</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("GS45000F100000S", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f", pdf);
        TestAssert.True(!pdf.Contains(" TJ", StringComparison.Ordinal), "Transparent synthetic-italic PPTX text should use sheared glyph outline paths instead of positioned PDF text.");
        TestAssert.True(!pdf.Contains("> Tj", StringComparison.Ordinal), "Transparent synthetic-italic PPTX text should use sheared glyph outline paths instead of simple PDF text.");
        TestAssert.True(!pdf.Contains(" Tr", StringComparison.Ordinal), "Transparent synthetic-italic PPTX text should not use PDF text rendering mode.");
    }

    public static void PptxSyntheticTransparentSyntheticBoldItalicTextUsesGlyphOutlinePaths()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3600" b="1" i="1"><a:solidFill><a:srgbClr val="663399"><a:alpha val="45000"/></a:srgbClr></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><a:t>Both</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("GS45000F100000S", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f", pdf);
        TestAssert.True(!pdf.Contains(" TJ", StringComparison.Ordinal), "Transparent synthetic-bold-italic PPTX text should use sheared glyph outline paths instead of positioned PDF text.");
        TestAssert.True(!pdf.Contains("> Tj", StringComparison.Ordinal), "Transparent synthetic-bold-italic PPTX text should use sheared glyph outline paths instead of simple PDF text.");
        TestAssert.True(!pdf.Contains(" Tr", StringComparison.Ordinal), "Transparent synthetic-bold-italic PPTX text should not use PDF text rendering mode.");
    }

    public static void PptxSyntheticFlippedRotatedTextBoxKeepsTextReadable()
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
                    <p:spPr><a:xfrm rot="10800000" flipV="1"><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Flipped text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0046", pdf);
        TestAssert.True(!Regex.IsMatch(pdf, @"-1\s+-?0\s+-?0\s+-1\s+[0-9.]+\s+[0-9.]+\s+cm"), "Flipped rotated text should not be rendered upside down.");
    }

    public static void PptxSyntheticBodyPrRotationOverridesShapeTextTransform()
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
                    <p:spPr><a:xfrm rot="10800000" flipV="1"><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr rot="0" lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Readable text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0052", pdf);
        TestAssert.True(!Regex.IsMatch(pdf, @"0\s+-1\s+1\s+0\s+[0-9.]+\s+[0-9.]+\s+cm"), "Explicit bodyPr rotation should override the shape text rotation.");
    }

    public static void PptxSyntheticTextOrientationVariantsProduceTransforms()
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
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="vert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>VERT</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="2286000" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="vert270" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>V270</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="eaVert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>EAV</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="5029200" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="wordArtVert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>WAV</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="6400800" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="wordArtVertRtl" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>WAR</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="7772400" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="mongolianVert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>MON</a:t></a:r></a:p></p:txBody></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 -1 1 0", pdf);
        TestAssert.Contains("-0 1 -1 -0", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TEXT_ORIENTATION"), "Known text orientation variants should be routed through the orientation model.");
    }

    public static void PptxSyntheticTextBoxEmbedsFontAndDrawsGlyphs()
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
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p>
                            <a:r>
                              <a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>
                              <a:t>Hello</a:t>
                            </a:r>
                          </a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Contains("/ToUnicode", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.Contains("/F1 24 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    // Office merges adjacent same-style source runs into one text operation;
    // the whitespace-controls visual probe pins 13 Office ops where the old
    // boundary rule emitted 15. Same-style runs merge; hyperlink, math/autofit,
    // alignment, and PreventCoalesce guards are unchanged.
    public static void PptxSyntheticTextBoxMergesSameStyleEmphasisSourceRuns()
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
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/>
                          <a:lstStyle/>
                          <a:p>
                            <a:r><a:rPr sz="1800" b="1"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha</a:t></a:r>
                            <a:r><a:rPr sz="1800" b="1"><a:latin typeface="Arial"/></a:rPr><a:t>Beta</a:t></a:r>
                            <a:r><a:rPr sz="1800" b="1"><a:latin typeface="Arial"/></a:rPr><a:t>Gamma</a:t></a:r>
                          </a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, PptxTests.CountOccurrences(pdf, " TJ"));
    }

    public static void PptxSyntheticTextBoxPreservesFractionalFontSize()
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
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p>
                            <a:r>
                              <a:rPr sz="996"><a:latin typeface="Arial"/></a:rPr>
                              <a:t>Fractional</a:t>
                            </a:r>
                          </a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 9.96 Tf", pdf);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal(9.96d, frame.Paragraphs[0].Runs[0].FontSize);

        PptxTextFrameLayoutSnapshot layoutFrame = PptxRenderer.InspectTextLayout(document, package, 0).Frames.Single();
        TestAssert.Equal(9.96d, layoutFrame.Paragraphs[0].Lines[0].Spans[0].FontSize);
    }

    public static void PptxSyntheticTextBoxClipIntersectsSlideBounds()
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
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="-45720"/><a:ext cx="914400" cy="274320"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr vertOverflow="clip" lIns="0" tIns="0" rIns="0" bIns="0"/>
                          <a:lstStyle/>
                          <a:p>
                            <a:r>
                              <a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr>
                              <a:t>Clip</a:t>
                            </a:r>
                          </a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 522 72 18 re W* n", pdf);
        TestAssert.DoesNotContain("72 522 72 21.6 re W* n", pdf);
    }

    public static void PptxSyntheticRoundRectUsesPresetTextRectangle()
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
                        <p:nvSpPr><p:cNvPr id="2" name="RectTextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>RECT</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                      <p:sp>
                        <p:nvSpPr><p:cNvPr id="3" name="RoundRectTextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                        <p:spPr>
                          <a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>
                          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>ROUND</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        IReadOnlyList<PptxTextFrameModelSnapshot> frames = PptxRenderer.InspectTextFrameModels(document, package, 0);

        PptxTextFrameModelSnapshot rect = frames.Single(frame => frame.Paragraphs[0].Runs[0].Text == "RECT");
        PptxTextFrameModelSnapshot round = frames.Single(frame => frame.Paragraphs[0].Runs[0].Text == "ROUND");
        double roundInset = (108d * 16667d / 100000d) * (1d - Math.Sqrt(0.5d));
        TestAssert.Equal(79.2d, rect.TextX);
        TestAssert.Equal(Math.Round(288d + 7.2d + roundInset, 6), Math.Round(round.TextX, 6));
        TestAssert.Equal(Math.Round(144d - 14.4d - (2d * roundInset), 6), Math.Round(round.TextWidth, 6));
    }

    public static void PptxSyntheticTextBoxAppliesOfficePdfFontSizeGridOnlyAtEmission()
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
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr><a:noAutofit/></a:bodyPr>
                          <a:lstStyle/>
                          <a:p>
                            <a:r>
                              <a:rPr sz="1000"><a:latin typeface="Arial"/></a:rPr>
                              <a:t>Office grid</a:t>
                            </a:r>
                          </a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 9.96 Tf", pdf);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal(10d, frame.Paragraphs[0].Runs[0].FontSize);

        PptxTextFrameLayoutSnapshot layoutFrame = PptxRenderer.InspectTextLayout(document, package, 0).Frames.Single();
        TestAssert.Equal(10d, layoutFrame.Paragraphs[0].Lines[0].Spans[0].FontSize);

        PptxTextGlyphRunSnapshot glyphRun = PptxRenderer.InspectTextGlyphRuns(document, package, 0).Single();
        TestAssert.Equal(0, glyphRun.FrameIndex);
        TestAssert.Equal(0, glyphRun.ParagraphIndex);
        TestAssert.Equal(0, glyphRun.LineIndex);
        TestAssert.Equal(frame.FontScale, glyphRun.FrameFontScale);
        TestAssert.Equal(72d, glyphRun.FrameShapeX);
        TestAssert.Equal(72d, glyphRun.FrameShapeTopY);
        TestAssert.Equal("Square", glyphRun.FrameWrapMode);
        TestAssert.Equal("noAutofit", glyphRun.FrameAutofitMode);
        TestAssert.Equal(frame.TextX, glyphRun.FrameTextX);
        TestAssert.Equal(frame.TextWidth, glyphRun.FrameTextWidth);
        TestAssert.Equal(1, glyphRun.FrameColumnCount);
        TestAssert.Equal(0d, glyphRun.FrameColumnSpacing);
        TestAssert.Equal(10d, glyphRun.LineMaxFontSize);
        TestAssert.Equal(10d, glyphRun.LayoutFontSize);
        TestAssert.Equal(9.96d, glyphRun.PdfFontSize);
    }

    public static void PptxSyntheticTextBoxAppliesOfficePdfFontSizeGridAcrossSizes()
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
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="5486400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"><a:noAutofit/></a:bodyPr>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr sz="700"><a:latin typeface="Arial"/></a:rPr><a:t>Size7</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="800"><a:latin typeface="Arial"/></a:rPr><a:t>Size8</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="900"><a:latin typeface="Arial"/></a:rPr><a:t>Size9</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="1000"><a:latin typeface="Arial"/></a:rPr><a:t>Size10</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="1300"><a:latin typeface="Arial"/></a:rPr><a:t>Size13</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="1400"><a:latin typeface="Arial"/></a:rPr><a:t>Size14</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="1600"><a:latin typeface="Arial"/></a:rPr><a:t>Size16</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="1900"><a:latin typeface="Arial"/></a:rPr><a:t>Size19</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="2000"><a:latin typeface="Arial"/></a:rPr><a:t>Size20</a:t></a:r></a:p>
                          <a:p><a:r><a:rPr sz="3000"><a:latin typeface="Arial"/></a:rPr><a:t>Size30</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        Dictionary<double, double> grid = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
            .ToDictionary(run => run.LayoutFontSize, run => run.PdfFontSize);

        TestAssert.Equal(6.96d, grid[7d]);
        TestAssert.Equal(8.04d, grid[8d]);
        TestAssert.Equal(9d, grid[9d]);
        TestAssert.Equal(9.96d, grid[10d]);
        TestAssert.Equal(12.96d, grid[13d]);
        TestAssert.Equal(14.04d, grid[14d]);
        TestAssert.Equal(15.96d, grid[16d]);
        TestAssert.Equal(18.96d, grid[19d]);
        TestAssert.Equal(20.04d, grid[20d]);
        TestAssert.Equal(30d, grid[30d]);
    }

    public static void PptxSyntheticTextBoxHonorsNoFillText()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="2400"><a:noFill/></a:rPr><a:t>Hidden</a:t></a:r>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t> visible</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/GS0F100000S gs", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTextBoxRendersTextOutline()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r>
                          <a:rPr sz="2400">
                            <a:noFill/>
                            <a:ln w="12700"><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln>
                            <a:latin typeface="Arial"/>
                          </a:rPr>
                          <a:t>Outline</a:t>
                        </a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.667 0 RG", pdf);
        TestAssert.Contains("1 Tr", pdf);
        TestAssert.Contains("0 Tr", pdf);
        TestAssert.Contains(" TJ", pdf);
        TestAssert.DoesNotContain(" c\r\nS", pdf);
        TestAssert.DoesNotContain(" h\r\nS", pdf);
    }

    public static void PptxSyntheticGradientTextFillEmitsDiagnosticUntilTextGradientExists()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400">
                        <a:gradFill><a:gsLst>
                          <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
                          <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
                        </a:gsLst><a:lin ang="0"/></a:gradFill>
                        <a:latin typeface="Arial"/>
                      </a:rPr><a:t>Gradient text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextRunModelSnapshot run = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .Single(textRun => textRun.Text == "Gradient text");
        TestAssert.Equal("FallbackBlack", run.ColorSource);
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Gradient text fill should remain diagnostic-covered until text gradients are rendered structurally.");
    }

    public static void PptxSyntheticTextBoxResolvesDrawingMlColorForms()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:prstClr val="orange"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Preset</a:t></a:r>
                        <a:br/>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:sysClr val="windowText" lastClr="112233"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>System</a:t></a:r>
                        <a:br/>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:scrgbClr r="50000" g="25000" b="0"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>ScRgb</a:t></a:r>
                        <a:br/>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:hslClr hue="7200000" sat="100000" lum="50000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Hsl</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0.647 0 rg", pdf);
        TestAssert.Contains("0.067 0.133 0.2 rg", pdf);
        TestAssert.Contains("0.502 0.251 0 rg", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsBodyInsets()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="914400" rIns="0" tIns="914400" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Inset</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("<< /Type /Pages /Count 1 /Kids [3 0 R] >>", pdf);
        TestAssert.Contains("/MediaBox [0 0 720 540]", pdf);
        TestAssert.True(pdf.Contains(" TJ", StringComparison.Ordinal) || pdf.Contains("> Tj", StringComparison.Ordinal), "Expected inset body text to be drawn as PDF text.");
    }

    public static void PptxTextModelExposesTypedBodyProperties()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="vert270" anchor="b" anchorCtr="1" wrap="none" vertOverflow="ellipsis" numCol="3" spcCol="914400" lIns="182880" rIns="0" tIns="457200" bIns="0" compatLnSpc="1" rot="5400000">
                        <a:normAutofit fontScale="80000" lnSpcReduction="12000"/>
                      </a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Body properties</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal(14.4d, frame.InsetLeft);
        TestAssert.Equal(0d, frame.InsetRight);
        TestAssert.Equal(36d, frame.InsetTop);
        TestAssert.Equal(0d, frame.InsetBottom);
        TestAssert.Equal("182880", frame.InsetLeftValue ?? string.Empty);
        TestAssert.Equal("0", frame.InsetRightValue ?? string.Empty);
        TestAssert.Equal("457200", frame.InsetTopValue ?? string.Empty);
        TestAssert.Equal("0", frame.InsetBottomValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.InsetLeftSource);
        TestAssert.Equal("DirectBodyPr", frame.InsetRightSource);
        TestAssert.Equal("DirectBodyPr", frame.InsetTopSource);
        TestAssert.Equal("DirectBodyPr", frame.InsetBottomSource);
        TestAssert.Equal("Vertical270", frame.Orientation);
        TestAssert.Equal("vert270", frame.OrientationValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.OrientationSource);
        TestAssert.Equal("Bottom", frame.VerticalAnchor);
        TestAssert.Equal("b", frame.VerticalAnchorValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.VerticalAnchorSource);
        TestAssert.Equal(true, frame.AnchorCenter);
        TestAssert.Equal("1", frame.AnchorCenterValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.AnchorCenterSource);
        TestAssert.Equal("None", frame.WrapMode);
        TestAssert.Equal("none", frame.WrapValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.WrapSource);
        TestAssert.Equal("Ellipsis", frame.VerticalOverflow);
        TestAssert.Equal("ellipsis", frame.VerticalOverflowValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.VerticalOverflowSource);
        TestAssert.Equal(3, frame.ColumnCount);
        TestAssert.Equal(72d, frame.ColumnSpacing);
        TestAssert.Equal("DirectBodyPr", frame.ColumnSource);
        TestAssert.Equal("DirectBodyPr", frame.ColumnCountSource);
        TestAssert.Equal("DirectBodyPr", frame.ColumnSpacingSource);
        TestAssert.Equal("3", frame.ColumnCountValue ?? string.Empty);
        TestAssert.Equal("914400", frame.ColumnSpacingValue ?? string.Empty);
        TestAssert.Equal("normAutofit", frame.AutofitModeValue);
        TestAssert.Equal("DirectBodyPr", frame.AutofitModeSource);
        TestAssert.Equal(0.8d, frame.FontScale);
        TestAssert.Equal("80000", frame.FontScaleValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.FontScaleSource);
        TestAssert.Equal(0.88d, frame.LineSpacingScale);
        TestAssert.Equal("12000", frame.LineSpacingReductionValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.LineSpacingScaleSource);
        TestAssert.Equal(true, frame.CompatibleLineSpacing);
        TestAssert.Equal("1", frame.CompatibleLineSpacingValue);
        TestAssert.Equal("DirectBodyPr", frame.CompatibleLineSpacingSource);
        TestAssert.Equal(90d, frame.RotationDegrees ?? double.NaN);
        TestAssert.Equal("5400000", frame.RotationValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.RotationDegreesSource);
    }

    public static void PptxTextModelKeepsUnknownBodyPropertyEnumsObservable()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="futureVert" anchor="futureAnchor" anchorCtr="futureAnchorCenter" wrap="futureWrap" vertOverflow="futureOverflow"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Unknown body properties</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal("Unknown", frame.Orientation);
        TestAssert.Equal("futureVert", frame.OrientationValue ?? string.Empty);
        TestAssert.Equal("Unknown", frame.VerticalAnchor);
        TestAssert.Equal("futureAnchor", frame.VerticalAnchorValue ?? string.Empty);
        TestAssert.Equal(false, frame.AnchorCenter);
        TestAssert.Equal("futureAnchorCenter", frame.AnchorCenterValue ?? string.Empty);
        TestAssert.Equal("Unknown", frame.WrapMode);
        TestAssert.Equal("futureWrap", frame.WrapValue ?? string.Empty);
        TestAssert.Equal("Unknown", frame.VerticalOverflow);
        TestAssert.Equal("futureOverflow", frame.VerticalOverflowValue ?? string.Empty);
    }

    public static void PptxTextModelPreservesInvalidNumericBodyPropertyTokens()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="futureColumns" spcCol="futureSpacing" lIns="futureLeft" rIns="futureRight" tIns="futureTop" bIns="futureBottom" rot="futureRotation">
                        <a:normAutofit fontScale="futureScale" lnSpcReduction="futureReduction"/>
                      </a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Invalid numeric body properties</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal(7.2d, frame.InsetLeft);
        TestAssert.Equal(7.2d, frame.InsetRight);
        TestAssert.Equal(3.6d, frame.InsetTop);
        TestAssert.Equal(3.6d, frame.InsetBottom);
        TestAssert.Equal("futureLeft", frame.InsetLeftValue ?? string.Empty);
        TestAssert.Equal("futureRight", frame.InsetRightValue ?? string.Empty);
        TestAssert.Equal("futureTop", frame.InsetTopValue ?? string.Empty);
        TestAssert.Equal("futureBottom", frame.InsetBottomValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.InsetLeftSource);
        TestAssert.Equal(1, frame.ColumnCount);
        TestAssert.Equal(0d, frame.ColumnSpacing);
        TestAssert.Equal("futureColumns", frame.ColumnCountValue ?? string.Empty);
        TestAssert.Equal("futureSpacing", frame.ColumnSpacingValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.ColumnCountSource);
        TestAssert.Equal("DirectBodyPr", frame.ColumnSpacingSource);
        TestAssert.Equal(1d, frame.FontScale);
        TestAssert.Equal("futureScale", frame.FontScaleValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.FontScaleSource);
        TestAssert.Equal(1d, frame.LineSpacingScale);
        TestAssert.Equal("futureReduction", frame.LineSpacingReductionValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.LineSpacingScaleSource);
        TestAssert.Equal(null, frame.RotationDegrees);
        TestAssert.Equal("futureRotation", frame.RotationValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.RotationDegreesSource);
    }

    public static void PptxTextModelDirectTopAnchorOverridesInheritedBottomAnchor()
    {
        string contentTypes = PptxTests.BasicContentTypes().Replace(
            "</Types>",
            """
              <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
              <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
            </Types>
            """,
            StringComparison.Ordinal);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sldMaster>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Layout Body"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr anchor="b"><a:normAutofit fontScale="80000"/></a:bodyPr><a:lstStyle/><a:p/></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sldLayout>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="3" name="Slide Body"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                    <p:txBody><a:bodyPr anchor="t"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Direct top anchor</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal("Top", frame.VerticalAnchor);
        TestAssert.Equal("t", frame.VerticalAnchorValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", frame.VerticalAnchorSource);
        TestAssert.Equal(0d, frame.VerticalOffset);
    }

    public static void PptxTextModelPreservesRunStyleTokens()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:pPr><a:defRPr u="dbl" strike="dblStrike" cap="small"/></a:pPr>
                        <a:r><a:rPr sz="1800" u="futureUnderline" strike="futureStrike" cap="futureCaps"/><a:t>Unknown</a:t></a:r>
                        <a:r><a:rPr sz="1800" u="none" strike="noStrike" cap="all"/><a:t>Off</a:t></a:r>
                        <a:r><a:rPr sz="1800"/><a:t>Default</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextRunModelSnapshot[] runs = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .Single()
            .Paragraphs.Single()
            .Runs.ToArray();
        TestAssert.Equal("futureUnderline", runs[0].UnderlineValue ?? string.Empty);
        TestAssert.True(runs[0].Underline, "Expected unknown underline tokens to remain enabled until an Office-aligned renderer handles them.");
        TestAssert.Equal("futureStrike", runs[0].StrikeValue ?? string.Empty);
        TestAssert.True(runs[0].Strike, "Expected unknown strike tokens to remain enabled until an Office-aligned renderer handles them.");
        TestAssert.Equal("futureCaps", runs[0].CapsValue ?? string.Empty);
        TestAssert.Equal("none", runs[1].UnderlineValue ?? string.Empty);
        TestAssert.True(!runs[1].Underline, "Expected explicit underline none to remain disabled.");
        TestAssert.Equal("noStrike", runs[1].StrikeValue ?? string.Empty);
        TestAssert.True(!runs[1].Strike, "Expected explicit noStrike to remain disabled.");
        TestAssert.Equal("all", runs[1].CapsValue ?? string.Empty);
        TestAssert.Equal("dbl", runs[2].UnderlineValue ?? string.Empty);
        TestAssert.Equal("dblStrike", runs[2].StrikeValue ?? string.Empty);
        TestAssert.Equal("small", runs[2].CapsValue ?? string.Empty);
    }

    public static void PptxSyntheticTextBoxHonorsLineBreaks()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 79.2 447.581 Tm", pdf);
        TestAssert.Contains("1 0 0 1 79.2 425.261 Tm", pdf);
    }

    public static void PptxSyntheticTextRunLineFeedForcesManualBreak()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Cambria"/></a:rPr><a:t>Alpha&#xA;Beta Gamma</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.True(frame.Paragraphs[0].HasManualLineBreak, "Expected literal line-feed text to preserve manual line-break state.");

        PptxTextFlowSegmentSnapshot[] segments = PptxRenderer.InspectTextFlow(document, package, 0)
            .Frames.Single()
            .Paragraphs.Single()
            .Runs.Single()
            .Segments.ToArray();
        TestAssert.Equal("Break", segments[1].Kind);

        string[] renderedLines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames.Single()
            .Paragraphs.Single()
            .Lines
            .Select(line => string.Concat(line.Spans.Select(span => span.Text)))
            .ToArray();
        TestAssert.Equal(2, renderedLines.Length);
        TestAssert.Equal("Alpha", renderedLines[0]);
        TestAssert.Equal("Beta Gamma", renderedLines[1]);
    }

    public static void PptxSyntheticTextBoxLineBreaksUseExplicitLineSpacing()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.True(frame.Paragraphs[0].HasVisibleContent, "Expected the text model to preserve paragraph-visible-content state.");
        TestAssert.True(frame.Paragraphs[0].HasManualLineBreak, "Expected the text model to preserve manual line-break state.");
        TestAssert.True(Math.Abs(frame.Paragraphs[0].FirstLineFallbackFontSize - 18d) < 0.01d, "Expected the text model to preserve the first-line fallback font size.");

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextLineLayoutSnapshot[] lines = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines.All(line => line.LineSpacingKind == "Absolute"), "Expected manual-break lines to keep absolute spcPts line spacing.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 24d) < 0.01d), "Expected absolute spcPts line spacing to own a 24pt line box advance.");
        TestAssert.True(Math.Abs((lines[0].BaselineY - lines[1].BaselineY) - 24d) < 0.01d, "Expected manual line break baselines to step by the absolute line spacing.");
    }

    public static void PptxSyntheticTextCompressedExplicitLineSpacingMovesBaselineByLineBoxReduction()
    {
        string input = Path.Combine(AppContext.BaseDirectory, "Cases", "pptx-ladder-04-typography-section-baseline-probe.pptx");
        if (!File.Exists(input))
        {
            input = Path.GetFullPath(Path.Combine("tests", "Lokad.OoxPdf.Tests", "Cases", "pptx-ladder-04-typography-section-baseline-probe.pptx"));
        }

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .Single()
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        const double officeBaselineFallback = 0.974d;
        const double normalLineHeight = 1.2d;
        const double lineSpacingCompression = 0.10d;
        double Expected(double fontSize) => fontSize * (officeBaselineFallback - normalLineHeight * lineSpacingCompression);

        TestAssert.True(lines.All(line => line.LineSpacingKind == "Multiple"), "Expected the section baseline probe to keep explicit percentage line-spacing provenance.");
        TestAssert.True(Math.Abs(lines[0].BaselineOffset - Expected(66d)) < 0.01d,
            $"Expected 90% explicit line spacing to move the 66pt baseline by the compressed normal line-box amount; got {lines[0].BaselineOffset.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(lines[2].BaselineOffset - Expected(36d)) < 0.01d,
            $"Expected 90% explicit line spacing to move the 36pt baseline by the compressed normal line-box amount; got {lines[2].BaselineOffset.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(lines[3].BaselineOffset - Expected(24d)) < 0.01d,
            $"Expected 90% explicit line spacing to move the 24pt baseline by the compressed normal line-box amount; got {lines[3].BaselineOffset.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticTrailingBreakUsesEndParagraphFontSize()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r><a:br/><a:endParaRPr sz="7200"/></a:p>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
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

        TestAssert.Equal(2, lines.Length);
        double baselineGap = lines[0].BaselineY - lines[1].BaselineY;
        TestAssert.True(baselineGap > 100d, $"Expected 72pt endParaRPr trailing line to push the next paragraph down, got {baselineGap.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxSyntheticEmptyParagraphUsesEndParagraphFontSize()
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
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr/><a:lstStyle/>
                        <a:p><a:endParaRPr sz="7200"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="4572000" y="914400"/><a:ext cx="2743200" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr/><a:lstStyle/>
                        <a:p><a:endParaRPr sz="1800"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextFrameModelSnapshot[] models = PptxRenderer.InspectTextFrameModels(document, package, 0).ToArray();
        TestAssert.Equal(2, layout.Frames.Count);
        TestAssert.Equal(2, models.Count(model => model.Paragraphs.Any(paragraph => paragraph.HasEndParagraphProperties)));

        PptxTextLineLayoutSnapshot largeEndParagraphLine = layout.Frames[0]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .Single();
        PptxTextLineLayoutSnapshot smallEndParagraphLine = layout.Frames[1]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .Single();

        double baselineGap = smallEndParagraphLine.BaselineY - largeEndParagraphLine.BaselineY;
        TestAssert.True(baselineGap > 50d, $"Expected empty 72pt endParaRPr paragraph to advance more than the 18pt control, got {baselineGap.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxSyntheticVerticalAnchorUsesEndParagraphFontSizeForEmptySpacing()
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
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p>
                        <a:p><a:pPr><a:spcBef><a:spcPct val="100000"/></a:spcBef></a:pPr><a:endParaRPr sz="7200"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="4572000" y="914400"/><a:ext cx="2743200" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p>
                        <a:p><a:pPr><a:spcBef><a:spcPts val="7200"/></a:spcBef></a:pPr><a:endParaRPr sz="7200"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot[] models = PptxRenderer.InspectTextFrameModels(document, package, 0).ToArray();
        TestAssert.Equal(2, models.Length);
        TestAssert.True(models.All(model => model.Paragraphs[1].HasEndParagraphProperties), "Expected text model inspection to expose the empty endParaRPr owners.");
        TestAssert.True(models.All(model => model.Paragraphs[1].HasLayoutContent), "Expected endParaRPr-only paragraphs to be marked as layout content by the text model.");
        TestAssert.True(models.All(model => Math.Abs(model.Paragraphs[1].EndParagraphFontSize - 72d) < 0.001d), "Expected resolved end-paragraph font size to be owned by the text model before layout.");
        TestAssert.True(models.All(model => Math.Abs(model.Paragraphs[1].EmptySpacingBefore - 72d) < 0.001d), "Expected empty-paragraph spacing before to be resolved by the text model before layout.");
        TestAssert.True(models.All(model => Math.Abs(model.Paragraphs[1].EmptySpacingAfter) < 0.001d), "Expected absent empty-paragraph spacing after to be resolved by the text model before layout.");

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        TestAssert.Equal(2, layout.Frames.Count);

        PptxTextLineLayoutSnapshot[] percentSpacingLines = layout.Frames[0]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();
        PptxTextLineLayoutSnapshot[] pointSpacingLines = layout.Frames[1]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, percentSpacingLines.Length);
        TestAssert.Equal(2, pointSpacingLines.Length);
        TestAssert.True(Math.Abs(percentSpacingLines[0].BaselineY - pointSpacingLines[0].BaselineY) < 0.01d, "Expected percent and equivalent point spacing before an empty endParaRPr paragraph to estimate the same anchored text height.");
        TestAssert.True(Math.Abs(percentSpacingLines[1].BaselineY - pointSpacingLines[1].BaselineY) < 0.01d, "Expected anchored follow-up paragraph baselines to stay equivalent after empty endParaRPr spacing.");
    }

    public static void PptxSyntheticVerticalAnchorUsesLineAdvanceForVisibleLines()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="4572000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>First</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        const double fontSize = 18d;
        double expectedOffset = (frame.TextHeight - 2d * fontSize * 1.2d) / 2d;

        TestAssert.Equal("Middle", frame.VerticalAnchor);
        TestAssert.True(Math.Abs(frame.VerticalOffset - expectedOffset) < 0.01d,
            $"Expected middle-anchor offset to use line advance (Office anchor-slack rule, anchor probe 2026-09-06); expected {expectedOffset.ToString("0.###", CultureInfo.InvariantCulture)}pt, got {frame.VerticalOffset.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxSyntheticMiddleAnchorUsesSignedActualSlackWhenWrappedTextOverflows()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="365760"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha beta gamma delta epsilon zeta eta theta</a:t></a:r></a:p>
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

        TestAssert.True(lines.Length > 1, "Expected narrow text to wrap before evaluating vertical anchoring.");
        TestAssert.True(Math.Abs(lines[0].TopY - 496.8d) < 0.01d,
            $"Expected overflowing middle-anchored text to use signed actual-layout slack instead of clamping to the frame top, got top Y {lines[0].TopY.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxSyntheticTextWrapKeepsBreakSpaceAtLineEnd()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="650000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta</a:t></a:r></a:p>
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

        TestAssert.Equal(2, lines.Length);
        string firstLine = string.Concat(lines[0].Spans.Select(span => span.Text));
        string secondLine = string.Concat(lines[1].Spans.Select(span => span.Text));
        TestAssert.Equal("Alpha ", firstLine);
        TestAssert.Equal("Beta", secondLine);
    }

    public static void PptxSyntheticLeadingManualBreakUsesBreakFontForAdvance()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>First line</a:t></a:r></a:p>
                      <a:p>
                        <a:pPr><a:defRPr sz="1800"><a:latin typeface="Cambria Math"/></a:defRPr></a:pPr>
                        <a:br><a:rPr sz="1100"><a:latin typeface="Cambria Math"/></a:rPr></a:br>
                        <a:r><a:rPr sz="1100"><a:latin typeface="Cambria Math"/></a:rPr><a:t>After break</a:t></a:r>
                      </a:p>
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

        TestAssert.Equal(2, lines.Length);
        const double normalLineHeight = 1.2d;
        const double manualBreakLineHeight = 1.24d;
        double expectedTopGap = 12d * normalLineHeight + 11d * manualBreakLineHeight;
        double actualTopGap = lines[0].TopY - lines[1].TopY;
        TestAssert.True(Math.Abs(actualTopGap - expectedTopGap) < 0.01d,
            $"Expected a leading manual break to advance by the break run font size; expected {expectedTopGap.ToString("0.###", CultureInfo.InvariantCulture)}pt, got {actualTopGap.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxSyntheticTextBeforeManualBreakKeepsNormalLineGrid()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Before break</a:t></a:r>
                        <a:br><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr></a:br>
                        <a:r><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>After break</a:t></a:r>
                      </a:p>
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

        TestAssert.Equal(2, lines.Length);
        double expectedAdvance = 12d * 1.2d;
        TestAssert.True(Math.Abs((lines[0].TopY - lines[1].TopY) - expectedAdvance) < 0.01d,
            $"Expected a non-leading manual break to keep the normal line grid; expected {expectedAdvance.ToString("0.###", CultureInfo.InvariantCulture)}pt, got {(lines[0].TopY - lines[1].TopY).ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.True(Math.Abs(lines[0].BaselineOffset - (12d * 0.9344d)) < 0.01d, "Expected the first line before a manual break to keep the manual-break baseline offset (Office fit-fixture: manual-break first baselines match the manual constant, not the floor).");
    }

    public static void PptxSyntheticTextWrapNoneKeepsLongTextOnOneLine()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="650000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta</a:t></a:r></a:p>
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

        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal("Alpha Beta", string.Concat(lines[0].Spans.Select(span => span.Text)));
    }

    public static void PptxSyntheticTextBoxUsesCompatibleLineSpacing()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr compatLnSpc="1"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t>Second</a:t></a:r></a:p>
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

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines.All(line => line.LineSpacingKind == "Multiple"), "Expected compatLnSpc to keep a multiple line-spacing model.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 14.4d) < 0.01d), $"Expected non-autofit compatLnSpc to use Office's normal default line advance, got {string.Join(",", lines.Select(line => line.Advance.ToString("0.###", CultureInfo.InvariantCulture)))}.");
        TestAssert.True(Math.Abs((lines[0].BaselineY - lines[1].BaselineY) - 14.4d) < 0.01d, $"Expected compatible line spacing to keep normal default baseline steps without shape autofit, got {(lines[0].BaselineY - lines[1].BaselineY).ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticTextBoxShapeAutoFitUsesTightCompatibleLineSpacing()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr compatLnSpc="1"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t>Second</a:t></a:r></a:p>
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

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines.All(line => line.LineSpacingKind == "Multiple"), "Expected shape-autofit compatLnSpc to keep a multiple line-spacing model.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 13.2d) < 0.01d), $"Expected shape-autofit compatLnSpc to use Office's tight default line advance, got {string.Join(",", lines.Select(line => line.Advance.ToString("0.###", CultureInfo.InvariantCulture)))}.");
        TestAssert.True(Math.Abs((lines[0].BaselineY - lines[1].BaselineY) - 13.2d) < 0.01d, $"Expected shape-autofit compatible line spacing to tighten default baseline steps, got {(lines[0].BaselineY - lines[1].BaselineY).ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticShapeAutoFitExplicitParagraphSpacingKeepsActualAnchorLineGrid()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2286000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr compatLnSpc="1" anchor="ctr" anchorCtr="0" lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr sz="1200"/><a:t>Control first</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t>Control second</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t>Control third</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                  <p:sp>
                    <p:spPr><a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="2286000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr compatLnSpc="1" anchor="ctr" anchorCtr="0" lIns="0" rIns="0" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr sz="1200"/><a:t>Autofit first</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t>Autofit second</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t>Autofit third</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameLayoutSnapshot[] frames = PptxRenderer.InspectTextLayout(document, package, 0).Frames.ToArray();
        PptxTextLineLayoutSnapshot[] controlLines = frames[0].Paragraphs.SelectMany(paragraph => paragraph.Lines).ToArray();
        PptxTextLineLayoutSnapshot[] autofitLines = frames[1].Paragraphs.SelectMany(paragraph => paragraph.Lines).ToArray();

        TestAssert.Equal(3, controlLines.Length);
        TestAssert.Equal(3, autofitLines.Length);
        TestAssert.True(autofitLines.All(line => line.LineSpacingKind == "Default"), "Expected explicit paragraph spacing to keep shape-autofit compatLnSpc on the normal default line-spacing model.");
        TestAssert.True(autofitLines.All(line => Math.Abs(line.Advance - 14.4d) < 0.01d), $"Expected explicit paragraph spacing to keep the normal 12pt default advance, got {string.Join(",", autofitLines.Select(line => line.Advance.ToString("0.###", CultureInfo.InvariantCulture)))}.");
        TestAssert.True(Math.Abs(autofitLines[0].BaselineY - controlLines[0].BaselineY) < 0.25d, $"Expected shape-autofit explicit paragraph spacing to use the same actual-line-box middle anchor as noAutofit, got {autofitLines[0].BaselineY.ToString("0.###", CultureInfo.InvariantCulture)} vs {controlLines[0].BaselineY.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticTextBoxCompatibleLineSpacingKeepsWrappedDefaultAdvance()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr compatLnSpc="1" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"/><a:t>Alpha Beta Gamma Delta</a:t></a:r></a:p>
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

        TestAssert.True(lines.Length >= 2, "Expected narrow text box to wrap without manual line breaks.");
        TestAssert.True(lines.All(line => line.LineSpacingKind == "Default"), "Expected compatLnSpc not to replace default spacing for ordinary wrapped lines.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 14.4d) < 0.01d), $"Expected wrapped lines to keep Office's normal 12pt default advance, got {string.Join(",", lines.Select(line => line.Advance.ToString("0.###", CultureInfo.InvariantCulture)))}.");
    }

    public static void PptxSyntheticTextBoxRendersFieldText()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum"><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>7</a:t></a:fld></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTextBoxIgnoresStandaloneTabElements()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>A</a:t></a:r><a:tab/><a:r><a:rPr sz="1800"/><a:t>B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 79.2d);
        PptxTests.AssertDoesNotContainTextMatrixAtX(pdf, 130.806d, "Standalone a:tab elements should not move following text.");
    }

    public static void PptxSyntheticTextBoxIgnoresStandaloneExplicitTabStops()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr><a:tabLst><a:tab pos="1828800"/></a:tabLst></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A</a:t></a:r><a:tab/><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Contains(" TJ", StringComparison.Ordinal) || pdf.Contains("> Tj", StringComparison.Ordinal), "Expected body text to be drawn as PDF text.");
        PptxTests.AssertDoesNotContainTextMatrixAtX(pdf, 216d, "Standalone a:tab elements should not move following text.");
    }
    public static void PptxCffFontSubstitutesFallbackWithDiagnostic()
    {
        byte[] cffBytes = TestFontBuilder.CreateCffKindFont("CffFamily");
        byte[] fallbackBytes = TestFontBuilder.CreateTestFont();
        string input = WriteCffProbePackage();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(
            input,
            output,
            new OoxPdfOptions { FontResolver = new CffFallbackPptxResolver(cffBytes, fallbackBytes), DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 1, "Substituted text must be painted.");
        TestAssert.DoesNotContain("CffFamily", pdf);
        TestAssert.Contains("TestFont", pdf);
        OoxPdfDiagnostic warning = TestAssert.NotNull(diagnostics.SingleOrDefault(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"));
        TestAssert.Equal(OoxPdfSeverity.Warning, warning.Severity);
        TestAssert.Contains("CffFamily", warning.Message);
        TestAssert.Equal("Per-glyph fallback typeface", warning.Fallback);
    }

    public static void PptxCffSubstitutionMatchesTrueTypeRendering()
    {
        byte[] cffBytes = TestFontBuilder.CreateCffKindFont("CffFamily");
        byte[] fallbackBytes = TestFontBuilder.CreateTestFont();
        string input = WriteCffProbePackage();
        string cffOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var cffDiagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(
            input,
            cffOutput,
            new OoxPdfOptions { FontResolver = new CffFallbackPptxResolver(cffBytes, fallbackBytes), DiagnosticSink = cffDiagnostics.Add });
        string ttOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var ttDiagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(
            input,
            ttOutput,
            new OoxPdfOptions { FontResolver = new CffFallbackPptxResolver(cffBytes, fallbackBytes, emitCff: false), DiagnosticSink = ttDiagnostics.Add });

        TestAssert.Equal(File.ReadAllText(ttOutput, Encoding.ASCII), File.ReadAllText(cffOutput, Encoding.ASCII));
        TestAssert.True(!ttDiagnostics.Any(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"), "TrueType rendering must not report a substitution.");
        TestAssert.True(cffDiagnostics.Any(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"), "CFF rendering must report the substitution.");
    }

    public static void PptxChartCffTextSubstitutesFallbackWithDiagnostic()
    {
        byte[] cffBytes = TestFontBuilder.CreateCffKindFont("CffFamily");
        byte[] fallbackBytes = TestFontBuilder.CreateTestFont();
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/charts/chart1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="CffFamily"/></a:rPr><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:bubbleChart>
                  <c:dLbls><c:showVal val="0"/><c:showBubbleSize val="1"/><c:dLblPos val="t"/><c:numFmt formatCode="0"/></c:dLbls>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
                    <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt></c:numLit></c:yVal>
                    <c:bubbleSize><c:numLit><c:formatCode>0</c:formatCode><c:pt idx="0"><c:v>9</c:v></c:pt><c:pt idx="1"><c:v>16</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:bubbleSize>
                  </c:ser>
                </c:bubbleChart></c:plotArea></c:chart></c:chartSpace>
                """
        });
        string cffOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var cffDiagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(
            input,
            cffOutput,
            new OoxPdfOptions { FontResolver = new CffFallbackPptxResolver(cffBytes, fallbackBytes, matchAllFamilies: true), DiagnosticSink = cffDiagnostics.Add });
        string ttOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var ttDiagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(
            input,
            ttOutput,
            new OoxPdfOptions { FontResolver = new CffFallbackPptxResolver(cffBytes, fallbackBytes, emitCff: false), DiagnosticSink = ttDiagnostics.Add });

        string cffPdf = File.ReadAllText(cffOutput, Encoding.ASCII);
        TestAssert.True(DocxTests.CountPdfTextShows(cffPdf) >= 1, "Substituted chart text must be painted.");
        TestAssert.DoesNotContain("CffFamily", cffPdf);
        TestAssert.True(cffDiagnostics.Any(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"), "Chart CFF rendering must report the substitution.");
        TestAssert.Equal(File.ReadAllText(ttOutput, Encoding.ASCII), cffPdf);
        TestAssert.True(!ttDiagnostics.Any(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"), "TrueType rendering must not report a substitution.");
    }

    private static string WriteCffProbePackage()
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
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
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:noFill/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="CffFamily"/></a:rPr><a:t>Hello</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
    }

    private sealed class CffFallbackPptxResolver(byte[] cffBytes, byte[] fallbackBytes, bool emitCff = true, string[]? cffFamilies = null, bool matchAllFamilies = false) : IFontResolver, IFontCatalog
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            bool cff = emitCff && (matchAllFamilies ? !request.FamilyName.Equals("FallbackFamily", StringComparison.OrdinalIgnoreCase) : (cffFamilies ?? ["CffFamily"]).Contains(request.FamilyName, StringComparer.OrdinalIgnoreCase));
            return new FontFaceResolution(
                request.FamilyName,
                cff ? "CffFamily" : "FallbackFamily",
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                new MemoryFontProgramSource(cff ? "cff-test" : "tt-test", cff ? cffBytes : fallbackBytes),
                IsFallback: false);
        }

        public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
        {
            return [Resolve(new FontRequest("FallbackFamily", false, false))];
        }
    }

}
