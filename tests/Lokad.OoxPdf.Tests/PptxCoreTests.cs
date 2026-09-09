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

internal static class PptxCoreTests
{
    public static void PptxSyntheticTwoSlidesProducesTwoPdfPages()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldSz cx="12192000" cy="6858000"/>
                  <p:sldIdLst>
                    <p:sldId id="256" r:id="rId1"/>
                    <p:sldId id="257" r:id="rId2"/>
                  </p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>",
            ["ppt/slides/slide2.xml"] = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>"
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>", pdf);
        TestAssert.Contains("/MediaBox [0 0 960 540]", pdf);
    }

    public static void PptxSyntheticLineEndPresetVariantsRender()
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
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="stealth"/></a:ln></p:spPr></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="1371600"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="diamond"/></a:ln></p:spPr></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="oval"/></a:ln></p:spPr></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("216 468 m", pdf);
        TestAssert.Contains("210 471 l", pdf);
        TestAssert.Contains("212 468 l", pdf);
        TestAssert.Contains("210 465 l", pdf);
        TestAssert.DoesNotContain("72 468 m 216 468 l S", pdf);
        TestAssert.Contains("212 435.2 l", pdf);
        TestAssert.Contains("208 432 l", pdf);
        TestAssert.Contains("212 428.8 l", pdf);
        TestAssert.Contains(" c", pdf);
    }

    public static void PptxSyntheticNoAutoFitOverflowColumnsKeepExplicitMultipleBaselineOnAscent()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        string repeated = string.Join(" ", Enumerable.Repeat("Structural layout keeps Office column baselines aligned", 34));
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="2194560"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" numCol="3" spcCol="182880" vertOverflow="overflow"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>{{repeated}}</a:t></a:r></a:p>
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
            .Single()
            .Paragraphs
            .Single()
            .Lines
            .ToArray();

        TestAssert.True(lines.Select(line => Math.Round(line.StartX, 2)).Distinct().Count() >= 2, "Expected explicit multiple line-spacing text to overflow into multiple columns.");
        TestAssert.True(lines.All(line => line.LineSpacingKind == "Multiple"), "Expected layout inspection to preserve explicit multiple line-spacing provenance.");
        TestAssert.True(lines.All(line => line.Advance > line.BaselineOffset + 1d),
            $"Expected 120% explicit line spacing to keep extra leading in the line advance; got advances {string.Join(",", lines.Select(line => line.Advance.ToString("0.###", CultureInfo.InvariantCulture)).Distinct(StringComparer.Ordinal))} and offsets {string.Join(",", lines.Select(line => line.BaselineOffset.ToString("0.###", CultureInfo.InvariantCulture)).Distinct(StringComparer.Ordinal))}.");
        TestAssert.True(lines.All(line => Math.Abs(line.BaselineOffset - 11.688d) < 0.01d),
            $"Expected noAutoFit overflow columns to place explicit percentage line-spacing baselines on Office's unscaled 12pt ascent while leaving leading in the line advance; got offsets {string.Join(",", lines.Select(line => line.BaselineOffset.ToString("0.###", CultureInfo.InvariantCulture)).Distinct(StringComparer.Ordinal))}.");
    }

    public static void PptxSyntheticVerticalAnchorIgnoresTerminalSpacingAfter()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p>
                        <a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcAft><a:spcPts val="1200"/></a:spcAft></a:pPr>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Centered</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        double expectedOffset = (frame.TextHeight - 18d * 1.2d * 0.9d) / 2d;

        TestAssert.Equal("Middle", frame.VerticalAnchor);
        TestAssert.True(Math.Abs(frame.VerticalOffset - expectedOffset) < 0.01d,
            $"Expected terminal paragraph spacing-after to be excluded from middle-anchor height; expected {expectedOffset.ToString("0.###", CultureInfo.InvariantCulture)}pt, got {frame.VerticalOffset.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxCambriaMathDenseWrapProbeKeepsHeadingOnOneLine()
    {
        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-04-cambria-math-dense-wrap-probe.pptx");
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();
        string[] texts = lines.Select(line => string.Concat(line.Spans.Select(span => span.Text))).ToArray();

        TestAssert.True(
            texts.Length > 0 && texts[0].EndsWith("France.", StringComparison.Ordinal),
            $"Expected Office-compatible first-line wrap. Lines: {string.Join(" | ", texts.Take(4))}. Widths: {string.Join(" | ", lines.Take(4).Select(line => (line.EndX - line.StartX).ToString("0.###", CultureInfo.InvariantCulture)))}");
    }

    public static void PptxBoundaryInvarianceProbeCoalescesLeftAlignedHighlightBoundaries()
    {
        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-04-typography-boundary-invariance-probe.pptx");
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);
        string[] glyphRunTexts = glyphRuns.Select(run => run.Text).ToArray();

        TestAssert.Equal(4, glyphRunTexts.Length);
        TestAssert.Equal("The scale and growth", glyphRunTexts[0]);
        TestAssert.Equal("The scale and growth", glyphRunTexts[1]);
        TestAssert.Equal("Large Global Supply", glyphRunTexts[2]);
        TestAssert.True(
            glyphRunTexts[3].StartsWith("D", StringComparison.Ordinal) &&
            glyphRunTexts[3].Contains("pendance", StringComparison.Ordinal) &&
            glyphRunTexts[3].EndsWith("e", StringComparison.Ordinal),
            $"Expected the accented left-aligned line to remain one glyph run, got '{glyphRunTexts[3]}'.");
    }

    public static void PptxSyntheticMongolianVerticalAutoFitDoesNotUseRotatedWidthShrink()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="365760"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="mongolianVert" wrap="square" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>VERTICAL STACKED TEXT</a:t></a:r></a:p>
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

        TestAssert.True(lines.Length > 1, "Stacked vertical text should use vertical chunk layout instead of the rotated single-line autofit path.");
        TestAssert.True(lines.All(line => Math.Abs(line.MaxFontSize - 24d) < 0.01d), "Stacked vertical spAutoFit should not shrink against the narrow rotated text width.");
    }

    public static void PptxSyntheticNoAutoFitWrapKeepsLineEndingSpace()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-bold-wrap-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        string[] renderedLines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Select(line => string.Concat(line.Spans.Select(span => span.Text)))
            .ToArray();

        TestAssert.Equal(2, renderedLines.Length);
        TestAssert.Equal("Quality decisions depend on careful operational planning and ", renderedLines[0]);
        TestAssert.Equal("reliable daily execution.", renderedLines[1]);
    }

    public static void PptxSyntheticOverflowColumnsPlaceBaselineBelowFrameBottom()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="495300"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="3" spcCol="144000" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Bravo</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Charlie</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Delta</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Echo</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Foxtrot</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Golf</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Hotel</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>India</a:t></a:r></a:p>
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

        TestAssert.True(
            lines.Count(line => Math.Abs(line.StartX - 72d) < 0.01d) == 3,
            "Expected overflow columns to place the last first-column baseline just below the frame bottom before advancing columns. Starts: " + string.Join(", ", lines.Select(line => line.StartX.ToString("0.###", CultureInfo.InvariantCulture))));
        TestAssert.True(lines.Count(line => Math.Abs(line.StartX - 219.78d) < 0.01d) == 3, "Expected balanced overflow text in the second column.");
        TestAssert.True(lines.Count(line => Math.Abs(line.StartX - 367.56d) < 0.01d) == 3, "Expected balanced overflow text in the third column.");
    }

    public static void PptxSyntheticMiddleAnchorUsesActualWrappedLineHeight()
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
                    <p:spPr><a:xfrm><a:off x="457200" y="914400"/><a:ext cx="12065000" cy="567531"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2000"><a:latin typeface="Arial"/></a:rPr><a:t>Operational planning aligns demand commitments, replenishment timing, pricing, staffing, and supplier constraints across the same weekly review cycle with finance governance checkpoints.</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot model = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(0d, model.VerticalOffset);
        TestAssert.Equal(2, lines.Length);
        TestAssert.True(
            lines[0].BaselineY < 452d,
            "Expected middle anchoring to use actual two-line layout height, not the pre-layout overestimate. Baselines: " +
            string.Join(", ", lines.Select(line => line.BaselineY.ToString("0.###", CultureInfo.InvariantCulture))));
    }

    public static void PptxStackedColumnBottomLegendPlotBoxUsesOfficeHorizontalPadding()
    {
        Type frameType = typeof(PptxRenderer).GetNestedType(
            "ChartFrameBox",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart frame box.");
        Type plotBoxType = typeof(PptxRenderer).GetNestedType(
            "ChartPlotBox",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart plot box.");
        Type legendLayoutType = typeof(PptxRenderer).GetNestedType(
            "ChartLegendLayout",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart legend layout.");
        Type shapeStyleType = typeof(PptxRenderer).GetNestedType(
            "ChartShapeStyle",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart shape style.");

        object frame = Activator.CreateInstance(frameType, [120d, 234d, 360d, 216d]) ?? throw new InvalidOperationException("Expected chart frame.");
        object plotBox = Activator.CreateInstance(plotBoxType, [140.24d, 280d, 333.32d, 159d]) ?? throw new InvalidOperationException("Expected plot box.");
        object emptyShapeStyle = shapeStyleType.GetProperty(
            "Empty",
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)?.GetValue(null) ?? throw new InvalidOperationException("Expected empty chart shape style.");
        object legend = Activator.CreateInstance(
            legendLayoutType,
            [
                PptxSceneChartLegendPosition.Bottom,
                "b",
                false,
                true,
                default(PptxSceneChartManualLayout),
                default(PptxSceneChartTextBodyProperties),
                emptyShapeStyle
            ]) ?? throw new InvalidOperationException("Expected chart legend layout.");
        System.Reflection.MethodInfo adjust = typeof(PptxRenderer).GetMethod(
            "AdjustStackedColumnBottomLegendPlotBox",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected stacked column bottom legend plot-box adjuster.");

        object adjusted = adjust.Invoke(null, [plotBox, frame, false, PptxSceneChartGrouping.Stacked, false, legend]) ?? throw new InvalidOperationException("Expected adjusted plot box.");
        double x = (double)(adjusted.GetType().GetProperty("X")?.GetValue(adjusted) ?? 0d);
        double width = (double)(adjusted.GetType().GetProperty("Width")?.GetValue(adjusted) ?? 0d);

        TestAssert.Equal(142.04d, Math.Round(x, 2));
        TestAssert.Equal(327.62d, Math.Round(width, 2));
    }

    public static void PptxUnsupportedFeaturesEmitDiagnostics()
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
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <a:graphic><a:graphicData uri="urn:vendor:unknown-graphic"/></a:graphic>
                    </p:graphicFrame>
                    <p:pic><p:nvPicPr><p:cNvPr id="2" name="Video"/><p:cNvPicPr/><p:nvPr><p:video/></p:nvPr></p:nvPicPr><p:blipFill><a:blip><a:videoFile/></a:blip></p:blipFill></p:pic>
                    <p:pic><p:nvPicPr><p:cNvPr id="3" name="Audio"/><p:cNvPicPr/><p:nvPr><p:audio/></p:nvPr></p:nvPicPr><p:blipFill><a:blip><a:audioFile/></a:blip></p:blipFill></p:pic>
                    <p:oleObj/>
                    <p:sp><p:spPr><a:gradFill/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:pattFill/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></p:spPr></p:sp>
                    <p:sp><p:spPr><a:effectLst><a:reflection/></a:effectLst></p:spPr></p:sp>
                    <p:sp><p:spPr><a:custGeom/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:prstGeom prst="wedgeRoundRectCallout"/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:prstGeom prst="heart"/><a:blipFill><a:blip/></a:blipFill></p:spPr></p:sp>
                    <p:pic><p:blipFill><a:blip><a:grayscl/></a:blip><a:tile/></p:blipFill></p:pic>
                    <p:sp><p:txBody><a:bodyPr vert="vert270" vertOverflow="ellipsis"/><a:lstStyle/><a:p/></p:txBody></p:sp>
                  </p:spTree></p:cSld>
                  <p:transition/>
                  <p:timing/>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string[] ids = diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal).ToArray();
        TestAssert.Equal(16, ids.Length);
        TestAssert.Contains("PPTX_UNSUPPORTED_ANIMATION", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_AUDIO", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_CALLOUT", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_CHART", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_CUSTOM_GEOMETRY", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_EFFECT", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_GRADIENT_FILL", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_GRAPHIC_FRAME", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_IMAGE_TILE", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_OLE_OBJECT", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_PATTERN_FILL", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_PICTURE_FILL", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_SMARTART", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_TEXT_OVERFLOW", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_TRANSITION", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_VIDEO", string.Join("|", ids));
        TestAssert.True(diagnostics.All(d => d.Severity == OoxPdfSeverity.Warning && d.SlideIndex == 1), "Unsupported PPTX diagnostics should be slide-scoped warnings.");
    }
}
