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

internal static class PptxTextLayoutTests
{
    public static void PptxSyntheticTextBoxHonorsVerticalAnchor()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Centered</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 7.2d);
    }

    public static void PptxSyntheticTextBoxVerticalAnchorUsesWrappedHeight()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta Gamma Delta</a:t></a:r></a:p>
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

        TestAssert.True(lines.Length >= 3, "Expected the narrow center-anchored text frame to wrap.");
        TestAssert.True(lines[0].BaselineY > 475d, "Expected vertical centering to account for wrapped line count instead of one logical paragraph line.");
    }

    public static void PptxSyntheticTextBoxVerticalAnchorUsesResolvedCharacterSpacing()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800" spc="2000"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta Gamma Delta</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        PptxTextLineLayoutSnapshot[] lines = layout
            .Frames
            .SelectMany(renderedFrame => renderedFrame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.True(lines.Length >= 4, "Expected resolved character spacing to increase wrapping before vertical centering.");
        TestAssert.True(frame.VerticalOffset < 40d, "Expected center anchoring to use the resolved character-spaced wrapped height.");
    }

    public static void PptxSyntheticTextBoxMiddleAnchorUsesVisibleRunFontSizes()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="4144300"/><a:ext cx="1634286" cy="699793"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr anchor="ctr"/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400"><a:latin typeface="Arial"/></a:rPr><a:t>       Data</a:t></a:r></a:p>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="900"><a:latin typeface="Arial"/></a:rPr><a:t>ERP/WMS/OMS/CRM</a:t></a:r></a:p>
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
        TestAssert.True(Math.Abs(lines[0].MaxFontSize - 14d) < 0.01d, "Expected the first line height to come from its visible 14pt run, not an unrelated 18pt fallback.");
        TestAssert.True(Math.Abs(lines[1].MaxFontSize - 9d) < 0.01d, "Expected the second line height to come from its visible 9pt run, not an unrelated 18pt fallback.");
        TestAssert.True(Math.Abs(lines[0].Advance - 16.8d) < 0.01d, "Expected default line advance to use the visible run font size.");
        TestAssert.True(Math.Abs(lines[1].Advance - 10.8d) < 0.01d, "Expected default line advance to use the visible run font size.");
    }

    public static void PptxSyntheticTransformedVerticalTextDoesNotDropPreclipBaselines()
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

        string emittedText = string.Concat(PptxRenderer.InspectTextGlyphRuns(document, package, 0).Select(run => run.Text));

        TestAssert.Contains("VERTIC", emittedText);
        TestAssert.Contains("TEXT", emittedText);
    }

    public static void PptxSyntheticHorizontalShapeAutoFitPreservesFontSizeWhenHeightOverflows()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-headline-wrap-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(4, lines.Length);
        TestAssert.True(lines.All(line => Math.Abs(line.MaxFontSize - 18d) < 0.01d), "Horizontal spAutoFit should preserve Office's run font size when only vertical text height overflows.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 21.6d) < 0.01d), "Horizontal spAutoFit should keep the normal 1.2 line advance instead of shrinking text.");

        string[] renderedLines = lines
            .Select(line => string.Concat(line.Spans.Select(span => span.Text)))
            .ToArray();

        TestAssert.Equal("Operational planning is decisions and execution. ", renderedLines[0]);
        TestAssert.Equal("Execute better, and you have a better operating ", renderedLines[1]);
        TestAssert.Equal("model. Make decisions better and you have a ", renderedLines[2]);
        TestAssert.Equal("better company.", renderedLines[3]);
    }

    public static void PptxTextOfficeTrailingEmphasisRunOwnsLineEndGlyphOperation()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-trailing-emphasis-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("Quality decisions depend on careful operational planning", string.Concat(lines[0].Spans.Select(span => span.Text)));
        TestAssert.Equal("and reliable daily execution.", string.Concat(lines[1].Spans.Select(span => span.Text)));
        TestAssert.Equal("planning", lines[0].Spans[^1].Text);
        double trailingSeparatorAdvance = lines[0].EndX - (lines[0].Spans[^1].X + lines[0].Spans[^1].Width);
        TestAssert.True(
            trailingSeparatorAdvance > 0d && trailingSeparatorAdvance < 4d,
            "Expected the emphasized trailing visible run to be followed only by the hidden line-ending separator advance.");

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        TestAssert.Equal(3, glyphRuns.Length);
        TestAssert.Equal("Quality decisions depend on careful operational ", glyphRuns[0].Text);
        TestAssert.Equal("planning", glyphRuns[1].Text);
        TestAssert.Equal("and reliable daily execution.", glyphRuns[2].Text);
        TestAssert.Equal(0, glyphRuns[1].LineIndex);
        TestAssert.Equal(1, glyphRuns[2].LineIndex);
        TestAssert.True(
            Math.Abs(glyphRuns[1].X - (glyphRuns[0].X + glyphRuns[0].Width)) < 0.01d,
            "Expected the emphasized glyph operation to start at the preceding run advance while remaining a separate text operation.");
    }

    public static void PptxTextManualBreakSameStyleLineCoalesces()
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
                    <p:spPr><a:xfrm><a:off x="2724150" y="3657600"/><a:ext cx="1498600" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr wrap="square" vertOverflow="overflow"><a:noAutofit/></a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:pPr><a:defRPr sz="900" i="1"><a:latin typeface="Cambria"/></a:defRPr></a:pPr><a:r><a:rPr sz="900" i="1"><a:latin typeface="Cambria"/></a:rPr><a:t>“A 10% better forecast&#xA;wouldn’t even register.”</a:t></a:r></a:p>
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
        TestAssert.Equal("“A 10% better forecast", string.Concat(lines[0].Spans.Select(span => span.Text)));
        TestAssert.Equal("wouldn’t even register.”", string.Concat(lines[1].Spans.Select(span => span.Text)));

        PptxTextGlyphRunSnapshot[] secondLineRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
            .Where(run => run.LineIndex == 1)
            .ToArray();

        TestAssert.Equal(1, secondLineRuns.Length);
        TestAssert.Equal("wouldn’t even register.”", secondLineRuns[0].Text);
    }

    public static void PptxSyntheticTextBoxHonorsNormAutofitFontScale()
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
                      <a:bodyPr><a:normAutofit fontScale="80000"/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3000"><a:latin typeface="Arial"/></a:rPr><a:t>Scaled</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 24 Tf", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TEXT_AUTOFIT"), "normAutofit fontScale should be handled.");
    }

    public static void PptxSyntheticTextBoxHonorsNormAutofitLineSpacingReduction()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:normAutofit lnSpcReduction="25000"/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="3000"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="3000"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        IReadOnlyList<PptxTextLineLayoutSnapshot> lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Count);
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 32.4d) < 0.01d), "Expected normAutofit line spacing reduction to scale explicit percentage line spacing over the normal line box.");
        TestAssert.True(Math.Abs((lines[0].BaselineY - lines[1].BaselineY) - 32.4d) < 0.01d, "Expected reduced line spacing to drive manual-break baseline steps.");
    }

    public static void PptxSyntheticTextBoxInheritsParagraphIndentFromListStyle()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/>
                      <a:lstStyle><a:lvl1pPr marL="91440" indent="45720"/></a:lstStyle>
                      <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1200"/><a:t>Indented</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot line = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single();
        PptxTextFrameModelSnapshot model = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();

        TestAssert.True(
            Math.Abs(line.Spans[0].X - 82.8d) < 0.01d,
            "Expected inherited marL plus first-line indent to move the emitted text span from the text body origin. Actual span: " + line.Spans[0].X.ToString("0.###", CultureInfo.InvariantCulture) +
            "; line start: " + line.StartX.ToString("0.###", CultureInfo.InvariantCulture) +
            "; cascade sources: " + model.Paragraphs[0].ResolvedCascadeSourceCount.ToString(CultureInfo.InvariantCulture) +
            "; margin: " + model.Paragraphs[0].MarginLeft.ToString("0.###", CultureInfo.InvariantCulture) +
            "; hanging: " + model.Paragraphs[0].HangingIndent.ToString("0.###", CultureInfo.InvariantCulture));
    }

    public static void PptxSyntheticEllipseTextUsesPresetTextRectangle()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:prstGeom prst="ellipse"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1800"/><a:t>7</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot line = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single();

        double expectedTextX = 72d + 144d * 0.1464466094067262d;
        TestAssert.True(
            Math.Abs(line.StartX - expectedTextX) < 0.01d,
            "Expected ellipse text to start inside the preset geometry text rectangle. Actual start: " + line.StartX.ToString("0.###", CultureInfo.InvariantCulture) +
            "; expected: " + expectedTextX.ToString("0.###", CultureInfo.InvariantCulture));

        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);
        TestAssert.Equal(1, glyphRuns.Count);
        TestAssert.True(
            Math.Abs(glyphRuns[0].X - expectedTextX) < 0.01d,
            "Expected clipped ellipse text to keep emitting a glyph run at the preset text-rectangle origin.");
    }

    public static void PptxSyntheticRectAndEllipseTextUseOfficeBaselineFloor()
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
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"/>
                        <a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Rect</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"/>
                        <a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Ellipse</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
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
            .OrderBy(line => line.StartX)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        PptxTextBaselineMetricSnapshot rectMetric = lines[0].BaselineMetric;
        PptxTextBaselineMetricSnapshot ellipseMetric = lines[1].BaselineMetric;
        TestAssert.True(
            Math.Abs(rectMetric.Ratio - 0.974d) < 0.001d,
            "Expected rectangular text frame to use the Office baseline floor. Actual ratio: " + rectMetric.Ratio.ToString("0.###", CultureInfo.InvariantCulture));
        TestAssert.True(
            Math.Abs(ellipseMetric.Ratio - 0.974d) < 0.001d,
            "Expected ellipse preset text frame to use the Office baseline floor (small-label-origin Office probe 2026-09-06). Actual ratio: " + ellipseMetric.Ratio.ToString("0.###", CultureInfo.InvariantCulture));
    }

    public static void PptxSyntheticRectTextKeepsCalibriMetricBelowOfficeBaselineFloor()
    {
        string calibri = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "calibri.ttf");
        if (!File.Exists(calibri))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(calibri))");
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Calibri"/></a:rPr><a:t>Calibri</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextBaselineMetricSnapshot metric = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single()
            .BaselineMetric;

        TestAssert.True(
            metric.Ratio > 0.94d && metric.Ratio < 0.974d,
            "Expected Calibri rectangular text to keep its resolved font metric instead of the Office fallback floor. Actual ratio: " + metric.Ratio.ToString("0.###", CultureInfo.InvariantCulture));
    }

    public static void PptxSyntheticRectTextUsesOfficeBaselineFloorForSmallDescenderFonts()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Cambria"/></a:rPr><a:t>Cambria</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextBaselineMetricSnapshot metric = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single()
            .BaselineMetric;

        TestAssert.True(
            Math.Abs(metric.Ratio - 0.974d) < 0.001d,
            "Expected rectangular text whose resolved font has a small Windows descender to use Office's fallback baseline floor. Actual ratio: " + metric.Ratio.ToString("0.###", CultureInfo.InvariantCulture));
    }

    public static void PptxSyntheticRectTextUsesOfficeBaselineFloorForSmallMetricFonts()
    {
        string cambriaMath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambriaMath))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambriaMath))");
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Math</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextBaselineMetricSnapshot metric = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single()
            .BaselineMetric;

        TestAssert.True(
            Math.Abs(metric.Ratio - 0.974d) < 0.001d,
            "Expected rectangular text whose resolved font metrics meet the baseline-floor thresholds to use Office's fallback baseline floor. Actual ratio: " + metric.Ratio.ToString("0.###", CultureInfo.InvariantCulture));
    }

    public static void PptxSyntheticTextBoxFlowsAcrossColumns()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="3" spcCol="144000" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>One two</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Three four</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Five six</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Seven eight</a:t></a:r></a:p>
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
        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();

        TestAssert.True(lines.Any(line => Math.Abs(line.StartX - 72d) < 0.01d), "Expected first-column text.");
        TestAssert.True(lines.Any(line => Math.Abs(line.StartX - 219.78d) < 0.01d), "Expected second-column text. Starts: " + string.Join(", ", lines.Select(line => line.StartX.ToString("0.###", CultureInfo.InvariantCulture))));
        TestAssert.True(glyphRuns.All(run => run.FrameColumnCount == 3), "Expected glyph-run inspection to preserve text-frame column count.");
        TestAssert.True(glyphRuns.All(run => Math.Abs(run.FrameColumnSpacing - 144000d / 12700d) < 0.01d), "Expected glyph-run inspection to preserve text-frame column spacing.");
    }

    public static void PptxSyntheticNoAutoFitTextOverflowColumnsBalanceOverloadedLastColumn()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        string text = string.Join(
            " ",
            Enumerable.Repeat(
                "Operational planning aligns demand commitments, replenishment timing, pricing, staffing, and supplier constraints across the same weekly review cycle with finance governance checkpoints.",
                10));
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="457200" y="914400"/><a:ext cx="6858000" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="3" spcCol="144000" lIns="91440" rIns="91440" tIns="45720" bIns="45720" vertOverflow="overflow"><a:noAutofit/></a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>{text}</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        int[] columnCounts = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .GroupBy(line => Math.Round(line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.Count())
            .ToArray();

        TestAssert.True(columnCounts.SequenceEqual([24, 24, 22]), "Expected Office-like overflow column balance. Counts: " + string.Join(", ", columnCounts));
    }

    public static void PptxSyntheticNoAutoFitTextOverflowColumnsBalanceEvenContinuedParagraph()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        var paragraph = new StringBuilder();
        for (int i = 1; i <= 66; i++)
        {
            paragraph.Append(CultureInfo.InvariantCulture, $"""<a:r><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Line {i:00}</a:t></a:r>""");
            if (i != 66)
            {
                paragraph.Append("<a:br/>");
            }
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="457200" y="914400"/><a:ext cx="6858000" cy="4267200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="3" spcCol="144000" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"><a:noAutofit/></a:bodyPr>
                      <a:lstStyle/>
                      <a:p>{paragraph}</a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        int[] columnCounts = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraphLayout => paragraphLayout.Lines)
            .GroupBy(line => Math.Round(line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.Count())
            .ToArray();

        TestAssert.True(columnCounts.SequenceEqual([22, 23, 21]), "Expected the continued paragraph to move one line from the final column into the penultimate column. Counts: " + string.Join(", ", columnCounts));
    }

    public static void PptxSyntheticTextBoxOverflowColumnsReserveNormalLineAdvance()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="350520"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="2" spcCol="0" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"><a:noAutofit/></a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Bravo</a:t></a:r></a:p>
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
        TestAssert.True(Math.Abs(lines[0].StartX - 72d) < 0.01d, "Expected the first line in the first overflow column.");
        TestAssert.True(
            Math.Abs(lines[1].StartX - 216d) < 0.01d,
            "Expected overflow column fit to reserve default normal line advance, not only font size. Starts: " +
            string.Join(", ", lines.Select(line => line.StartX.ToString("0.###", CultureInfo.InvariantCulture))));
    }

    public static void PptxSyntheticTextBoxOverflowColumnsUseSlideClip()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="685800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="3" spcCol="144000" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>One two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 720 540 re W* n", pdf);
        TestAssert.True(
            !pdf.Contains("72 414 135.78 54 re W* n", StringComparison.Ordinal),
            "Expected multi-column text with overflow enabled to keep Office's slide-wide text clip instead of clipping to the first column.");
        TestAssert.True(
            !pdf.Contains("219.78 414 135.78 54 re W* n", StringComparison.Ordinal),
            "Expected multi-column text with overflow enabled to keep Office's slide-wide text clip instead of clipping to the second column.");
    }

    public static void PptxSyntheticTextManualBreaksStayInActiveOverflowColumn()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="2" spcCol="0" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>First column line</a:t></a:r><a:br/><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Second first column</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Second column first</a:t></a:r><a:br/><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Second column second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot[] secondParagraphLines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .Select((paragraph, index) => new { paragraph, index })
            .Where(item => item.index == 1)
            .Select(item => item.paragraph)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, secondParagraphLines.Length);
        TestAssert.True(
            secondParagraphLines.All(line => Math.Abs(line.StartX - 216d) < 0.01d),
            "Expected manual-break lines to align against the active second column, not the full text frame. Starts: " +
            string.Join(", ", secondParagraphLines.Select(line => line.StartX.ToString("0.###", CultureInfo.InvariantCulture))));
    }

    public static void PptxSyntheticTextBoxSplitsOverwideFirstSegmentAcrossColumns()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="685800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="3" spcCol="144000" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</a:t></a:r></a:p>
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

        TestAssert.True(lines.Any(line => Math.Abs(line.StartX - 72d) < 0.01d), "Expected overwide first segment in first column.");
        TestAssert.True(lines.Any(line => Math.Abs(line.StartX - 219.78d) < 0.01d), "Expected overwide first segment to flow into second column. Starts: " + string.Join(", ", lines.Select(line => line.StartX.ToString("0.###", CultureInfo.InvariantCulture))));
        TestAssert.True(lines.Any(line => Math.Abs(line.StartX - 367.56d) < 0.01d), "Expected overwide first segment to flow into third column. Starts: " + string.Join(", ", lines.Select(line => line.StartX.ToString("0.###", CultureInfo.InvariantCulture))));
        TestAssert.True(lines.All(line => line.NaturalEndX <= line.EndX + 0.01d), "Expected split lines to stay inside their column widths.");
    }

    public static void PptxSyntheticNoAutoFitTextOverwideFirstSegmentUsesOfficeWrapFitTolerance()
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
                    <p:spPr><a:xfrm><a:off x="817304" y="914400"/><a:ext cx="10626315" cy="650875"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr anchor="ctr" lIns="91440" rIns="91440" tIns="45720" bIns="45720"><a:noAutofit/></a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Cambria Math"/></a:rPr><a:t>XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX</a:t></a:r></a:p>
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
        int[] lineTextLengths = lines
            .Select(line => line.Spans.Sum(span => span.Text.Length))
            .ToArray();
        TestAssert.True(lineTextLengths.All(length => length == 80),
            "Expected Office-fit tolerance to keep the dense overwide no-autofit run at two 80-character lines. Lengths: " +
            string.Join(", ", lineTextLengths.Select(length => length.ToString(CultureInfo.InvariantCulture))));
    }

    // Office drops the 48pt line whose baseline sits outside the 36pt clip
    // rectangle: no text operation is emitted for out-of-frame clip lines.
    public static void PptxSyntheticTextBoxClipDropsLineWithBaselineOutsideClip()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="clip"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="4800"/><a:t>UnbreakableOverflowingText</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(0, PptxTests.CountTextMatrices(pdf));
    }

    private static string FindVisualCase(string name)
    {
        string[] candidates = new[]
        {
            Path.Combine("tests", "Lokad.OoxPdf.Tests", "Cases", name),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name),
        };
        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException("Case file not found: " + name);
    }

    public static void PptxEllipseMicroLabelKeepsOfficePositionedLine()
    {
        // F01: an ellipse micro-box under vertOverflow clip keeps the line whose
        // baseline sits inside shape bounds (Office Y=359.11) even though preset
        // inscription shrinks the flow rect past it. The vertical clip spans shape
        // bounds; horizontal placement keeps the inscribed rect (Office X=186.38).
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = FindVisualCase("pptx-ladder-04-typography-small-label-origin-probe.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, PptxTests.CountTextMatrices(pdf));
        TestAssert.True(pdf.Contains("/FontFile2", StringComparison.Ordinal), "Kept line must embed its font.");
    }

    public static void PptxSyntheticTextBoxEllipsisAddsMarkerAtLastVisibleLine()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="ellipsis"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Visible line</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Clipped line</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            string[] glyphRunTexts = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
                .Select(run => run.Text)
                .ToArray();

            TestAssert.True(
                glyphRunTexts.Contains("Visible line", StringComparer.Ordinal),
                "Expected vertOverflow=\"ellipsis\" to keep text whose baseline remains inside the text rectangle.");
            TestAssert.True(
                !glyphRunTexts.Contains("Clipped line", StringComparer.Ordinal),
                "Expected vertOverflow=\"ellipsis\" to drop text whose glyph outline does not intersect the text rectangle.");
            TestAssert.True(
                glyphRunTexts.Contains("…", StringComparer.Ordinal),
                "Expected vertOverflow=\"ellipsis\" to render an Office-style ellipsis marker after the last visible line.");
        }

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 450 216 18 re W* n", pdf);
        TestAssert.True(
            diagnostics.All(diagnostic => diagnostic.Id != "PPTX_UNSUPPORTED_TEXT_OVERFLOW"),
            "Expected shape text ellipsis overflow to be handled by the shared text-frame renderer.");
    }

    // Office drops out-of-frame lines under vertOverflow clip: the live
    // anchor-overflow reference emits 5 text operations with no Clip two,
    // whose baseline sits outside the text rectangle even though its glyph
    // outline intersects it.
    // Office drops ellipsis lines whose baseline sits outside the text rectangle
    // and appends the marker inline after the last kept line: the ellipsis reference
    // emits Visible plus U+2026 on one baseline with no second line.
    public static void PptxSyntheticEllipsisDropsBaselineOutsideLineAndMarksInline()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="ellipsis"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Visible line</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Hidden line</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            string[] glyphRunTexts = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
                .Select(run => run.Text)
                .ToArray();

            TestAssert.Equal(2, glyphRunTexts.Length);
            TestAssert.True(
                glyphRunTexts.Contains("Visible line", StringComparer.Ordinal),
                "Expected vertOverflow=\"ellipsis\" to keep text whose baseline remains inside the text rectangle.");
            TestAssert.True(
                !glyphRunTexts.Contains("Hidden line", StringComparer.Ordinal),
                "Expected vertOverflow=\"ellipsis\" to drop text whose baseline sits outside the text rectangle instead of emitting a clipped sliver.");
            TestAssert.True(
                glyphRunTexts.Contains("…", StringComparer.Ordinal),
                "Expected vertOverflow=\"ellipsis\" to render an Office-style ellipsis marker after the last visible line.");
        }

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 432 216 36 re W* n", pdf);
        TestAssert.True(
            diagnostics.All(diagnostic => diagnostic.Id != "PPTX_UNSUPPORTED_TEXT_OVERFLOW"),
            "Expected shape text ellipsis overflow to be handled by the shared text-frame renderer.");    }

    public static void PptxTextFrameVerticalClipDropsLinesWithBaselineOutsideClip()
    {
        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-03-text-anchor-overflow.pptx");
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        string[] glyphRunTexts = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
            .Select(run => run.Text)
            .ToArray();

        TestAssert.True(
            glyphRunTexts.Contains("Clip one", StringComparer.Ordinal),
            "Expected first clipped line to remain visible.");
        TestAssert.True(
            !glyphRunTexts.Contains("Clip two", StringComparer.Ordinal),
            "Expected vertOverflow=\"clip\" to drop the line whose baseline sits outside the text rectangle, matching the 5-operation Office reference.");
        TestAssert.True(
            glyphRunTexts.Contains("Flow one", StringComparer.Ordinal),
            "Expected overflow-enabled companion frame to keep its first line.");
        TestAssert.True(
            glyphRunTexts.Contains("Flow two", StringComparer.Ordinal),
            "Expected overflow-enabled companion frame to keep its second line.");
    }

    public static void PptxSyntheticTextBoxAllowsVerticalOverflowByDefault()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="4800"/><a:t>Overflow</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 720 540 re W* n", pdf);
    }

    public static void PptxSyntheticTextBoxOverflowKeepsOffSlidePdfStructure()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="-2286000"/><a:ext cx="3657600" cy="6400800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Overflow structure above slide</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);

        TestAssert.True(
            glyphRuns.Any(run => run.BaselineY > document.SlideHeightPoints),
            "Expected overflow text to keep PDF text structure outside the slide clip instead of being pre-culled by the renderer.");
    }

    public static void PptxSyntheticNoAutoFitTextBoxUsesSlideWidthTextClip()
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
                    <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"/><a:t>No autofit text wraps with Office slide clip.</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 720 540 re W* n", pdf);
    }

    public static void PptxSyntheticShapeAutoFitTextBoxUsesSlideWidthTextClip()
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
                    <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"/><a:t>Shape autofit text wraps with Office slide clip.</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 720 540 re W* n", pdf);
    }

    public static void PptxSyntheticPlainShapeTextUsesSlideWidthTextClip()
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
                    <p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2550000" cy="276000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Plain shape text uses Office slide clip.</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 720 540 re W* n", pdf);
    }

    public static void PptxAlignmentValuesKeepDistinctOfficeTextDistributionModes()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-alignment-values-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFlowSnapshot flow = PptxRenderer.InspectTextFlow(document, package, 0);
        string[] alignments = flow.Frames
            .SelectMany(frame => frame.Paragraphs)
            .Select(paragraph => paragraph.Alignment)
            .ToArray();

        string alignmentList = string.Join("|", alignments);
        TestAssert.Contains("Justify", alignmentList);
        TestAssert.Contains("Distributed", alignmentList);
        TestAssert.Contains("JustLow", alignmentList);
        TestAssert.Contains("ThaiDistributed", alignmentList);

        PptxTextFrameModelSnapshot[] models = PptxRenderer.InspectTextFrameModels(document, package, 0).ToArray();
        string alignmentValueList = string.Join("|", models
            .SelectMany(frame => frame.Paragraphs)
            .Select(paragraph => paragraph.AlignmentValue ?? string.Empty));
        TestAssert.Contains("just", alignmentValueList);
        TestAssert.Contains("dist", alignmentValueList);
        TestAssert.Contains("justLow", alignmentValueList);
        TestAssert.Contains("thaiDist", alignmentValueList);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextLineLayoutSnapshot distributed = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .First(line => line.Alignment == "Distributed");

        TestAssert.True(distributed.Spans.Count > 10, "Expected distributed alignment to own per-glyph positioned spans.");
        TestAssert.True(distributed.Spans.All(span => span.Text.Length <= 2), "Expected distributed alignment to avoid word-level spans that hide letter spacing.");
        TestAssert.True(distributed.Spans.Last().X - distributed.StartX > distributed.Advance * 0.75d, "Expected distributed alignment to stretch glyph positions across the text frame.");
    }

    public static void PptxHighlightedTextRunDoesNotApplyImplicitTracking()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-tracking-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFlowSnapshot flow = PptxRenderer.InspectTextFlow(document, package, 0);
        PptxTextFlowRunSnapshot[] runs = flow.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.SourceText.Length != 0)
            .ToArray();

        PptxTextFlowRunSnapshot highlighted = runs.First(run => run.SourceText == "AI");
        PptxTextFlowRunSnapshot following = runs.First(run => run.SourceText.StartsWith(" boundary", StringComparison.Ordinal));

        TestAssert.True(Math.Abs(highlighted.FontSize - 12d) < 0.01d, "Expected the probe highlight run to stay at 12pt.");
        TestAssert.True(highlighted.Segments.All(segment => Math.Abs(segment.FontScale - 1d) < 0.01d), "Expected highlight tracking to avoid fake font scaling.");

        TestAssert.True(following.FontSize > 0d, "Expected the following run to remain present in text flow.");

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        PptxTextGlyphRunSnapshot[] trackedRuns = glyphRuns
            .Where(run => Math.Abs(run.PdfCharacterSpacing + 0.036d) < 0.001d)
            .ToArray();

        TestAssert.Equal(2, trackedRuns.Length);
        TestAssert.True(
            trackedRuns.Any(run => run.HighlightColor is not null),
            "Expected Office's PDF character-spacing state to start at the highlighted run.");
        TestAssert.True(
            trackedRuns.Any(run => run.HighlightColor is null),
            "Expected the same-paragraph continuation after the highlight to keep Office's PDF character-spacing state.");
        TestAssert.True(
            trackedRuns.All(run => Math.Abs(run.LayoutCharacterSpacing) < 0.001d),
            "Expected highlight continuation character spacing to be an emission-only PDF text state.");
    }

    public static void PptxSingleParagraphHighlightDoesNotStartContinuationTextState()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-highlight-single.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();

        TestAssert.True(
            glyphRuns.Any(run => run.HighlightColor is not null),
            "Expected the probe to contain a highlighted glyph run.");
        TestAssert.True(
            glyphRuns.All(run => Math.Abs(run.PdfCharacterSpacing) < 0.001d),
            "Expected a first-paragraph highlight to leave Office's PDF character-spacing state at zero.");
    }

    public static void PptxNoAutofitHeadlineHighlightDoesNotStartContinuationTextState()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-highlighted-headline-runs.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();

        TestAssert.True(
            glyphRuns.Any(run => run.FrameAutofitMode == "noAutofit" && run.HighlightColor is not null),
            "Expected the probe to contain highlighted noAutofit headline runs.");
        TestAssert.True(
            glyphRuns.All(run => Math.Abs(run.PdfCharacterSpacing) < 0.001d),
            "Expected highlighted noAutofit headline text to keep Office's PDF character-spacing state at zero.");
    }

    public static void PptxTextGlyphRunInspectionReportsInterGlyphAdjustmentAggregates()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-tracking-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot glyphRun = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
            .First(run => run.Glyphs.Count > 2);
        double[] adjustments = glyphRun.Glyphs.Skip(1).Select(glyph => glyph.AdjustmentBefore).ToArray();

        TestAssert.Equal(adjustments.Length, glyphRun.InterGlyphAdjustmentCount);
        TestAssert.True(Math.Abs(adjustments.Sum() - glyphRun.InterGlyphAdjustmentSum) < 0.0001d, "Expected glyph-run aggregate sum to match exposed atom adjustments.");
        TestAssert.True(Math.Abs(adjustments.Min() - glyphRun.InterGlyphAdjustmentMin) < 0.0001d, "Expected glyph-run aggregate minimum to match exposed atom adjustments.");
        TestAssert.True(Math.Abs(adjustments.Max() - glyphRun.InterGlyphAdjustmentMax) < 0.0001d, "Expected glyph-run aggregate maximum to match exposed atom adjustments.");
        TestAssert.True(Math.Abs(adjustments.Average() - glyphRun.InterGlyphAdjustmentAverage) < 0.0001d, "Expected glyph-run aggregate average to match exposed atom adjustments.");
    }

    public static void PptxHighlightedContinuationEmitsOfficeCharacterSpacingTextStateAcrossAutofitWrap()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-tracking-narrow-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        PptxTextGlyphRunSnapshot[] trackedRuns = glyphRuns
            .Where(run => Math.Abs(run.PdfCharacterSpacing + 0.036d) < 0.001d)
            .ToArray();

        TestAssert.Equal(3, trackedRuns.Length);
        TestAssert.True(
            trackedRuns.Any(run => run.HighlightColor is not null),
            "Expected the highlighted autofit run to receive Office's PDF character-spacing state.");
        TestAssert.True(
            trackedRuns.Any(run => run.HighlightColor is null),
            "Expected same-paragraph autofit continuation text to keep the same PDF character-spacing state.");
        TestAssert.True(
            trackedRuns.All(run => Math.Abs(run.LayoutCharacterSpacing) < 0.001d),
            "Expected the Office character-spacing state to be an emission-only decomposition, not layout tracking.");
    }

    public static void PptxHighlightedTextStateContinuesAcrossFollowingParagraph()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-slide3-narrow-cambria-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        PptxTextGlyphRunSnapshot[] trackedRuns = glyphRuns
            .Where(run => Math.Abs(run.PdfCharacterSpacing + 0.036d) < 0.001d)
            .ToArray();

        TestAssert.Equal(7, trackedRuns.Length);
        TestAssert.True(
            trackedRuns.Any(run => run.HighlightColor is not null),
            "Expected the highlighted run to start Office's PDF character-spacing state.");
        TestAssert.True(
            trackedRuns.Any(run => run.ParagraphIndex == 2),
            "Expected the following paragraph in the same frame to keep Office's PDF character-spacing state.");
        TestAssert.True(
            trackedRuns.All(run => Math.Abs(run.LayoutCharacterSpacing) < 0.001d),
            "Expected paragraph-spanning highlight continuation spacing to be an emission-only PDF text state.");
    }

    public static void PptxNumberedAutofitFramesEmitOfficeCharacterSpacingTextState()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-numbered-tc-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        AssertNumberedFrameCharacterSpacing(glyphRuns);
        TestAssert.True(
            glyphRuns.Where(run => Math.Abs(run.PdfCharacterSpacing) > 0.001d).All(run => Math.Abs(run.LayoutCharacterSpacing) < 0.001d),
            "Expected numbered autofit character spacing to be an emission-only PDF text state.");
        TestAssert.True(
            glyphRuns.Any(run => run.ParagraphBulletKind == "AutoNumber"),
            "Expected the probe to expose auto-numbered paragraph metadata to the emission layer.");
    }

    public static void PptxNumberedNoAutofitFramesEmitOfficeCharacterSpacingTextState()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-noautofit-numbered-tc-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        AssertNumberedFrameCharacterSpacing(glyphRuns);
        TestAssert.True(
            glyphRuns.Where(run => Math.Abs(run.PdfCharacterSpacing) > 0.001d).All(run => Math.Abs(run.LayoutCharacterSpacing) < 0.001d),
            "Expected numbered noAutofit character spacing to be an emission-only PDF text state.");
        TestAssert.True(
            glyphRuns.Any(run => run.ParagraphBulletKind == "AutoNumber") &&
            glyphRuns.Any(run => run.FrameAutofitMode == "noAutofit"),
            "Expected the probe to expose noAutofit auto-numbered paragraph metadata to the emission layer.");
    }

    private static void AssertNumberedFrameCharacterSpacing(PptxTextGlyphRunSnapshot[] glyphRuns)
    {
        // These Aptos fixtures can wrap differently when the host substitutes a font.
        // Check every emitted run in each frame instead of assuming a fixed line count.
        IGrouping<int, PptxTextGlyphRunSnapshot>[] frames = glyphRuns.GroupBy(run => run.FrameIndex).OrderBy(frame => frame.Key).ToArray();
        TestAssert.Equal(3, frames.Length);
        double[] expectedSpacing = [0d, -0.048d, -0.024d];
        int[] expectedParagraphCounts = [2, 4, 4];
        for (int index = 0; index < frames.Length; index++)
        {
            TestAssert.Equal(index, frames[index].Key);
            TestAssert.Equal(expectedParagraphCounts[index], frames[index].Select(run => run.ParagraphIndex).Distinct().Count());
            TestAssert.True(
                frames[index].All(run => Math.Abs(run.PdfCharacterSpacing - expectedSpacing[index]) < 0.001d),
                $"Expected every glyph run in frame {index} to use PDF character spacing {expectedSpacing[index]}.");
        }
    }

    public static void PptxNumberedAutofitRunSplitUsesContinuationCharacterSpacingTextState()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-numbered-run-split-tc-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        Dictionary<double, int> buckets = glyphRuns
            .GroupBy(run => Math.Round(run.PdfCharacterSpacing, 3))
            .ToDictionary(group => group.Key, group => group.Count());

        TestAssert.Equal(17, buckets[-0.024d]);
        TestAssert.True(
            glyphRuns.All(run => Math.Abs(run.LayoutCharacterSpacing) < 0.001d),
            "Expected the run-split numbered autofit probe to keep OOXML layout tracking at zero.");
        TestAssert.True(
            glyphRuns.Any(run => run.ParagraphAutoNumberStartAt == 2) &&
            glyphRuns.Any(run => run.ParagraphBulletKind == "None"),
            "Expected the dense numbered-autofit branch to be driven by explicit numbering followed by body continuations.");
    }

    public static void PptxTextLeadingSpaceAfterStyleBoundaryUsesHiddenAdvance()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/>
                      <a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha</a:t></a:r>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t xml:space="preserve"> beta</a:t></a:r>
                      </a:p>
                      <a:p>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t xml:space="preserve">Alpha </a:t></a:r>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>beta</a:t></a:r>
                      </a:p>
                      <a:p>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha</a:t></a:r>
                        <a:r><a:rPr sz="1800" b="1"><a:latin typeface="Arial"/></a:rPr><a:t xml:space="preserve"> beta</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFlowRunSnapshot[] flowRuns = PptxRenderer.InspectTextFlow(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .ToArray();

        PptxTextFlowRunSnapshot sameStyleLeadingSpace = flowRuns.First(run => run.SourceText == " beta" && Math.Abs(run.FontSize - 18d) < 0.01d);
        TestAssert.Equal(" beta", string.Concat(sameStyleLeadingSpace.Segments.Where(segment => segment.Draw).Select(segment => segment.Text)));
        TestAssert.True(!sameStyleLeadingSpace.Segments.Any(segment => segment.Kind == "HiddenAdvance"), "Expected same-style split runs to avoid hidden leading-space advance.");

        PptxTextFlowRunSnapshot styleBoundaryLeadingSpace = flowRuns.Last(run => run.SourceText == " beta");
        TestAssert.True(styleBoundaryLeadingSpace.Segments.Any(segment => segment.Kind == "HiddenAdvance" && segment.AdvanceText == " "), "Expected style-boundary leading space to become measured hidden advance.");
        TestAssert.True(styleBoundaryLeadingSpace.Segments.Any(segment => segment.Kind == "Text" && segment.Text == "beta"), "Expected style-boundary visible text to start after the hidden leading space.");

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        TestAssert.True(
            glyphRuns.Any(run => run.ParagraphIndex == 0 && run.Text == "Alpha beta"),
            "Expected same-style authored runs to merge into one Office-compatible PDF text operation.");
        PptxTextGlyphRunSnapshot boldBeta = glyphRuns.Single(run => run.ParagraphIndex == 2 && run.Text == "beta");
        PptxTextGlyphRunSnapshot precedingAlpha = glyphRuns.Single(run => run.ParagraphIndex == 2 && run.Text == "Alpha");
        TestAssert.True(boldBeta.X > precedingAlpha.X + precedingAlpha.Width, "Expected style-boundary glyph run to start after the hidden leading-space advance.");
        TestAssert.True(!glyphRuns.Any(run => run.ParagraphIndex == 2 && run.Text == " beta"), "Expected emitted glyph runs to omit the style-boundary leading space.");
    }

    public static void PptxTypographyTextHyphenBoundariesRemainSeparateSpans()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-punctuation-boundaries.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        string[] firstLineTexts = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .First()
            .Spans
            .Select(span => span.Text)
            .Take(3)
            .ToArray();

        TestAssert.Equal("SKU", firstLineTexts[0]);
        TestAssert.Equal("-", firstLineTexts[1]);
        TestAssert.True(firstLineTexts[2].StartsWith("123", StringComparison.Ordinal), "Expected text after the hyphen to remain a separate positioned span for Office-style PDF text operations.");
    }

    public static void PptxTypographyTextEmDashBoundariesRemainSeparateSpans()
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
                        <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:nvPr/></p:nvSpPr>
                        <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm></p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Plan &#x2014; execute</a:t></a:r></a:p>
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

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        string[] firstLineTexts = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .First()
            .Spans
            .Select(span => span.Text)
            .Take(3)
            .ToArray();

        TestAssert.Equal("Plan ", firstLineTexts[0]);
        TestAssert.Equal("\u2014", firstLineTexts[1]);
        TestAssert.Equal("execute", firstLineTexts[2]);
    }

    public static void PptxSyntheticStyledTextProducesStyleOperators()
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
                        <a:pPr algn="ctr"/>
                        <a:r>
                          <a:rPr sz="2400" b="1" i="1" u="sng"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr>
                          <a:t>Styled text</a:t>
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
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains(" re f", pdf);
        TestAssert.True(PptxTests.CountOccurrences(pdf, " TJ") >= 1, "Expected styled text to be emitted.");
    }

    public static void PptxMongolianVerticalEmitsClockwiseRowMatrices()
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
                    <p:spPr><a:xfrm><a:off x="2794000" y="762000"/><a:ext cx="6647974" cy="276999"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="mongolianVert" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>AB</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(2, PptxTests.CountOccurrences(pdf, "0 -1 1 0 "));
        TestAssert.Equal(0, PptxTests.CountOccurrences(pdf, " cm"));
        TestAssert.Contains("0 -1 1 0 226.15 480 Tm", pdf);
        TestAssert.Contains("0 -1 1 0 254.95 480 Tm", pdf);
    }

    public static void PptxVerticalAnchorLadderIsLinear()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        // One-line vertical runs stack 28.8pt per line at 24pt; an 87.24pt-wide shape with
        // 14.4pt side insets leaves a 58.44pt stacking box, so middle and bottom anchors sit
        // exactly half and full slack past top: (58.44 - 28.8) / 2 = 14.82pt per step. The sign
        // is negative in layout Tm-Y: the content is smaller than the box, so the positive
        // offset starts lines later in flow (smaller Tm-Y, mirrored to larger device-X).
        double topY = RenderVerticalAnchorProbe(null);
        double middleY = RenderVerticalAnchorProbe("ctr");
        double bottomY = RenderVerticalAnchorProbe("b");
        TestAssert.True(Math.Abs((middleY - topY) + 14.82d) < 0.05d, "Expected the middle anchor one half-slack past top.");
        TestAssert.True(Math.Abs((bottomY - middleY) + 14.82d) < 0.05d, "Expected the bottom anchor one full slack past top.");

        static double RenderVerticalAnchorProbe(string? anchor)
        {
            string anchorAttribute = anchor is null ? string.Empty : " anchor=\"" + anchor + "\"";
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
                        <p:spPr><a:xfrm><a:off x="3311604" y="1219200"/><a:ext cx="1107996" cy="707886"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                        <p:txBody>
                          <a:bodyPr vert="vert" lIns="182880" tIns="45720" rIns="182880" bIns="45720"__ANCHOR__ rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                          <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>A</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp></p:spTree></p:cSld>
                    </p:sld>
                    """.Replace("__ANCHOR__", anchorAttribute)
            });
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

            OoxPdfConverter.Convert(input, output);

            string pdf = File.ReadAllText(output, Encoding.ASCII);
            MatchCollection matrices = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ ([0-9.]+) Tm");
            TestAssert.Equal(1, matrices.Count);
            return double.Parse(matrices[0].Groups[1].Value, CultureInfo.InvariantCulture);
        }
    }
    // RV03: split words break at word boundaries, not run seams.
    public static void FragmentedWordBreaksAtWordBoundary()
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
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1574800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>alpha base</a:t></a:r><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>line beta</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        string[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Select(line => string.Concat(line.Spans.Select(span => span.Text)))
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("alpha ", lines[0]);
        TestAssert.Equal("baseline beta", lines[1]);
    }

    public static void PptxVerticalAutoFitOverflowKeepsWrappedFullSizeText()
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
                    <p:spPr><a:xfrm><a:off x="2794000" y="762000"/><a:ext cx="6647974" cy="276999"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="vert" lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Vertical Text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        // Eight strict-greedy runs at full size; legacy shrinks to one 3.96pt run.
        TestAssert.Equal(8, PptxTests.CountOccurrences(pdf, " TJ"));
        TestAssert.Contains(" 24 Tf", pdf);
    }
}
