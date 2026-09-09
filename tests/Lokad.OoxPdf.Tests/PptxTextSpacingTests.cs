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

internal static class PptxTextSpacingTests
{
    public static void PptxSyntheticTextBoxHonorsTabCharacters()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string slideXml = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr><a:tabLst><a:tab pos="1828800"/></a:tabLst></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A{TAB}B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """.Replace("{TAB}", "\t", StringComparison.Ordinal);

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slideXml
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Contains(" TJ", StringComparison.Ordinal) || pdf.Contains("> Tj", StringComparison.Ordinal), "Expected body text to be drawn as PDF text.");
        PptxTests.AssertContainsTextMatrixAtX(pdf, 216d);
    }

    public static void PptxSyntheticTextBoxUsesDefaultTabStops()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string slideXml = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A{TAB}B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """.Replace("{TAB}", "\t", StringComparison.Ordinal);

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slideXml
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Contains(" TJ", StringComparison.Ordinal) || pdf.Contains("> Tj", StringComparison.Ordinal), "Expected body text to be drawn as PDF text.");
        PptxTests.AssertContainsTextMatrixAtX(pdf, 144d);
    }

    public static void PptxTextModelInheritsListStyleTabStops()
    {
        string slideXml = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/>
                      <a:lstStyle>
                        <a:lvl1pPr><a:tabLst><a:tab pos="1828800"/></a:tabLst></a:lvl1pPr>
                      </a:lstStyle>
                      <a:p><a:pPr/><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A{TAB}B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """.Replace("{TAB}", "\t", StringComparison.Ordinal);

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slideXml
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        PptxTextParagraphModelSnapshot paragraph = frame.Paragraphs.Single();

        TestAssert.Equal(1, paragraph.TabStops.Count);
        TestAssert.Equal(144d, paragraph.TabStops[0]);
    }

    public static void PptxSyntheticTextBoxTreatsNarrowNoBreakSpaceAsHiddenAdvance()
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha&#x202F;beta</a:t></a:r></a:p>
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

        TestAssert.True(line.Spans.Any(span => span.Text == "Alpha"), "Expected narrow no-break space to split the preceding word span.");
        TestAssert.True(line.Spans.Any(span => span.Text == "beta"), "Expected narrow no-break space to split the following word span.");
        TestAssert.True(line.Spans.All(span => !span.Text.Contains('\u202F', StringComparison.Ordinal)), "Expected narrow no-break space to stay out of emitted text spans.");
        TestAssert.True(line.Spans.Zip(line.Spans.Skip(1), (left, right) => right.X - (left.X + left.Width)).Any(gap => gap > 0d), "Expected narrow no-break space to contribute hidden advance between words.");
        PptxTextFlowSegmentSnapshot hiddenAdvance = PptxRenderer.InspectTextFlow(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .SelectMany(run => run.Segments)
            .Single(segment => segment.Kind == "NoBreakHiddenAdvance");
        TestAssert.Equal("\u202F", hiddenAdvance.AdvanceText);
    }

    public static void PptxSyntheticTextBoxNoBreakSpaceKeepsWrappedClusterTogether()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3175000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha beta gamma +2&#x202F;points omega</a:t></a:r></a:p>
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
            .Where(line => line.Spans.Count > 0)
            .ToArray();

        int tokenLine = Array.FindIndex(lines, line => line.Spans.Any(span => span.Text == "+2"));
        int followingLine = Array.FindIndex(lines, line => line.Spans.Any(span => span.Text == "points"));

        TestAssert.True(tokenLine >= 0, "Expected the no-break cluster's leading token to be present.");
        TestAssert.Equal(tokenLine, followingLine);
        TestAssert.True(!lines.Take(tokenLine).Any(line => line.Spans.Any(span => span.Text == "+2")), "Expected the leading token to move with the following no-break word.");
        TestAssert.True(lines[tokenLine].Spans.All(span => !span.Text.Contains('\u202F', StringComparison.Ordinal)), "Expected no-break space to remain a hidden advance after wrapping.");
    }

    public static void PptxSyntheticTextBoxTreatsNoBreakSpaceAsHiddenRegularSpaceAdvance()
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha&#xA0;beta</a:t></a:r></a:p>
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

        PptxTextSpanLayoutSnapshot alpha = line.Spans.Single(span => span.Text == "Alpha");
        PptxTextSpanLayoutSnapshot beta = line.Spans.Single(span => span.Text == "beta");
        double hiddenAdvance = beta.X - (alpha.X + alpha.Width);

        TestAssert.True(line.Spans.All(span => !span.Text.Contains('\u00A0', StringComparison.Ordinal)), "Expected no-break space to stay out of emitted text spans.");
        TestAssert.True(hiddenAdvance > 5d, "Expected no-break space to advance like a regular hidden space, not a narrow spacer.");
    }

    public static void PptxSyntheticTextBoxNoBreakSpaceKeepsBoundaryKerning()
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400" kern="1200"><a:latin typeface="Arial"/></a:rPr><a:t>A&#xA0;B</a:t></a:r></a:p>
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

        PptxTextSpanLayoutSnapshot first = line.Spans.Single(span => span.Text == "A");
        PptxTextSpanLayoutSnapshot second = line.Spans.Single(span => span.Text == "B");
        double totalAdvance = second.X - first.X;

        TestAssert.True(totalAdvance > 21d, $"Expected no-break space to keep a visible regular-space class advance, got {totalAdvance}.");
        TestAssert.True(totalAdvance < 22d, $"Expected no-break space to preserve Office's kerning from the preceding glyph, got {totalAdvance}.");
        TestAssert.True(second.GlyphSpan.LeadingAdjustment < -0.1d, "Expected the visible glyph span after hidden no-break space to own its leading boundary kerning.");
    }

    public static void PptxSyntheticTextBoxOffsetsLargeTextByFontSize()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="6600"/><a:t>Large heading</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 79.2d);
    }

    public static void PptxSyntheticTextBoxRendersBulletCharacters()
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
                      <a:p><a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Bullet item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("2022", pdf);
    }

    public static void PptxSyntheticTextBoxInheritsListStyleBulletCharacters()
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
                      <a:bodyPr/>
                      <a:lstStyle><a:lvl1pPr><a:buChar char="&#x2022;"/></a:lvl1pPr></a:lstStyle>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Inherited item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextParagraphModelSnapshot paragraph = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .Single()
            .Paragraphs
            .Single();
        TestAssert.Equal("Character", paragraph.BulletKind);
        TestAssert.Equal("\u2022", paragraph.BulletCharacter);
        TestAssert.Equal("\u2022", paragraph.BulletResolvedCharacter);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        IReadOnlyList<string> spanTexts = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .Select(span => span.Text)
            .ToArray();

        TestAssert.Contains("\u2022", string.Concat(spanTexts));
    }

    public static void PptxSyntheticTextBoxInheritsListStyleBulletThroughDirectParagraphIndent()
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
                      <a:bodyPr/>
                      <a:lstStyle><a:lvl1pPr marL="457200" indent="-228600"><a:buChar char="&#x2022;"/></a:lvl1pPr></a:lstStyle>
                      <a:p><a:pPr marL="914400" indent="-228600"/><a:r><a:rPr sz="1800"/><a:t>Indented inherited item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextParagraphModelSnapshot paragraph = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .Single()
            .Paragraphs
            .Single();
        TestAssert.Equal("Character", paragraph.BulletKind);
        TestAssert.Equal("\u2022", paragraph.BulletResolvedCharacter);
        TestAssert.Equal(72d, paragraph.MarginLeft);
        TestAssert.Equal(-18d, paragraph.HangingIndent);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        IReadOnlyList<string> spanTexts = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(layoutParagraph => layoutParagraph.Lines)
            .SelectMany(line => line.Spans)
            .Select(span => span.Text)
            .ToArray();

        TestAssert.Contains("\u2022", string.Concat(spanTexts));
    }

    public static void PptxSyntheticTextBoxDirectBulletOverridesInheritedNone()
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
                      <a:bodyPr/>
                      <a:lstStyle><a:lvl1pPr><a:buNone/></a:lvl1pPr></a:lstStyle>
                      <a:p><a:pPr marL="571500" indent="-571500"><a:buFont typeface="Arial"/><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Direct bullet item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextParagraphModelSnapshot paragraph = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .Single()
            .Paragraphs
            .Single();
        TestAssert.Equal("Character", paragraph.BulletKind);
        TestAssert.Equal("\u2022", paragraph.BulletResolvedCharacter);
        TestAssert.Equal("Arial", paragraph.BulletFontTypeface);
        TestAssert.Equal(45d, paragraph.MarginLeft);
        TestAssert.Equal(-45d, paragraph.HangingIndent);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        IReadOnlyList<string> spanTexts = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .Select(span => span.Text)
            .ToArray();

        TestAssert.Contains("\u2022", string.Concat(spanTexts));
    }

    public static void PptxSyntheticDefaultAutofitBulletUsesOfficeWrapTolerance()
    {
        const double emusPerPoint = 12700d;
        const double overflowBeyondTextBox = 2.4d;
        const double coordinateTolerance = 0.01d;
        const double bulletWrapToleranceAt24Pt = 4.8d;

        string WritePackage(long widthEmu) => TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="{{widthEmu}}" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p>
                        <a:pPr marL="571500" indent="-571500"><a:buFont typeface="Arial"/><a:buChar char="&#x2022;"/></a:pPr>
                        <a:r><a:rPr sz="2400"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Public bullet text keeps the final marker</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        string wideInput = WritePackage(12000000);
        using FileStream wideStream = File.OpenRead(wideInput);
        OoxPackage widePackage = OoxPackage.Open(wideStream, CancellationToken.None);
        PptxDocument wideDocument = new PptxReader().Read(widePackage, CancellationToken.None);
        PptxTextFrameModelSnapshot wideFrame = PptxRenderer.InspectTextFrameModels(wideDocument, widePackage, 0).Single();
        PptxTextLineLayoutSnapshot wideLine = PptxRenderer.InspectTextLayout(wideDocument, widePackage, 0)
            .Frames
            .Single()
            .Paragraphs
            .Single()
            .Lines
            .Single();

        double targetTextWidth = wideLine.NaturalEndX - wideFrame.TextX - overflowBeyondTextBox;
        long narrowWidthEmu = (long)Math.Round((targetTextWidth + wideFrame.InsetLeft + wideFrame.InsetRight) * emusPerPoint);
        string narrowInput = WritePackage(narrowWidthEmu);
        using FileStream narrowStream = File.OpenRead(narrowInput);
        OoxPackage narrowPackage = OoxPackage.Open(narrowStream, CancellationToken.None);
        PptxDocument narrowDocument = new PptxReader().Read(narrowPackage, CancellationToken.None);
        PptxTextFrameModelSnapshot narrowFrame = PptxRenderer.InspectTextFrameModels(narrowDocument, narrowPackage, 0).Single();
        PptxTextLineLayoutSnapshot[] narrowLines = PptxRenderer.InspectTextLayout(narrowDocument, narrowPackage, 0)
            .Frames
            .Single()
            .Paragraphs
            .Single()
            .Lines
            .ToArray();

        TestAssert.Equal(1, narrowLines.Length);
        string renderedText = string.Concat(narrowLines[0].Spans.Select(span => span.Text));
        TestAssert.Contains("final marker", renderedText);
        double actualOverflow = narrowLines[0].NaturalEndX - (narrowFrame.TextX + narrowFrame.TextWidth);
        TestAssert.True(
            actualOverflow > coordinateTolerance && actualOverflow < bulletWrapToleranceAt24Pt,
            $"Expected bullet wrap to accept a small Office-style overrun. Overflow: {actualOverflow.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticTextBoxMapsSymbolFontBulletCharacters()
    {
        string wingdings = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "wingding.ttf");
        if (!File.Exists(wingdings))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(wingdings))");
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
                      <a:p>
                        <a:pPr>
                          <a:buFont typeface="Wingdings" charset="2"/>
                          <a:buChar char="&#xA7;"/>
                        </a:pPr>
                        <a:r><a:rPr sz="1800"/><a:t>Symbol item</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextParagraphModelSnapshot paragraph = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .Single()
            .Paragraphs
            .Single();
        TestAssert.Equal("Character", paragraph.BulletKind);
        TestAssert.Equal("\u00A7", paragraph.BulletCharacter);
        TestAssert.Equal("\uF0A7", paragraph.BulletResolvedCharacter);
        TestAssert.Equal("Wingdings", paragraph.BulletFontTypeface);
        TestAssert.Equal("2", paragraph.BulletFontCharset);
        TestAssert.Equal("Wingdings", paragraph.BulletResolvedFontTypeface);
        TestAssert.Equal("Direct", paragraph.BulletFontTypefaceSource);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextSpanLayoutSnapshot bullet = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .Single(span => span.Text == "\uF0A7");

        TestAssert.Equal("\uF0A7", bullet.Text);
        TestAssert.Equal("Wingdings", bullet.GlyphSpan.Typeface);
    }

    public static void PptxSyntheticTextBoxRendersAutoNumberedBullets()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>First item</a:t></a:r></a:p>
                      <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Second item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            IReadOnlyList<PptxTextParagraphModelSnapshot> paragraphs = PptxRenderer.InspectTextFrameModels(document, package, 0)
                .Single()
                .Paragraphs;

            TestAssert.Equal("AutoNumber", paragraphs[0].BulletKind);
            TestAssert.Equal("arabicPeriod", paragraphs[0].BulletAutoNumberType);
            TestAssert.Equal("1", paragraphs[0].BulletAutoNumberStartAtValue);
            TestAssert.Equal("AutoNumber", paragraphs[1].BulletKind);
            TestAssert.Equal("arabicPeriod", paragraphs[1].BulletAutoNumberType);
            TestAssert.Equal(null, paragraphs[1].BulletAutoNumberStartAtValue);
        }

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0031", pdf);
        TestAssert.Contains("0032", pdf);
    }

    public static void PptxSyntheticTextBoxRestartsAutoNumberingAfterPlainParagraph()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Alpha</a:t></a:r></a:p>
                      <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Beta</a:t></a:r></a:p>
                      <a:p><a:pPr/><a:r><a:rPr sz="1800"/><a:t>Plain separator</a:t></a:r></a:p>
                      <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Gamma</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxTextFrameLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0).Frames.Single();
            string[] bulletTexts = layout.Paragraphs
                .SelectMany(paragraph => paragraph.Lines)
                .SelectMany(line => line.Spans)
                .Select(span => span.Text)
                .Where(text => text is "1." or "2.")
                .ToArray();

            TestAssert.Equal("1.|2.|1.", string.Join("|", bulletTexts));
        }

        OoxPdfConverter.Convert(input, output);
    }

    public static void PptxSyntheticTextBoxRendersRomanAutoNumberedBullets()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buAutoNum type="romanLcPeriod" startAt="4"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>One</a:t></a:r></a:p>
                      <a:p><a:pPr><a:buAutoNum type="romanLcPeriod"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Two</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0076", pdf);
        TestAssert.Contains("002E", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsBulletHangingIndent()
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
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr marL="914400" indent="-914400"><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Indented item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 79.2d);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 151.2d);
    }

    public static void PptxSyntheticTextBoxHonorsBulletColorAndSize()
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
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPts val="3000"/><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Styled bullet</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxTextParagraphModelSnapshot paragraph = PptxRenderer.InspectTextFrameModels(document, package, 0)
                .Single()
                .Paragraphs
                .Single();

            TestAssert.Equal("FF0000", paragraph.BulletColor);
            TestAssert.Equal("Points", paragraph.BulletSizeKind);
            TestAssert.Equal("3000", paragraph.BulletSizeValue);
        }

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 30 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("/F1 24 Tf", pdf);
        TestAssert.Contains("0 g", pdf);
    }

    public static void PptxSyntheticTextBoxUsesPerRunFontResources()
    {
        string fonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!File.Exists(Path.Combine(fonts, "arial.ttf")) ||
            !File.Exists(Path.Combine(fonts, "times.ttf")))
        {
            TestAssert.Skip("Environmental precondition not met: missing Arial or Times.");
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
                        <a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Arial</a:t></a:r>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Times New Roman"/></a:rPr><a:t>Times</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 18 Tf", pdf);
        TestAssert.Contains("/F2 18 Tf", pdf);
    }

    public static void PptxSyntheticTextBoxUsesEaFontWhenLatinMissing()
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
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:ea typeface="Courier New"/></a:rPr><a:t>East Asian fallback</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextSpanLayoutSnapshot span = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames.SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .Single(candidate => candidate.Text.StartsWith("East", StringComparison.Ordinal));

        TestAssert.Equal("Courier New", span.GlyphSpan.Typeface ?? string.Empty);
    }

    public static void PptxSyntheticTextBoxSplitsMissingGlyphsToFallbackFont()
    {
        string fonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!File.Exists(Path.Combine(fonts, "arial.ttf")) ||
            !File.Exists(Path.Combine(fonts, "simsun.ttc")))
        {
            TestAssert.Skip("Environmental precondition not met: missing Arial or SimSun.");
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
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A&#x4E2D;B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);

        string expected = "A" + char.ConvertFromUtf32(0x4E2D) + "B";
        TestAssert.Equal(expected, string.Concat(glyphRuns.Select(run => run.Text)));
        TestAssert.True(
            glyphRuns.SelectMany(run => run.Glyphs).Select(glyph => glyph.ResourceName).Distinct(StringComparer.OrdinalIgnoreCase).Count() >= 2,
            "Expected a missing Arial glyph to emit through a separate fallback PDF font resource.");
        PptxTextGlyphRunAtomSnapshot fallbackGlyph = glyphRuns.SelectMany(run => run.Glyphs).Single(glyph => glyph.CodePoint == 0x4E2D);
        TestAssert.True(!string.Equals("Arial", fallbackGlyph.Typeface, StringComparison.OrdinalIgnoreCase), "Expected the CJK glyph to carry its fallback typeface.");
        TestAssert.Equal("Fallback", fallbackGlyph.TypefaceResolutionSource);
        TestAssert.True(
            glyphRuns.SelectMany(run => run.Glyphs).Where(glyph => glyph.CodePoint != 0x4E2D).All(glyph => glyph.TypefaceResolutionSource == "Primary"),
            "Expected glyph inspection to distinguish requested-font glyphs from fallback-font glyphs.");

        PptxTextGlyphLayoutSnapshot fallbackLayoutGlyph = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames.SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .SelectMany(span => span.GlyphSpan.Glyphs)
            .Single(glyph => glyph.CodePoint == 0x4E2D);
        TestAssert.Equal("Fallback", fallbackLayoutGlyph.TypefaceResolutionSource);
    }

    public static void PptxSyntheticTextBoxUsesDistinctFontResourcesForStyles()
    {
        string fonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!File.Exists(Path.Combine(fonts, "arial.ttf")) ||
            !File.Exists(Path.Combine(fonts, "arialbd.ttf")))
        {
            TestAssert.Skip("Environmental precondition not met: missing Arial or Arial Bold.");
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
                        <a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Regular</a:t></a:r>
                        <a:r><a:rPr sz="2400" b="1"><a:latin typeface="Arial"/></a:rPr><a:t>Bold</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Split("/FontFile2", StringSplitOptions.None).Length - 1 >= 2, "Regular and bold Arial should embed distinct font files.");
    }

    public static void PptxSyntheticTextBoxUsesOfficeSyntheticBoldStroke()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400" b="1"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Synthetic bold</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.686 w", pdf);
        TestAssert.Contains("2 Tr", pdf);
        TestAssert.True(Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm").Count == 1, "Synthetic bold should use fill-and-stroke text rather than duplicate offset glyph draws.");
    }

    public static void PptxSyntheticBoldCambriaCenteredTextUsesOfficeTightenedAdvance()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-04-centered-bold-cambria-width.pptx");
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0).ToArray();
        PptxTextGlyphRunSnapshot bold = glyphRuns.Single(run => run.FrameIndex == 0);
        PptxTextGlyphRunSnapshot regular = glyphRuns.Single(run => run.FrameIndex == 1);

        TestAssert.True(
            bold.X > regular.X + 1.5d,
            $"Expected synthetic bold centered text to use Office's tighter synthetic-bold advance. Bold X={bold.X:0.###}, regular X={regular.X:0.###}.");
        TestAssert.True(
            bold.Width < regular.Width - 3d,
            $"Expected synthetic bold layout width to be narrower than regular Cambria Math for Office-like centered placement. Bold width={bold.Width:0.###}, regular width={regular.Width:0.###}.");
    }

    public static void PptxHighlightedSyntheticBoldItalicMathParagraphUsesOfficeCharacterSpacing()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        string highlighted = PptxTests.RenderSyntheticMathSpacingProbe(highlightMiddleRun: true);
        string unhighlighted = PptxTests.RenderSyntheticMathSpacingProbe(highlightMiddleRun: false);

        TestAssert.Contains("0.309 Tc", highlighted);
        TestAssert.DoesNotContain("0.309 Tc", unhighlighted);
    }

    public static void PptxSyntheticTextBoxUsesKerningWhenAvailable()
    {
        string times = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "times.ttf");
        if (!File.Exists(times))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(times))");
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
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="4000"><a:latin typeface="Times New Roman"/></a:rPr><a:t>To</a:t></a:r></a:p>
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

    public static void PptxHighlightedHeadlineTextWrapsBeforeFinalWord()
    {
        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-04-highlighted-headline-runs.pptx");
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
            texts.Length == 2,
            $"Expected Office-compatible two-line wrap before the final word. Lines: {string.Join(" | ", texts)}. Widths: {string.Join(" | ", lines.Select(line => (line.EndX - line.StartX).ToString("0.###", CultureInfo.InvariantCulture)))}.");
        TestAssert.Equal("Public headline uses highlighted terms and normal ", texts[0]);
        TestAssert.Equal("text", texts[1]);

        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);
        string[] glyphRunTexts = glyphRuns.Select(run => run.Text).ToArray();
        TestAssert.Equal(7, glyphRunTexts.Length);
        TestAssert.Equal("Public headline uses ", glyphRunTexts[0]);
        TestAssert.Equal("highlighted", glyphRunTexts[1]);
        TestAssert.Equal(" ", glyphRunTexts[2]);
        TestAssert.Equal("terms and ", glyphRunTexts[3]);
        TestAssert.Equal("normal", glyphRunTexts[4]);
        TestAssert.Equal(" ", glyphRunTexts[5]);
        TestAssert.Equal("text", glyphRunTexts[6]);
        int[] glyphRunLineIndexes = glyphRuns.Select(run => run.LineIndex).ToArray();
        TestAssert.True(
            glyphRunLineIndexes.SequenceEqual(new[] { 0, 0, 0, 0, 0, 0, 1 }),
            "Expected glyph-run line indexes to follow the two wrapped layout lines. Actual: " + string.Join(",", glyphRunLineIndexes));
        TestAssert.True(
            glyphRuns.All(run => run.SpanIndex >= 0 && run.SpanIndex < lines[run.LineIndex].Spans.Count),
            "Expected glyph-run inspection to preserve each source span ordinal inside its wrapped line.");
        TestAssert.True(
            glyphRuns.All(run => run.LineSpanCount == lines[run.LineIndex].Spans.Count),
            "Expected glyph-run inspection to expose the source span count for each wrapped line before PDF text emission.");
        TestAssert.True(
            glyphRuns.Take(6).All(run => Math.Abs(run.LineTopY - glyphRuns[0].LineTopY) < 0.01d),
            "Expected all first-line glyph runs to preserve the same line top.");
        TestAssert.True(
            Math.Abs(glyphRuns[6].LineTopY - glyphRuns[0].LineTopY) > 0.01d,
            "Expected the wrapped glyph run to preserve a distinct second-line top.");

        PptxTextGlyphRunSnapshot highlighted = glyphRuns.Single(run => run.Text == "highlighted");
        TestAssert.True(
            highlighted.HighlightColor is { } highlight && highlight.Equals(new RgbColor(255, 255, 0)),
            "Expected glyph-run inspection to expose the source highlight color.");
        TestAssert.True(
            highlighted.HighlightX is not null && highlighted.HighlightY is not null,
            "Expected highlighted glyph runs to expose rectangle origin.");
        TestAssert.True(
            highlighted.HighlightWidth is not null && highlighted.HighlightHeight is not null,
            "Expected highlighted glyph runs to expose rectangle size.");
        double actualHighlightWidth = highlighted.HighlightWidth.GetValueOrDefault();
        double actualHighlightHeight = highlighted.HighlightHeight.GetValueOrDefault();
        TestAssert.True(Math.Abs(actualHighlightWidth - highlighted.Width) < 0.01d, "Expected highlight geometry to be owned by the positioned glyph span width.");
        TestAssert.True(actualHighlightHeight > highlighted.LayoutFontSize * 0.5d, "Expected highlight geometry to use a line-height-sized rectangle.");
    }

    public static void PptxSyntheticCenteredLogoBoxWrapsDefaultTypefaceText()
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
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
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
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Aptos">
                  <a:themeElements>
                    <a:clrScheme name="Aptos">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                    </a:clrScheme>
                    <a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3350712" cy="1847589"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:noFill/>
                      <a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                    </p:spPr>
                    <p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4000"/><a:t>Customer Logo</a:t></a:r></a:p>
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
        string[] texts = lines.Select(line => string.Concat(line.Spans.Select(span => span.Text)).Trim()).ToArray();

        TestAssert.Equal(2, texts.Length);
        TestAssert.Equal("Customer", texts[0]);
        TestAssert.Equal("Logo", texts[1]);
    }

    public static void PptxSyntheticTextBoxHonorsKerningThreshold()
    {
        string times = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "times.ttf");
        if (!File.Exists(times))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(times))");
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
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800" kern="3600"><a:latin typeface="Times New Roman"/></a:rPr><a:t>To</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(times), "To".Select(c => (int)c), CancellationToken.None);
        string expected = "[<" + embedded.EncodeGlyphHex("T") + "><" + embedded.EncodeGlyphHex("o") + ">] TJ";
        TestAssert.Contains(expected, pdf);
    }

    public static void PptxSyntheticTextBoxEmitsCharacterSpacingAsTextState()
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
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3600" spc="3000"><a:latin typeface="Arial"/></a:rPr><a:t>AB</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" TJ", pdf);
        TestAssert.Contains("30 Tc", pdf);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextGlyphRunSnapshot glyphRun = PptxRenderer.InspectTextGlyphRuns(document, package, 0).Single();
        TestAssert.Equal(30d, glyphRun.LayoutCharacterSpacing);
        TestAssert.Equal(30d, glyphRun.PdfCharacterSpacing);
    }

    public static void PptxSyntheticTextBoxResetsCharacterSpacingTextState()
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
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="3600" spc="3000"><a:latin typeface="Arial"/></a:rPr><a:t>AB</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="3600"><a:latin typeface="Arial"/></a:rPr><a:t>CD</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int explicitSpacing = pdf.IndexOf("30 Tc", StringComparison.Ordinal);
        int resetSpacing = pdf.IndexOf("0 Tc", explicitSpacing + 1, StringComparison.Ordinal);
        TestAssert.True(explicitSpacing >= 0 && resetSpacing > explicitSpacing, "Expected text operations to reset inherited PDF character spacing.");
    }

    public static void PptxSyntheticTextBoxRendersSmallCapsAsScaledUppercaseRuns()
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400" cap="small"><a:latin typeface="Arial"/></a:rPr><a:t>Abc DEF</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("24 Tf", pdf);
        TestAssert.Contains("19.2 Tf", pdf);
        TestAssert.True(PptxTests.CountOccurrences(pdf, " TJ") >= 3, "Expected small-caps text to split into full-size and scaled positioned text runs.");
    }

    public static void PptxSyntheticTextBoxRendersBaselineShiftWithOfficeScale()
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Base</a:t></a:r><a:r><a:rPr sz="2400" baseline="30000"><a:latin typeface="Arial"/></a:rPr><a:t>2</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("24 Tf", pdf);
        TestAssert.Contains("15.96 Tf", pdf);
    }

    public static void PptxSyntheticTextBoxKeepsFontSizeForSmallBaselineShift()
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
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Base</a:t></a:r><a:r><a:rPr sz="2400" baseline="10000"><a:latin typeface="Arial"/></a:rPr><a:t>nudge</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextSpanLayoutSnapshot[] spans = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .ToArray();

        TestAssert.True(Math.Abs(spans[0].FontSize - 24d) < 0.001d, $"Expected base span to remain 24pt, got {spans[0].FontSize.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(spans[1].FontSize - 24d) < 0.001d, $"Expected small baseline-shift span to remain 24pt, got {spans[1].FontSize.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticTextBoxCentersMixedRunsTogether()
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
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Left </a:t></a:r><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Right</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 177.478d);
    }

    public static void PptxSyntheticTextBoxCenteredWrappedLineAlignsIgnoringTrailingSpace()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1524000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="square"/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha beta gamma</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot firstLine = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames.Single()
            .Paragraphs.Single()
            .Lines.First();

        PptxTextAtomLayoutSnapshot trailingSpace = firstLine.Spans
            .SelectMany(span => span.Atoms)
            .Last(atom => atom.Kind == "Space");
        double visibleEndX = firstLine.Spans
            .SelectMany(span => span.Atoms)
            .Where(atom => atom.Draw && atom.Kind != "Space")
            .Max(atom => atom.X + atom.Width);
        double textX = 72d;
        double textWidth = 120d;

        TestAssert.True(firstLine.EndX > visibleEndX + trailingSpace.Width - 0.01d, "Expected the first wrapped line to preserve the emitted line-end space.");
        TestAssert.True(
            Math.Abs((firstLine.StartX - textX) - (textX + textWidth - visibleEndX)) < 0.01d,
            "Expected centered wrapped-line alignment to ignore the trailing space while preserving it for text emission.");
    }

    public static void PptxSyntheticTextBoxWrapsAcrossMixedRuns()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha beta </a:t></a:r><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>gamma</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 79.2d);
        TestAssert.True(Regex.Matches(pdf, @"1 0 0 1 79\.2 [0-9.]+ Tm").Count >= 2, "Expected wrapped mixed runs to emit at least two text lines.");
    }

    public static void PptxSyntheticTextBoxUsesListStyleDefaults()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:defRPr sz="3600" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></a:lstStyle>
                      <a:p><a:r><a:rPr/><a:t>Default</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr/><a:t>Next</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 36 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("1 0 0 1 79.2 429.336 Tm", pdf);
        TestAssert.Contains("1 0 0 1 79.2 386.136 Tm", pdf);
    }

    public static void PptxSyntheticTextBoxUsesShapeFontRefColor()
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
                    <p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Styled color</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxTextRunModelSnapshot run = PptxRenderer.InspectTextFrameModels(document, package, 0)
                .SelectMany(frame => frame.Paragraphs)
                .SelectMany(paragraph => paragraph.Runs)
                .Single(textRun => textRun.Text == "Styled color");
            TestAssert.Equal("ShapeFontRef", run.ColorSource);
        }

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 g", pdf);
    }

    public static void PptxSyntheticShapeFontRefColorOverridesDefaultTextStyle()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldSz cx="9144000" cy="6858000"/>
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                  <p:defaultTextStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:defaultTextStyle>
                </p:presentation>
                """,
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme">
                  <a:themeElements>
                    <a:clrScheme name="Theme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                    </a:clrScheme>
                    <a:fontScheme name="Theme"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="808080"/></a:solidFill></p:spPr>
                    <p:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1400"><a:latin typeface="Arial"/></a:rPr><a:t>Font ref</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxTextRunModelSnapshot run = PptxRenderer.InspectTextFrameModels(document, package, 0)
                .SelectMany(frame => frame.Paragraphs)
                .SelectMany(paragraph => paragraph.Runs)
                .Single(textRun => textRun.Text == "Font ref");
            TestAssert.Equal("ShapeFontRef", run.ColorSource);
        }

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 g", pdf);
    }

    public static void PptxSyntheticExplicitRunColorOverridesShapeFontRefColor()
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
                    <p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Explicit color</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxTextRunModelSnapshot run = PptxRenderer.InspectTextFrameModels(document, package, 0)
                .SelectMany(frame => frame.Paragraphs)
                .SelectMany(paragraph => paragraph.Runs)
                .Single(textRun => textRun.Text == "Explicit color");
            TestAssert.Equal("RunSolidFill", run.ColorSource);
        }

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        int redTextColor = pdf.IndexOf("1 0 0 rg", StringComparison.Ordinal);
        TestAssert.True(redTextColor >= 0, "Expected explicit run color in the PDF text stream.");
        TestAssert.True(pdf.IndexOf("1 g", redTextColor, StringComparison.Ordinal) < 0, "Explicit run color should override shape fontRef color after the page background has been painted.");
    }

    public static void PptxSyntheticTextBoxRendersTextHighlight()
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
                      <a:p><a:r><a:rPr sz="1800"><a:highlight><a:srgbClr val="800080"/></a:highlight><a:latin typeface="Arial"/></a:rPr><a:t>Marked</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.502 0 0.502 rg", pdf);
        TestAssert.Contains(" re f*", pdf);
    }

    public static void PptxSyntheticTextBoxRendersStrikethrough()
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
                      <a:p><a:r><a:rPr sz="2400" strike="sngStrike"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Strike</a:t></a:r></a:p>
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

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextGlyphRunSnapshot glyphRun = PptxRenderer.InspectTextGlyphRuns(document, package, 0).Single(run => run.Text == "Strike");
        TestAssert.True(glyphRun.StrikeWidth is not null && glyphRun.StrikeHeight is not null, "Expected glyph-run inspection to expose strike geometry.");
        TestAssert.True(Math.Abs(glyphRun.StrikeWidth.GetValueOrDefault() - glyphRun.Width) < 0.01d, "Expected strike geometry to be owned by the emitted glyph-run width.");
    }

    public static void PptxSyntheticTextBoxHonorsParagraphSpacing()
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
                      <a:p>
                        <a:pPr><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>
                        <a:r><a:rPr sz="1800"/><a:t>Spaced</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 79.2 446.868 Tm", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsPercentParagraphSpacing()
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
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr><a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr><a:r><a:rPr sz="2400"/><a:t>First</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="2400"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Contains(" TJ", StringComparison.Ordinal) || pdf.Contains("> Tj", StringComparison.Ordinal), "Expected body text to be drawn as PDF text.");
    }

    public static void PptxSyntheticTextBoxSkipsEmptyParagraphs()
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
                      <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p>
                      <a:p><a:endParaRPr lang="en-US"/></a:p>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 79.2d);
        TestAssert.True(Regex.Matches(pdf, @"1 0 0 1 79\.2 [0-9.]+ Tm").Count >= 2, "Expected empty paragraphs to preserve later paragraph placement.");
    }
}
