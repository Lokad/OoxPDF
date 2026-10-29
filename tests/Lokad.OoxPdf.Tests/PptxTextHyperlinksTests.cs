using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxTextHyperlinksTests
{
    public static void PptxFallbackTextKeepsMeasuredAdvances()
    {
        foreach (string text in new[] { "Link", "A😀B" })
        {
            string input = WriteRunLinkPackage(
                $"""<a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>{text}</a:t></a:r></a:p>""",
                """<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/advances" TargetMode="External"/>""");
            (List<PdfPage> pages, _) = RenderSlidePages(input, new MissingTestFontResolver());
            PdfPage page = pages.Single();
            Match[] shows = Regex.Matches(page.Content, @"1 0 0 1 ([-0-9.]+) [-0-9.]+ Tm\s+<([0-9A-F]{2})> Tj").ToArray();
            Rune[] runes = text.EnumerateRunes().ToArray();
            TestAssert.Equal(runes.Length, shows.Length);
            double firstX = double.Parse(shows[0].Groups[1].Value, CultureInfo.InvariantCulture);
            int codeUnits = 0;
            for (int i = 0; i < runes.Length; i++)
            {
                double x = double.Parse(shows[i].Groups[1].Value, CultureInfo.InvariantCulture);
                TestAssert.True(Math.Abs(x - firstX - codeUnits * 18d * 0.42d) < 0.001d,
                    "Fallback glyphs must retain their measured advances instead of overlapping.");
                codeUnits += runes[i].Utf16SequenceLength;
            }
            TestAssert.True(Math.Abs(page.Annotations.Single().Width - codeUnits * 18d * 0.42d) < 0.001d,
                "The link must cover the same fallback advance as the painted text.");
        }
    }

    public static void PptxFallbackRunHyperlinksFollowClippedText()
    {
        string paragraph = """<a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>X</a:t></a:r></a:p>""";
        string input = WriteRunLinkPackage(string.Concat(Enumerable.Repeat(paragraph, 6)),
            """<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/clipped" TargetMode="External"/>""",
            """<a:bodyPr vertOverflow="clip" lIns="0" rIns="0" tIns="0" bIns="0"/>""");
        (List<PdfPage> pages, _) = RenderSlidePages(input, new MissingTestFontResolver());
        PdfPage page = pages.Single();
        int painted = Regex.Matches(page.Content, @"<58>\s+Tj").Count;
        TestAssert.True(painted > 0 && painted < 6, "The control must paint some rows and omit clipped rows.");
        TestAssert.Equal(painted, page.Annotations.Count);
        TestAssert.True(page.Annotations.All(link => link.Uri == "https://example.invalid/clipped"),
            "Only emitted fallback rows should keep their source link.");
    }

    public static void PptxRunExternalHyperlinkEmitsUriAnnotation()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdRunLink"/></a:rPr><a:t>Link</a:t></a:r></a:p>""",
            slideRelsXml: """<Relationship Id="rIdRunLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/run" TargetMode="External"/>""");

        foreach (IFontResolver resolver in new IFontResolver[] { new LinkTestFontResolver(), new MissingTestFontResolver() })
        {
            (List<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, resolver);

            PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
            TestAssert.Equal("https://example.invalid/run", annotation.Uri);
            TestAssert.True(annotation.Width > 0d, "Run hyperlink annotations should cover the linked text width.");
            TestAssert.True(Math.Abs((resolver is LinkTestFontResolver ? 19.8d : 16.65d) - annotation.Height) < 1e-6d, $"Run hyperlink height should follow the emitted font ascent/descent box, got {annotation.Height}.");
            TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 72d + 432d, "Run hyperlink X extent should stay inside the shape bounds.");
            TestAssert.True(annotation.Y >= 396d && annotation.Y + annotation.Height <= 396d + 72d, "Run hyperlink Y extent should stay inside the shape bounds.");
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable run hyperlink should not emit hyperlink diagnostics.");
        }
    }

    public static void PptxRunHyperlinkBoundaryKeepsAdjacentLinksSeparate()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdLink1"/></a:rPr><a:t>A </a:t></a:r>""" +
                """<a:r><a:rPr sz="1800" u="sng"><a:latin typeface="Link Sans"/></a:rPr><a:t>B</a:t></a:r>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdLink2"/></a:rPr><a:t>C</a:t></a:r></a:p>""",
            slideRelsXml: """<Relationship Id="rIdLink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/first" TargetMode="External"/>""" +
                """<Relationship Id="rIdLink2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/second" TargetMode="External"/>""");

        foreach (IFontResolver resolver in new IFontResolver[] { new LinkTestFontResolver(), new MissingTestFontResolver() })
        {
            (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input, resolver);

            PdfLinkAnnotation[] annotations = pages.Single().Annotations.ToArray();
            TestAssert.Equal(2, annotations.Length);
            TestAssert.Equal("https://example.invalid/first", annotations[0].Uri);
            TestAssert.Equal("https://example.invalid/second", annotations[1].Uri);
            TestAssert.True(annotations[1].X > annotations[0].X, "Adjacent link annotations should follow text order.");
        }
    }

    public static void PptxRunInternalSlideHyperlinkEmitsDestination()
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
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
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
                  <p:sldSz cx="9144000" cy="6858000"/>
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdRunJump" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide2.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="RunJump"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
                    </p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdRunJump"/></a:rPr><a:t>Jump</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/slide2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Target"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        foreach (IFontResolver resolver in new IFontResolver[] { new LinkTestFontResolver(), new MissingTestFontResolver() })
        {
            (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, resolver);

            TestAssert.Equal(2, pages.Count);
            PdfLinkAnnotation annotation = pages[0].Annotations.Single();
            TestAssert.True(annotation.Uri is null, "An internal run hyperlink should not carry a URI target.");
            TestAssert.True(annotation.Destination is { PageIndex: 1 }, "An internal run hyperlink should target the second page.");
            TestAssert.Equal(0, pages[1].Annotations.Count);
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable internal run hyperlink should not emit hyperlink diagnostics.");
        }
    }

    public static void PptxRunHyperlinkMissingRelationshipEmitsSingleDiagnostic()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdMissing"/></a:rPr><a:t>A </a:t></a:r>""" +
                """<a:r><a:rPr sz="1800" u="sng"><a:latin typeface="Link Sans"/></a:rPr><a:t>B</a:t></a:r>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdMissing"/></a:rPr><a:t>C</a:t></a:r></a:p>""",
            slideRelsXml: string.Empty);

        foreach (IFontResolver resolver in new IFontResolver[] { new LinkTestFontResolver(), new MissingTestFontResolver() })
        {
            (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, resolver);

            TestAssert.Equal(1, pages.Count);
            TestAssert.Equal(0, pages.Single().Annotations.Count);
            TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK"));
        }
    }

    public static void PptxTableCellRunHyperlinkEmitsUriAnnotation()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdCellLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/cell" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400"><a:tc>
                          <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdCellLink"/></a:rPr><a:t>Cell</a:t></a:r></a:p></a:txBody>
                          <a:tcPr/>
                        </a:tc></a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        foreach (IFontResolver resolver in new IFontResolver[] { new LinkTestFontResolver(), new MissingTestFontResolver() })
        {
            (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, resolver);

            PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
            TestAssert.Equal("https://example.invalid/cell", annotation.Uri);
            TestAssert.True(annotation.Width > 0d, "Table-cell hyperlink annotations should cover the linked text width.");
            TestAssert.True(Math.Abs((resolver is LinkTestFontResolver ? 19.8d : 16.65d) - annotation.Height) < 1e-6d, $"Table-cell hyperlink height should follow the emitted font ascent/descent box, got {annotation.Height}.");
            TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 72d + 144d, "Table-cell hyperlink X extent should stay inside the table frame.");
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable table-cell hyperlink should not emit hyperlink diagnostics.");
        }
    }

    public static void PptxRunHyperlinkActionWithoutTargetEmitsDiagnostic()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick action="ppaction://hlinkshowjump"/></a:rPr><a:t>Action</a:t></a:r></a:p>""",
            slideRelsXml: string.Empty);

        foreach (IFontResolver resolver in new IFontResolver[] { new LinkTestFontResolver(), new MissingTestFontResolver() })
        {
            (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, resolver);

            TestAssert.Equal(0, pages.Single().Annotations.Count);
            TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "An action-only run hyperlink should emit the unsupported-action diagnostic.");
        }
    }

    public static void PptxRotatedRunHyperlinkUsesTransformedBox()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdRunLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/rotated" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="RotatedLink"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm rot="5400000"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
                    </p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdRunLink"/></a:rPr><a:t>Link</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        foreach (IFontResolver resolver in new IFontResolver[] { new LinkTestFontResolver(), new MissingTestFontResolver() })
        {
            (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input, resolver);

            PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
            TestAssert.Equal("https://example.invalid/rotated", annotation.Uri);
            TestAssert.True(annotation.Height > annotation.Width, "A 90-degree rotated run link should cover a tall box, not the unrotated wide box.");
            TestAssert.True(annotation.X >= 0d && annotation.X + annotation.Width <= 720d, "Rotated run links should stay on the slide.");
            TestAssert.True(annotation.Y >= 0d && annotation.Y + annotation.Height <= 540d, "Rotated run links should stay on the slide.");
        }
    }
    private static string WriteRunLinkPackage(string paragraphXml, string slideRelsXml, string bodyProperties = "<a:bodyPr/>")
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  {slideRelsXml}
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="RunLink"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
                    </p:spPr>
                    <p:txBody>{bodyProperties}<a:lstStyle/>
                      {paragraphXml}
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
    }

    private sealed class LinkTestFontResolver : IFontResolver
    {
        private readonly byte[] fontBytes = TestFontBuilder.CreateTestFont();

        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(
                request.FamilyName,
                "LinkTestFont",
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                new MemoryFontProgramSource("link-test-font", fontBytes),
                IsFallback: false);
        }
    }

    private static (List<PdfPage> Pages, List<OoxPdfDiagnostic> Diagnostics) RenderSlidePages(string input, IFontResolver fontResolver)
    {
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        var diagnostics = new List<OoxPdfDiagnostic>();
        List<PdfPage> pages = new PptxRenderer(fontResolver).RenderPages(document, package, diagnostics.Add, CancellationToken.None).ToList();
        if (fontResolver is MissingTestFontResolver)
            TestAssert.True(pages[0].FallbackFonts.Count > 0 && pages[0].Fonts.Count == 0, "The missing-font control must use standard PDF fonts.");
        return (pages, diagnostics);
    }
}
