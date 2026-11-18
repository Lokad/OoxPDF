using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
    public static void PptxInheritancePartialParagraphDefaultsKeepInheritedTextMetrics()
    {
        const string font = "<a:defRPr sz=\"2400\"><a:latin typeface=\"DejaVu Sans\"/></a:defRPr>";
        foreach (string paint in new[] { "", "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>" })
        {
            var parts = TextDefaultParts(font, "", "");
            (string Text, double Size, double X, double Y)[] baseline = [];
            WithDeck(parts, (document, package) => baseline = DocxCommentsTests.ReadEmbeddedGlyphTextShows(Render(document, package).Single())
                .Select(show => (show.Text, show.Size, show.X, show.Y)).ToArray());
            Edit(parts, "ppt/slides/slide1.xml", xml => {
                foreach (var paragraph in xml.Descendants(A + "p"))
                    paragraph.AddFirst(XElement.Parse($"<a:pPr xmlns:a='{A}'><a:defRPr>{paint}</a:defRPr></a:pPr>"));
            });
            WithDeck(parts, (document, package) => {
                PdfPage page = Render(document, package).Single();
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
                TestAssert.Contains("Shape default", string.Concat(shows.Select(show => show.Text)));
                TestAssert.Contains("Cell default", string.Concat(shows.Select(show => show.Text)));
                TestAssert.True(shows.All(show => show.Size == 24d), "Partial paragraph defaults retain the independently authored 24pt size in shapes and cells.");
                TestAssert.True(baseline.SequenceEqual(shows.Select(show => (show.Text, show.Size, show.X, show.Y))), "Empty/default-paint declarations preserve text size and placement.");
            });
        }
    }

    public static void PptxInheritanceParagraphPaintOutranksShapeFontReference()
    {
        const string blue = "<a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill>";
        const string red = "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>";
        const string hidden = "<a:noFill/>";
        string[] paints = [red, "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>", hidden,
            "<a:solidFill><a:srgbClr val=\"000000\"><a:alpha val=\"0\"/></a:srgbClr></a:solidFill>"];
        foreach (string paint in paints)
        {
            PdfPage control = RenderPaintOwners(blue, "", "", paint);
            foreach (string fontReference in new[] { "<a:srgbClr val=\"FFFFFF\"/>", "<a:srgbClr val=\"000000\"/>" })
            {
                PdfPage actual = RenderPaintOwners(blue, fontReference, paint, "");
                TestAssert.Equal(control.Content, actual.Content);
                TestAssert.True(control.ExtGStates.SequenceEqual(actual.ExtGStates), "Moving paint to paragraph defaults keeps its complete opacity.");
                PdfPage overridden = RenderPaintOwners(blue, fontReference, paint, red);
                TestAssert.Contains("1 0 0 rg", overridden.Content);
                TestAssert.Contains("OWNER_PROBE", DocxPageFootnoteNumberingTests.DecodeText(overridden));
            }
        }
        // Presentation defaults remain below the font reference. In particular,
        // its noFill must not suppress an explicitly visible shape default.
        foreach (string presentation in new[] { blue, hidden })
        foreach (string paragraph in new[] { "", "<a:latin typeface=\"DejaVu Sans\"/>", "<a:gradFill/>",
            "<a:solidFill><a:schemeClr val=\"unknown\"/></a:solidFill>" })
        {
            PdfPage page = RenderPaintOwners(presentation, "<a:srgbClr val=\"FF0000\"/>", paragraph, "");
            TestAssert.Contains("1 0 0 rg", page.Content);
            TestAssert.Contains("OWNER_PROBE", DocxPageFootnoteNumberingTests.DecodeText(page));
        }
    }

    public static void PptxInheritanceShapeFontReferenceKeepsOpacityAndOverrides()
    {
        const string red = "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>";
        PdfPage opaqueOverride = RenderPaintOwners("", "", "", red);
        foreach (var (alpha, expectedAlpha) in new[] { ("<a:alpha val=\"0\"/>", 0d),
            ("<a:alpha val=\"50000\"/>", .5d), ("<a:alpha val=\"100000\"/>", 1d),
            ("<a:alpha val=\"50000\"/><a:alphaMod val=\"40000\"/>", .2d) })
        {
            string reference = $"<a:srgbClr val=\"000000\">{alpha}</a:srgbClr>";
            PdfPage control = RenderPaintOwners("", "", "", "<a:solidFill>" + reference + "</a:solidFill>");
            PdfPage actual = RenderPaintOwners(red, reference, "", "");
            TestAssert.Equal(control.Content, actual.Content);
            TestAssert.True(control.ExtGStates.SequenceEqual(actual.ExtGStates), "Font-reference alpha uses the same paint as a direct fill.");
            if (expectedAlpha == 0d) TestAssert.DoesNotContain(" Tf", actual.Content);
            else if (expectedAlpha < 1d)
                TestAssert.True(actual.ExtGStates.Any(state => state.FillAlpha == expectedAlpha), "The independently specified opacity must reach PDF paint.");
            else TestAssert.Contains("OWNER_PROBE", DocxPageFootnoteNumberingTests.DecodeText(actual));
            foreach (bool paragraph in new[] { false, true })
            {
                PdfPage overridden = RenderPaintOwners("", reference, paragraph ? red : "", paragraph ? "" : red);
                TestAssert.Equal(opaqueOverride.Content, overridden.Content);
                TestAssert.True(opaqueOverride.ExtGStates.SequenceEqual(overridden.ExtGStates), "Higher-priority opaque paint replaces inherited alpha.");
            }
        }
    }

    private static PdfPage RenderPaintOwners(string presentation, string fontReference, string paragraph, string direct)
    {
        var parts = TextDefaultParts("<a:defRPr sz=\"2400\"><a:latin typeface=\"DejaVu Sans\"/>" + presentation + "</a:defRPr>", "", direct);
        Edit(parts, "ppt/slides/slide1.xml", xml => {
            xml.Descendants(P + "graphicFrame").Remove();
            XElement shape = xml.Descendants(P + "sp").Single();
            if (fontReference.Length != 0)
                shape.Element(P + "spPr")!.AddAfterSelf(XElement.Parse($"<p:style xmlns:p='{P}' xmlns:a='{A}'><a:fontRef idx='minor'>{fontReference}</a:fontRef></p:style>"));
            if (paragraph.Length != 0)
                shape.Descendants(A + "p").Single().AddFirst(XElement.Parse($"<a:pPr xmlns:a='{A}'><a:defRPr>{paragraph}</a:defRPr></a:pPr>"));
            shape.Descendants(A + "t").Single().Value = "OWNER_PROBE";
        });
        PdfPage? result = null;
        WithDeck(parts, (document, package) => {
            var diagnostics = new List<OoxPdfDiagnostic>();
            result = Render(document, package, diagnostics).Single();
            TestAssert.Equal(0, diagnostics.Count);
        });
        PdfPage page = result ?? throw new InvalidOperationException("The paint-owner control produced no page.");
        PdfContentValidator.ValidatePage(page, 0, default);
        return page;
    }

    public static void PptxInheritanceRunFillChoicesUseTheMoreSpecificLayer()
    {
        foreach (var (higher, lower) in new[] { ("solidFill", "noFill"), ("noFill", "solidFill"), ("gradFill", "solidFill") })
        {
            XElement high = new(A + "lvl1pPr", new XElement(A + "defRPr", new XAttribute("b", "0"), new XElement(A + higher)));
            XElement low = new(A + "lvl1pPr", new XElement(A + "defRPr", new XAttribute("b", "1"),
                new XElement(A + lower), new XElement(A + "latin", new XAttribute("typeface", "DejaVu Sans"))));
            string highBefore = high.ToString(), lowBefore = low.ToString();
            XElement merged = PptxParagraphPropertyMerger.MergeRendererDefaultProperties(A + "defRPr", high, low)!;
            XElement run = merged.Element(A + "defRPr")!;
            TestAssert.Equal(higher, run.Elements().First().Name.LocalName);
            TestAssert.Equal(2, run.Elements().Count()); // One effective fill and the inherited typeface.
            TestAssert.Equal("0", (string?)run.Attribute("b"));
            TestAssert.Equal("DejaVu Sans", (string?)run.Element(A + "latin")?.Attribute("typeface"));
            TestAssert.Equal(highBefore, high.ToString());
            TestAssert.Equal(lowBefore, low.ToString());
        }
        XElement empty = new(A + "lvl1pPr", new XElement(A + "defRPr"));
        XElement hidden = new(A + "lvl1pPr", new XElement(A + "defRPr", new XElement(A + "noFill")));
        TestAssert.NotNull(PptxParagraphPropertyMerger.MergeRendererDefaultProperties(A + "defRPr", empty, hidden)!.Element(A + "defRPr")!.Element(A + "noFill"));
    }

    public static void PptxInheritanceSpacingAndTabsSurviveEmptyAndRedundantLayers()
    {
        XElement lower = XElement.Parse($"""
            <a:lvl1pPr xmlns:a="{A}"><a:lnSpc><a:spcPts val="4000"/></a:lnSpc><a:spcBef><a:spcPts val="700"/></a:spcBef><a:spcAft><a:spcPts val="1800"/></a:spcAft><a:buChar char="*"/><a:tabLst><a:tab pos="1016000"/></a:tabLst><a:defRPr sz="2400"/></a:lvl1pPr>
            """);
        foreach (bool repeatFont in new[] { false, true })
        {
            XElement higher = new(A + "lvl1pPr", repeatFont ? new XElement(A + "defRPr", new XAttribute("sz", "2400")) : null,
                new XElement(A + "extLst"));
            XElement merged = PptxParagraphPropertyMerger.MergeRendererDefaultProperties(A + "defRPr", higher, lower)!;
            foreach (string name in new[] { "lnSpc", "spcBef", "spcAft", "tabLst" })
                TestAssert.True(XNode.DeepEquals(lower.Element(A + name), merged.Element(A + name)), "An absent supported paragraph property inherits as a whole: " + name);
            TestAssert.Equal("lnSpc|spcBef|spcAft|buChar|tabLst|defRPr|extLst", string.Join("|", merged.Elements().Select(e => e.Name.LocalName)));
        }
        XElement overrides = XElement.Parse($"""
            <a:lvl1pPr xmlns:a="{A}"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:tabLst/></a:lvl1pPr>
            """);
        XElement result = PptxParagraphPropertyMerger.MergeRendererDefaultProperties(A + "defRPr", overrides, lower)!;
        TestAssert.Equal("125000", (string?)result.Element(A + "lnSpc")?.Element(A + "spcPct")?.Attribute("val"));
        TestAssert.Equal(1, result.Element(A + "lnSpc")!.Elements().Count());
        TestAssert.Equal(0, result.Element(A + "tabLst")!.Elements().Count());
        TestAssert.True(XNode.DeepEquals(lower.Element(A + "spcAft"), result.Element(A + "spcAft")), "An explicit line-spacing override does not erase independent after-spacing.");
    }

    public static void PptxInheritancePresentationDefaultsReachShapesAndCells()
    {
        const string blue = "<a:solidFill><a:srgbClr val=\"0066CC\"/></a:solidFill>";
        const string red = "<a:solidFill><a:srgbClr val=\"CC0000\"/></a:solidFill>";
        foreach (bool directOverride in new[] { false, true })
        {
            string defaults = "<a:defRPr sz=\"2400\">" + blue + "<a:latin typeface=\"DejaVu Sans\"/></a:defRPr>";
            WithDeck(TextDefaultParts(defaults, directOverride ? "sz=\"3000\"" : "", directOverride ? red : ""), (document, package) =>
            {
                TestAssert.NotNull(document.DefaultTextStyle);
                TestAssert.Equal("presentation", document.DefaultTextStyle!.Parent!.Name.LocalName);
                var scene = new PptxSceneBuilder().Build(document, package, default);
                var inherited = scene.Slides[0].SlideNodes.Single(node => node.Kind == PptxSceneNodeKind.Shape)
                    .TextBody!.Paragraphs.Single().CascadeLayers.Single(layer => layer.Name == "defaultTextStyle").Source;
                TestAssert.Equal("2400", (string?)inherited?.Element(A + "defRPr")?.Attribute("sz"));
                var fonts = new PresentationFontResolver(new PortableRenderingTests.PinnedResolver(), default);
                var glyphs = PptxRenderer.InspectTextGlyphRuns(document, package, 0, scene, fonts);
                string text = string.Concat(glyphs.Select(run => run.Text));
                TestAssert.Contains("Shape default", text);
                TestAssert.Contains("Cell default", text);
                TestAssert.True(glyphs.All(run => run.LayoutFontSize == (directOverride ? 30d : 24d)), "Both containers use the actual presentation default below direct run overrides.");
                string content = Render(document, package)[0].Content;
                TestAssert.Contains(directOverride ? "0.8 0 0 rg" : "0 0.4 0.8 rg", content);
            });
        }
    }

    private static Dictionary<string, byte[]> TextDefaultParts(string defaults, string runAttributes, string runPaint)
    {
        string Body(string text) => "<a:bodyPr lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\"/><a:lstStyle/><a:p><a:r><a:rPr " + runAttributes + ">" + runPaint + "</a:rPr><a:t>" + text + "</a:t></a:r></a:p>";
        var presentation = XDocument.Parse(PptxTests.BasicPresentation());
        presentation.Root!.Add(XElement.Parse($"<p:defaultTextStyle xmlns:p='{P}' xmlns:a='{A}'><a:lvl1pPr>{defaults}</a:lvl1pPr></p:defaultTextStyle>"));
        string slide = $$"""
            <p:sld xmlns:p="{{P}}" xmlns:a="{{A}}"><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="508000" y="508000"/><a:ext cx="3048000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody>{{Body("Shape default")}}</p:txBody></p:sp>
              <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="508000" y="1778000"/><a:ext cx="3048000" cy="1016000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="3048000"/></a:tblGrid><a:tr h="1016000"><a:tc><a:txBody>{{Body("Cell default")}}</a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
            </p:spTree></p:cSld></p:sld>
            """;
        return new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(presentation.ToString()),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide)
        };
    }
}
