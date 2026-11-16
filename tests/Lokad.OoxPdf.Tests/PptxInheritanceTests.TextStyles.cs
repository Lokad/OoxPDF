using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
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
