using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
    public static void PptxHighlightsResolveEquivalentColorsAcrossOwners()
    {
        const string yellow = "<a:srgbClr val=\"FFFF00\"/>";
        PdfPage expected = RenderTextEffects("", Highlight(yellow));
        TestAssert.Contains("1 1 0 rg", expected.Content);
        foreach (string color in new[] {
            "<a:scrgbClr r=\"100000\" g=\"100000\" b=\"0\"/>",
            "<a:prstClr val=\"yellow\"/>",
            "<a:sysClr val=\"windowText\" lastClr=\"FFFF00\"/>",
            "<a:hslClr hue=\"3600000\" sat=\"100000\" lum=\"50000\"/>",
            "<a:schemeClr val=\"accent1\"/>" })
        {
            TestAssert.Equal(expected.Content, RenderTextEffects("", Highlight(color)).Content);
            foreach (bool paragraphDefault in new[] { false, true })
                TestAssert.Equal(expected.Content, RenderTextEffects(Highlight(color), "", paragraphDefault).Content);
        }
        // Half linear light is independently specified as sRGB 188, not 128.
        TestAssert.Equal(RenderTextEffects("", Highlight("<a:srgbClr val=\"BCBCBC\"/>")).Content,
            RenderTextEffects("", Highlight("<a:scrgbClr r=\"50000\" g=\"50000\" b=\"50000\"/>")).Content);
    }

    public static void PptxHighlightsRetainOrderedColorAndOpacityTransforms()
    {
        const string source = "<a:srgbClr val=\"808080\"><a:lumMod val=\"50000\"/><a:alpha val=\"60000\"/><a:alphaMod val=\"50000\"/><a:alphaOff val=\"20000\"/></a:srgbClr>";
        const string canonical = "<a:srgbClr val=\"404040\"><a:alpha val=\"50000\"/></a:srgbClr>";
        PdfPage expected = RenderTextEffects("", Highlight(canonical));
        TestAssert.True(expected.ExtGStates.Any(state => state.FillAlpha == .5d), "The highlight owns half opacity independently of opaque text.");
        TestAssert.Contains("0.251 g", expected.Content);
        TestAssert.Equal(expected.Content, RenderTextEffects("", Highlight(source)).Content);
        foreach (bool paragraphDefault in new[] { false, true })
            TestAssert.Equal(expected.Content, RenderTextEffects(Highlight(source), "", paragraphDefault).Content);
    }

    public static void PptxHighlightsKeepAdjacentOpacityBoundaries()
    {
        var parts = TextEffectParts("", "");
        Edit(parts, "ppt/slides/slide1.xml", xml => {
            foreach (var paragraph in xml.Descendants(A + "p"))
            {
                var run = paragraph.Element(A + "r")!;
                string text = run.Element(A + "t")!.Value;
                var second = new XElement(run);
                run.Element(A + "t")!.Value = text[..5];
                second.Element(A + "t")!.Value = text[5..];
                run.Element(A + "rPr")!.Add(XElement.Parse($"<a:highlight xmlns:a='{A}'><a:srgbClr val='FFFF00'><a:alpha val='0'/></a:srgbClr></a:highlight>"));
                second.Element(A + "rPr")!.Add(XElement.Parse($"<a:highlight xmlns:a='{A}'><a:srgbClr val='FFFF00'/></a:highlight>"));
                run.AddAfterSelf(second);
            }
        });
        WithDeck(parts, (document, package) => {
            PdfPage page = Render(document, package).Single();
            TestAssert.Equal(4, page.Content.AsSpan().Count("1 1 0 rg"));
            TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == 0d), "Transparent highlights cannot merge into an opaque neighbor.");
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
            TestAssert.Contains("Shape default", text);
            TestAssert.Contains("Cell default", text);
            PdfContentValidator.ValidatePage(page, 0, default);
        });
    }

    public static void PptxFallbackHighlightsRetainTheirOwnOpacity()
    {
        PdfPage? expected = null;
        foreach (string color in new[] { "<a:srgbClr val=\"FFFF00\"><a:alpha val=\"50000\"/></a:srgbClr>",
            "<a:prstClr val=\"yellow\"><a:alpha val=\"50000\"/></a:prstClr>" })
        {
            WithDeck(TextEffectParts("", Highlight(color)), (document, package) => {
                PdfPage page = new PptxRenderer(new MissingTestFontResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
                TestAssert.True(page.FallbackFonts.Count > 0 && page.Fonts.Count == 0, "This control must reach the standard-font painter.");
                TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == .5d), "Fallback highlight opacity must reach PDF resources.");
                if (expected is { } baseline) TestAssert.Equal(baseline.Content, page.Content);
                else expected = page;
                PdfContentValidator.ValidatePage(page, 0, default);
            });
        }
    }

    public static void PptxBulletsRetainExplicitOpacityAcrossKindsAndContainers()
    {
        foreach (bool numbered in new[] { false, true })
        foreach (int alpha in new[] { 0, 50000, 100000 })
        {
            var parts = TextEffectParts("", "");
            Edit(parts, "ppt/slides/slide1.xml", xml => {
                foreach (var paragraph in xml.Descendants(A + "p"))
                    paragraph.AddFirst(XElement.Parse($"<a:pPr xmlns:a='{A}' marL='457200' indent='-228600'><a:buClr><a:srgbClr val='000000'><a:alpha val='{alpha}'/></a:srgbClr></a:buClr>" +
                        (numbered ? "<a:buAutoNum type='arabicPeriod'/>" : "<a:buChar char='*'/>") + "</a:pPr>"));
            });
            WithDeck(parts, (document, package) => {
                PdfPage page = Render(document, package).Single();
                if (alpha != 100000)
                {
                    TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == alpha / 100000d), "Explicit bullet alpha must survive " + (numbered ? "numbering" : "character") + " placement.");
                    string stateName = "GS" + alpha + "F100000S";
                    TestAssert.Equal(2, page.Content.AsSpan().Count("/" + stateName + " gs"));
                }
                string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
                TestAssert.Contains("Shape default", text);
                TestAssert.Contains("Cell default", text);
                PdfContentValidator.ValidatePage(page, 0, default);
            });
        }
    }

    private static string Highlight(string color) => "<a:highlight>" + color + "</a:highlight>";
}
