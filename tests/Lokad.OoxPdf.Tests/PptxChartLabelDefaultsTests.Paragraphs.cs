using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxLineLabelParagraphsWrapCompleteComponentsWithNormalAdvance()
    {
        foreach (bool key in new[] { false, true })
        {
            var page = LineParagraphPage(LineParagraphSource(key));
            var shows = LineParagraphShows(page);
            TestAssert.True(shows.Length > 1, "The automatic chart-relative capacity wraps composed line labels.");
            TestAssert.Equal("Actual,W3,3", string.Concat(shows.Select(s => s.Text)).Replace(" ", "", StringComparison.Ordinal));
            var font = shows[0].Font.Font;
            double advance = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) *
                shows[0].Size / font.UnitsPerEm;
            for (int i = 1; i < shows.Length; i++)
                TestAssert.True(Math.Abs(shows[i - 1].Y - shows[i].Y - advance) < .004,
                    "Wrapped lines use the same normal advance as other measured chart paragraphs.");
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    public static void PptxSizedLineLabelParagraphsElideAndKeepAuthoredNoWrap()
    {
        var clipped = LineParagraphShows(LineParagraphPage(LineParagraphSource(false, .1d, .08d)));
        TestAssert.Equal(1, clipped.Length);
        TestAssert.True(clipped[0].Text.EndsWith('\u2026') && clipped[0].Text.Length < "Actual, W3, 3".Length,
            "An authored short body elides the visible line instead of painting the whole source outside it.");
        var noWrap = LineParagraphShows(LineParagraphPage(LineParagraphSource(false, .1d, .08d, wrap: "none")));
        TestAssert.Equal(1, noWrap.Length); TestAssert.Equal("Actual, W3, 3", noWrap[0].Text);
        var mandatory = LineParagraphShows(LineParagraphPage(LineParagraphSource(false, .3d, .5d,
            wrap: "none", name: "Actual\nPlan")));
        TestAssert.Equal("Actual\nPlan, W3, 3", string.Join("\n", mandatory.Select(s => s.Text)));
    }

    public static void PptxSizedLineLabelHeightCentersTextAndKeepsKeyAtTop()
    {
        var small = LineParagraphPage(LineParagraphSource(true, .15d, .2d));
        var tall = LineParagraphPage(LineParagraphSource(true, .15d, .4d));
        var a = LineParagraphShows(small); var b = LineParagraphShows(tall);
        TestAssert.True(a.Length > 1 && a.Length == b.Length, "Both independent native-height controls retain the complete wrapped paragraph.");
        for (int i = 0; i < a.Length; i++)
        {
            TestAssert.Equal(a[i].Text, b[i].Text);
            TestAssert.True(Math.Abs(a[i].X - b[i].X) < .002 && Math.Abs(a[i].Y - b[i].Y) < .002,
                "A larger right-side body keeps the text group centered around the same source point.");
        }
        var firstKey = MarkerRectangles(small, new RgbColor(255, 0, 0)).Last();
        var secondKey = MarkerRectangles(tall, new RgbColor(255, 0, 0)).Last();
        TestAssert.True(Math.Abs(secondKey.Y - firstKey.Y - 28.8d) < .002,
            "The key follows the top of the288pt chart-relative body, independently of centered text.");
    }

    public static void PptxLineLabelKeyCapacityFollowsRootFontWithoutChangingLabelFont()
    {
        foreach ((int root, double expected) in new[] { (6, 3d), (12, 6.96d), (14, 8.04d), (18, 9.96d), (24, 12d) })
        {
            var page = LineParagraphPage(LineParagraphSource(true, rootFont: root));
            double[] markers = MarkerRectangles(page, new RgbColor(255, 0, 0)).Select(m => m.Width).Order().ToArray();
            TestAssert.True(markers.Length == 2 && Math.Abs(markers[0] - expected) < .002 && Math.Abs(markers[1] - 12d) < .002,
                "Independent native root-font controls change key capacity while the authored point stays12pt.");
            TestAssert.True(LineParagraphShows(page).All(s => Math.Abs(s.Size - 20.04d) < .002),
                "Root typography controls key scale without replacing the label's own20pt font.");
        }
    }

    private static (PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] LineParagraphShows(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).ToArray();

    private static XDocument LineParagraphSource(bool key, double? width = null, double? height = null,
        string wrap = "square", string name = "Actual", int rootFont = 10)
    {
        var xml = KeyPlacementSource("r", 12, key, font: rootFont);
        XElement labels = xml.Descendants(C + "dLbls").Single();
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        labels.Add(new XElement(C + "txPr", new XElement(a + "bodyPr", new XAttribute("wrap", wrap)),
            new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", 2000))))));
        labels.AddFirst(new XElement(C + "dLbl", new XElement(C + "idx", new XAttribute("val", 2)),
            width.HasValue || height.HasValue ? SizedFrameLayout(width: width, height: height) : null,
            Components.Select(c => new XElement(C + c.Element, new XAttribute("val",
                c.Property is "ShowValue" or "ShowCategoryName" or "ShowSeriesName" || key && c.Property == "ShowLegendKey" ? 1 : 0)))));
        xml.Descendants(C + "ser").Single().Element(C + "tx")!.Element(C + "v")!.Value = name;
        return xml;
    }

    private static PdfPage LineParagraphPage(XDocument xml) => RenderChartPage(xml, new Dictionary<string, byte[]>
    {
        ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="1828800" y="914400"/><a:ext cx="6096000" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            """)
    });
}
