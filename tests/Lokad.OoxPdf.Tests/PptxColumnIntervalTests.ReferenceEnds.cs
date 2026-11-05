using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    internal static void PptxSkippedCategoryPaintKeepsLogicalReferenceEnds()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (double angle in new[] { -45d, 45d, -30d, 30d })
        foreach (bool high in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string text in new[] { "I", "WWW" })
        {
            PdfPage Render(int? skip, string? mutation = null) => RenderCaption(row, 18d, high, 18d, 100,
                "ordinary", angle, skip: skip, text: text, edit: xml =>
                {
                    xml.Descendants(c + "dLbls").Remove();
                    xml.Descendants(c + "catAx").Single().Element(c + "scaling")!
                        .Element(c + "orientation")!.SetAttributeValue("val", reversed ? "maxMin" : "minMax");
                    foreach (var point in xml.Descendants(c + "cat").Descendants(c + "pt"))
                    {
                        int index = (int)point.Attribute("idx")!;
                        if (mutation == "last-long" && index == 7 || mutation == "middle-long" && index == 1)
                            point.Element(c + "v")!.Value = "LongQuarterSummary";
                        if (mutation == "first-empty" && index == 0 || mutation == "only-middle" && index != 2)
                            point.Element(c + "v")!.Value = string.Empty;
                    }
                });
            var complete = Render(1);
            foreach (int skip in new[] { 2, 3 })
            {
                var sparse = Render(skip);
                SameFrame(complete, sparse);
                TestAssert.Equal((8 + skip - 1) / skip, CategoryCount(sparse));
            }
            var baseline = Render(2);
            foreach (string mutation in new[] { "last-long", "middle-long", "first-empty", "only-middle" })
            {
                var modified = Render(2, mutation);
                SameFrame(baseline, modified);
                TestAssert.Equal(mutation == "only-middle" ? 1 : mutation == "first-empty" ? 3 : 4, CategoryCount(modified));
            }
        }

        static int CategoryCount(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
            usePageCoordinates: false, fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Count();
        static void SameFrame(PdfPage first, PdfPage second)
        {
            var a = AdmissionBounds(first); var b = AdmissionBounds(second);
            TestAssert.True(new[] { a.Left - b.Left, a.Top - b.Top, a.Right - b.Right, a.Bottom - b.Bottom }
                .All(delta => Math.Abs(delta) < .002d),
                "Independent Office controls retain logical reference ends when paint skips or hidden words change.");
        }
    }
}
