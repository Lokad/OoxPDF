using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    internal static void PptxAutomaticCategoryReferencesKeepSourceEndpoints()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const string word = "LongQuarterSummary";
        foreach (int rotation in new[] { -45, 45 })
        foreach (bool reversed in new[] { false, true })
        foreach (bool high in new[] { false, true })
        {
            PdfPage Render(int skip, int? wordIndex, bool hiddenLast = false) => RenderCaption(row,
                18, high, 18, 100, "ordinary", rotation, "automatic", skip, text: "I", edit: xml =>
                {
                    xml.Descendants(c + "catAx").Single().Element(c + "scaling")!
                        .Element(c + "orientation")!.SetAttributeValue("val", reversed ? "maxMin" : "minMax");
                    foreach (var point in xml.Descendants(c + "cat").Descendants(c + "pt"))
                    {
                        int index = (int)point.Attribute("idx")!;
                        if (index == wordIndex || hiddenLast && index == 7)
                            point.Element(c + "v")!.Value = word;
                    }
                });
            PdfPage sparse = Render(2, 6), dense = Render(1, 6), hidden = Render(2, 6, hiddenLast: true);
            foreach (var (page, count) in new[] { (sparse, 4), (dense, 8), (hidden, 4) })
            {
                TestAssert.Contains("72 72 720 432 re W n", page.Content);
                var captions = Captions(page);
                TestAssert.Equal(1, captions.Count(text => text == word));
                TestAssert.Equal(count - 1, captions.Count(text => text == "I"));
            }
            var s = AdmissionBounds(sparse); var d = AdmissionBounds(dense);
            var h = AdmissionBounds(hidden);
            TestAssert.True(s == h && Captions(sparse).SequenceEqual(Captions(hidden)),
                "An omitted source word changes neither the measured edge fragment nor the source interval.");
            if (!reversed && rotation > 0)
                TestAssert.True(s.Right < d.Right - 1 && Math.Abs(s.Left - d.Left) < .003,
                    "Sparse painting extends the last prepared reference to the source end; numeric placement stays independent.");
            if (reversed)
            {
                var shortBounds = AdmissionBounds(Render(2, null));
                TestAssert.True(s.Left >= shortBounds.Left - .003 && s.Right <= shortBounds.Right + .003,
                    "A wider source reference cannot reclaim the reversed interval's endpoint bands.");
            }
        }
        var ordinary = AdmissionBounds(RenderCaption(row, 18, false, 18, 100, "ordinary", 45, "automatic", 2));
        var reversedShort = AdmissionBounds(RenderCaption(row, 18, false, 18, 100, "ordinary", 45,
            "automatic", 2, edit: xml => xml.Descendants(c + "catAx").Single().Element(c + "scaling")!
                .Element(c + "orientation")!.SetAttributeValue("val", "maxMin")));
        TestAssert.True(reversedShort.Left > ordinary.Left + 1 && reversedShort.Right < ordinary.Right - 1,
            "Independent Office short-caption controls retain the reversed reference pitch, not only painted ink.");

        static string[] Captions(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
            usePageCoordinates: true, fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal))
            .Select(show => show.Text).ToArray();
    }
}
