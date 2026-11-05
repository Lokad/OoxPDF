using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    // Independent Office controls vary source/physical height, width, caption
    // length, three faces, size, side and overflow. These portable relations
    // exercise the shared capacity without baking in one font's prefix length.
    internal static void PptxOuterCategoryFittingSharesDefaultEllipsisAndOccupiedFrames()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        const string source = "LongQuarterSummary";
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (int size in new[] { 18, 24 })
        foreach (int rotation in new[] { -45, 45 })
        foreach (bool high in new[] { false, true })
        {
            int previousLength = 0;
            bool fitted = false;
            foreach (double height in new[] { .45, .65, .85 })
            {
                PdfPage Render(string? overflow) => RenderCaption(row, size, high, 18, 100, "ordinary",
                    rotation, text: source, edit: xml =>
                    {
                        xml.Descendants(c + "manualLayout").Single().Element(c + "h")!.SetAttributeValue("val",
                            height.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        xml.Descendants(c + "catAx").Single().Element(c + "txPr")!.Element(a + "bodyPr")!
                            .SetAttributeValue("vertOverflow", overflow);
                    });
                PdfPage inherited = Render(null), ellipsis = Render("ellipsis");
                AssertSame(inherited, ellipsis);
                string text = CaptionText(inherited);
                string prefix = text.EndsWith('\u2026') ? text[..^1] : text;
                TestAssert.True(source.StartsWith(prefix, StringComparison.Ordinal) && prefix.Length >= previousLength,
                    "Increasing the source height preserves a monotone source prefix; a marker is only an overflow suffix.");
                previousLength = prefix.Length;
                fitted |= text.EndsWith('\u2026');

                PdfPage clip = Render("clip"), overflow = Render("overflow");
                string clipped = CaptionText(clip);
                TestAssert.True(CaptionText(overflow) == clipped && source.StartsWith(clipped, StringComparison.Ordinal) &&
                    !clipped.EndsWith('\u2026') && clipped.Length >= prefix.Length,
                    "Explicit clip and overflow preserve the same source prefix without a generated marker.");
                var cb = AdmissionBounds(clip); var ob = AdmissionBounds(overflow);
                TestAssert.True(ob.Bottom - ob.Top <= cb.Bottom - cb.Top + .003 &&
                    Math.Abs(high ? ob.Bottom - cb.Bottom : ob.Top - cb.Top) < .003,
                    "Explicit overflow retains capacity on the caption edge; clipping reclaims unused capacity.");
            }
            TestAssert.True(fitted, "The smallest source frame must exercise fitting, not only unchanged short captions.");
        }
        foreach (string target in new[] { "inner", "automatic" })
        {
            PdfPage Render(string? overflow) => RenderCaption(row, 18, false, 18, 100, "ordinary", 45,
                target, text: source, edit: xml => xml.Descendants(c + "catAx").Single()
                    .Element(c + "txPr")!.Element(a + "bodyPr")!.SetAttributeValue("vertOverflow", overflow));
            AssertSame(Render(null), Render("ellipsis"));
        }

        static string CaptionText(PdfPage page)
        {
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(8, shows.Length);
            TestAssert.True(shows.All(show => show.Text == shows[0].Text), "Fitting keeps every authored caption slot.");
            return shows[0].Text;
        }

        static void AssertSame(PdfPage first, PdfPage second)
        {
            var left = AdmissionBounds(first); var right = AdmissionBounds(second);
            TestAssert.True(left == right, "Default and explicit ellipsis reserve the same data rectangle.");
            var x = DocxCommentsTests.ReadEmbeddedGlyphTextShows(first, usePageCoordinates: true).ToArray();
            var y = DocxCommentsTests.ReadEmbeddedGlyphTextShows(second, usePageCoordinates: true).ToArray();
            TestAssert.True(x.Length == y.Length && x.Zip(y).All(pair => pair.First.Text == pair.Second.Text &&
                pair.First.X == pair.Second.X && pair.First.Y == pair.Second.Y && pair.First.Size == pair.Second.Size),
                "Equivalent overflow declarations preserve all caption/numeric content and origins.");
        }
    }
}
