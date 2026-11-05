using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    internal static void PptxShortRotatedCaptionsKeepFiniteSpacingAndAuthoredSkips()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        foreach (double angle in new[] { -30d, 30d, -45d, 45d, -60d, -90d })
        foreach (bool high in new[] { false, true })
        {
            var page = RenderCaption(row, 18d, high, 18d, 100, "ordinary", angle, skip: null);
            var labels = Captions(page);
            TestAssert.True(labels.SequenceEqual(Enumerable.Repeat("I", 8)),
                "Independent Office controls retain every short caption when its finite rotated frame fits.");
            PptxSourceUnitBoundsTests.CheckCaptionInk(page, (int)angle, expectedCount: 8,
                frameRight: 792d, frameBottom: 72d, frameTop: 504d);
        }
        foreach (string text in new[] { "I", "WWW", "I I I" })
        foreach (int? skip in new int?[] { null, 1, 2, 3 })
        {
            var page = RenderCaption(row, 18d, false, 18d, 100, "ordinary", -45d, skip: skip, text: text);
            // Native controls separate a short word, a wide word and a phrase
            // of short words. The phrase needs its full advance, not only the
            // widest word. An explicit skip remains authoritative in all three.
            int effective = skip ?? (text == "I" ? 1 : 2);
            TestAssert.True(Captions(page).SequenceEqual(Enumerable.Repeat(text, (8 + effective - 1) / effective)),
                "Automatic density follows whole-caption geometry; authored density is unchanged.");
        }

        static string[] Captions(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
            usePageCoordinates: false, fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal))
            .Select(show => show.Text).ToArray();
    }
}
