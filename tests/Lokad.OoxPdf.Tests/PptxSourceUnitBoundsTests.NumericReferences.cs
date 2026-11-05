using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxSourceUnitBoundsTests
{
    public static void PptxNumericOnlyAdmissionRetainsFittedCaptionReferences()
    {
        const string word = "LongQuarterSummary";
        foreach (string kind in new[] { "bar", "line", "area" })
        foreach (int angle in new[] { -60, 45 })
        foreach (bool high in new[] { false, true })
        foreach (int size in new[] { 1200, 1800 })
        {
            PdfPage Render(string text, int numericSize) => RenderCaptionFrames(angle, 2, kind,
                highSide: high, manual: true, categoryFontSize: size, valueFontSize: numericSize,
                categoryText: text, height: 432, manualHeight: .88d,
                manualY: high ? .09d : .03d, labelSkip: 1);
            var source = Captions(Render(word, 800));
            TestAssert.True(source.Length == 19 && source.All(x => x.Text == source[0].Text) &&
                source[0].Text.EndsWith('\u2026'), "The original source must exercise caption fitting.");
            string prefix = source[0].Text[..^1];
            PdfPage complete = Render(word, 2400), reference = Render(prefix, 2400);
            var shows = Captions(complete);
            TestAssert.True(shows.Length == 19 && shows.All(x => x.Text == word),
                "Numeric admission retains complete captions within the admitted fitting capacity.");
            var font = shows[0].Font.Font;
            double Width(string text) => text.Sum(ch => (double)font.GetAdvanceWidth(font.MapCodePoint(ch))) *
                (size / 100d) / font.UnitsPerEm;
            double extra = (Width(word) - Width(prefix)) * Math.Sin(Math.Abs(angle) * Math.PI / 180d);
            double actual = CaptionPlotHeight(reference) - CaptionPlotHeight(complete);
            TestAssert.True(Math.Abs(actual - extra) < .004,
                $"Numeric-only admission retains the fitted reference before restoring its span: {kind}/{angle}/{high}/{size}, {actual} vs {extra}.");
            var a = Interval(reference); var b = Interval(complete);
            TestAssert.True(Math.Abs((high ? a.Low : a.High) - (high ? b.Low : b.High)) < .003,
                "Restoring a caption preserves the opposite numeric endpoint.");
        }

        static (PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] Captions(PdfPage page) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
    }

    public static void PptxNumericOnlySameEdgeAdmissionPreservesTheOppositeEndpoint()
    {
        foreach (string kind in new[] { "bar", "line", "area" })
        foreach (int angle in new[] { -60, 45 })
        {
            PdfPage Render(string text) => RenderCaptionFrames(angle, 2, kind,
                highSide: true, manual: true, categoryFontSize: 1200, valueFontSize: 2400,
                categoryText: text, height: 432, manualHeight: .88d, manualY: .03d, labelSkip: 1);
            var shortInterval = Interval(Render("L"));
            var longInterval = Interval(Render("LongQuarterSummary"));
            TestAssert.True(Math.Abs(shortInterval.Low - longInterval.Low) < .003 &&
                longInterval.High + 20d < shortInterval.High,
                "When captions share the overflowing numeric edge, their band consumes only that side.");
        }
    }

    private static (double Low, double High) Interval(PdfPage page)
    {
        double[] axis = PptxColumnIntervalTests.Lines(page).Single(line =>
            line.Color == new RgbColor(255, 0, 0) && Math.Abs(line.V[0] - line.V[2]) < .001 &&
            Math.Abs(line.V[1] - line.V[3]) > 1).V;
        return (Math.Min(axis[1], axis[3]), Math.Max(axis[1], axis[3]));
    }
}
