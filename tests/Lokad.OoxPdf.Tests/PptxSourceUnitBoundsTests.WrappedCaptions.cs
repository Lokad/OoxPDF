using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxSourceUnitBoundsTests
{
    public static void PptxOverflowCategoryCaptionsShareMeasuredWrappedLines()
    {
        foreach (string kind in new[] { "bar", "line", "area" })
        foreach (int angle in new[] { -60, -45, 45, 60 })
        foreach (bool high in new[] { false, true })
        {
            PdfPage Render(int count) => RenderWrappedCaption(new string('L', count), kind, angle, high);
            PdfPage two = Render(2), four = Render(4);
            var a = WrappedShows(two).Where(x => x.Text == "L").ToArray();
            var b = WrappedShows(four).Where(x => x.Text == "L").ToArray();
            TestAssert.True(a.Length == 2 && b.Length == 4 &&
                WrappedShows(four).Count(x => x.Text == "I") == 18,
                "Explicit overflow preserves every source line and neighboring caption.");
            var font = b[0].Font.Font;
            double advance = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender +
                font.Hhea.HorizontalLineGap) * 12d / font.UnitsPerEm;
            double sin = -Math.Sin(angle * Math.PI / 180d), cos = Math.Cos(angle * Math.PI / 180d);
            for (int index = 1; index < b.Length; index++)
                TestAssert.True(Math.Abs(b[index - 1].X - b[index].X + advance * sin) < .004 &&
                    Math.Abs(b[index - 1].Y - b[index].Y - advance * cos) < .004,
                    "Wrapped category baselines use the measured normal-line advance in both coordinates.");
            TestAssert.True(Math.Abs(b[0].X - a[0].X + advance * sin) < .004 &&
                Math.Abs(b[0].Y - a[0].Y - (high ? 2d * advance * cos : 0d)) < .004,
                "The full line stack centers across the tick while its nearest baseline stays on the selected side.");
            TestAssert.True(Math.Abs(CaptionPlotHeight(two) - CaptionPlotHeight(four)) < .003 &&
                Labels(four).SequenceEqual(new double[] { 0, 2, 4, 6 }),
                "Wrapped overflow preserves the authored plot and numeric interval.");
        }
    }

    public static void PptxOverflowCategoryNoWrapRetainsItsCompleteSourceLine()
    {
        const string source = "LongQuarterSummary";
        foreach (string kind in new[] { "bar", "line", "area" })
        foreach (int angle in new[] { -90, -45, 45, 90 })
        foreach (bool high in new[] { false, true })
        {
            PdfPage page = RenderWrappedCaption(source, kind, angle, high, wrap: "none");
            var shows = WrappedShows(page);
            TestAssert.True(shows.Length == 19 && shows.Count(x => x.Text == source) == 1 &&
                shows.Count(x => x.Text == "I") == 18,
                "Authored no-wrap preserves one complete source line under explicit overflow.");
        }
    }

    public static void PptxVerticalWrappedCaptionsKeepTheirFittedSourceReferenceEdge()
    {
        foreach (string kind in new[] { "bar", "line", "area" })
        foreach (int angle in new[] { -90, 90 })
        foreach (bool high in new[] { false, true })
        {
            PdfPage two = RenderWrappedCaption("LL", kind, angle, high);
            PdfPage four = RenderWrappedCaption("LLLL", kind, angle, high);
            var a = WrappedShows(two).Where(x => x.Text == "L").ToArray();
            var b = WrappedShows(four).Where(x => x.Text == "L").ToArray();
            TestAssert.True(a.Length == 2 && b.Length == 4 && WrappedShows(four).Count(x => x.Text == "I") == 18,
                "Vertical source-reference placement retains every source line and neighboring caption.");
            var font = b[0].Font.Font;
            double advance = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender +
                font.Hhea.HorizontalLineGap) * 12d / font.UnitsPerEm;
            for (int index = 1; index < b.Length; index++)
                TestAssert.True(Math.Abs(b[index].X - b[index - 1].X + Math.Sign(angle) * advance) < .004 &&
                    Math.Abs(b[index].Y - b[index - 1].Y) < .004,
                    "Vertical baselines retain their measured advance and common along-line alignment.");
            double sourceEdgeA = angle > 0 ? a[^1].X : a[0].X;
            double sourceEdgeB = angle > 0 ? b[^1].X : b[0].X;
            TestAssert.True(Math.Abs(sourceEdgeA - sourceEdgeB) < .004 && Math.Abs(a[0].Y - b[0].Y) < .004,
                "Adding wrapped lines preserves the fitted reference edge on either side and rotation direction.");
            TestAssert.True(Math.Abs(CaptionPlotHeight(two) - CaptionPlotHeight(four)) < .003 &&
                Labels(four).SequenceEqual(new double[] { 0, 2, 4, 6 }),
                "Vertical reference placement leaves the source plot and numeric interval unchanged.");
            PdfContentValidator.ValidatePage(four, 0, default);
        }
    }

    public static void PptxOverflowCategoryWrappingKeepsCompleteTextElements()
    {
        const string source = "e\u0301\U0001F600e\u0301\U0001F600";
        foreach (int angle in new[] { -45, 45 })
        foreach (bool high in new[] { false, true })
        {
            var shows = WrappedShows(RenderWrappedCaption(source, "bar", angle, high))
                .Where(x => x.Text != "I").ToArray();
            TestAssert.True(string.Concat(shows.Select(x => x.Text)) == source &&
                shows.All(x => x.Text is "e\u0301" or "\U0001F600"),
                "Shared chart wrapping keeps combining sequences and surrogate pairs on one line.");
        }
    }

    private static PdfPage RenderWrappedCaption(string text, string kind, int angle, bool high, string? wrap = null) =>
        RenderCaptionFrames(angle, 2, kind, highSide: high, manual: true, height: 432,
            categoryFontSize: 1200, valueFontSize: 800, categoryText: text, interiorCaptionIndex: 17,
            manualHeight: .8d, manualY: high ? .06d : .14d, labelSkip: 1, overflow: "overflow", wrap: wrap);

    private static (PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] WrappedShows(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
}
