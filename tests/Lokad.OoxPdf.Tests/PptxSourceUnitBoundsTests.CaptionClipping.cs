using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxSourceUnitBoundsTests
{
    public static void PptxAutomaticCaptionOverflowKeepsTheBaselineDataFrame()
    {
        foreach (int width in new[] { 288, 576 })
        foreach (int angle in new[] { -45, 45 })
        foreach (string kind in new[] { "bar", "line", "area" })
        {
            var full = RenderCaptionFrames(angle, 2, kind, highSide: true,
                categoryText: "Regional sales performance", width: width, onTickMarks: true, labelSkip: 1);
            var shortFirst = RenderCaptionFrames(angle, 2, kind, highSide: true,
                categoryText: "Regional sales performance", width: width, onTickMarks: true, labelSkip: 1, firstCaptionText: "Q");
            double[] Axis(PdfPage page) => PptxColumnIntervalTests.Lines(page)
                .Single(line => line.Color == new RgbColor(255, 0, 0) &&
                    Math.Abs(line.V[0] - line.V[2]) < .001 && Math.Abs(line.V[1] - line.V[3]) > 100).V;
            TestAssert.True(Axis(full).Zip(Axis(shortFirst)).All(pair => Math.Abs(pair.First - pair.Second) < .002),
                "Shortening one overflowing caption preserves the baseline numeric axis/data placement.");
            CheckChartCaptionClip(full, width);
            CheckChartCaptionClip(shortFirst, width);
        }
    }

    private static void CheckChartCaptionClip(PdfPage page, int width = 576)
    {
        var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal(19, captions.Length);
        int firstCaption = page.Content.IndexOf("/CCA", StringComparison.Ordinal);
        string frameClip = $"72 72 {width} 360 re W n";
        int clip = page.Content.IndexOf(frameClip, StringComparison.Ordinal);
        TestAssert.True(clip >= 0 && firstCaption > clip,
            "All source captions are emitted under their physical chart-frame clip, rather than moving the plot to fit their full ink.");
    }
}
