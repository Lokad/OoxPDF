using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxSourceUnitBoundsTests
{
    public static void PptxFittedCategoryCaptionsReturnUnusedVerticalCapacity()
    {
        // Independent Office resize controls retain the same fitted captions
        // and give the entire added point to the plot. Challenge that contract
        // with portable typography, both directions and both label sides.
        foreach (int size in new[] { 1200, 1600 })
        foreach (int angle in new[] { -45, 45, -90, 90 })
        foreach (bool high in new[] { false, true })
        foreach (string overflow in new[] { "ellipsis", "clip" })
        {
            int stablePairs = 0;
            (string Text, double Height) Read(int height)
            {
                PdfPage page = RenderCaptionFrames(angle, 2, highSide: high,
                    height: height, categoryFontSize: size, overflow: overflow);
                string text = string.Join("|", DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    usePageCoordinates: false, fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal))
                    .Select(show => show.Text));
                double[] axis = PptxColumnIntervalTests.Lines(page).Single(line =>
                    line.Color == new RgbColor(255, 0, 0) && Math.Abs(line.V[0] - line.V[2]) < .001 &&
                    Math.Abs(line.V[1] - line.V[3]) > 1).V;
                return (text, Math.Abs(axis[3] - axis[1]));
            }
            var previous = Read(154);
            for (int height = 155; height <= 158; height++)
            {
                var next = Read(height);
                if (next.Text == previous.Text)
                {
                    stablePairs++;
                    TestAssert.True(Math.Abs(next.Height - previous.Height - 1) < .003,
                        $"Unchanged {size}/{angle}/{high}/{overflow} captions retain their occupied height: {previous.Height} -> {next.Height}.");
                }
                previous = next;
            }
            TestAssert.True(stablePairs >= 2, "Expected independently stable fitted-caption resize pairs.");
        }
    }
}
