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
        {
            int stablePairs = 0;
            (string Text, double Height) Read(int height)
            {
                PdfPage page = RenderCaptionFrames(angle, 2, highSide: high,
                    height: height, categoryFontSize: size, overflow: "ellipsis");
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
                        $"Unchanged {size}/{angle}/{high} ellipsized captions retain their occupied height: {previous.Height} -> {next.Height}.");
                }
                previous = next;
            }
            TestAssert.True(stablePairs >= 2, "Expected independently stable fitted-caption resize pairs.");
        }
    }

    public static void PptxCategoryOverflowMarkerSeparatesOccupiedHeightFromPainting()
    {
        foreach (int size in new[] { 1200, 1600 })
        foreach (int angle in new[] { -45, 45, -90, 90 })
        foreach (bool high in new[] { false, true })
        {
            PdfPage generated = RenderCaptionFrames(angle, 2, highSide: high, height: 154,
                categoryFontSize: size, overflow: "ellipsis", categoryText: new string('W', 80));
            var first = DocxCommentsTests.ReadEmbeddedGlyphTextShows(generated, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(19, first.Length);
            string fitted = first[0].Text;
            TestAssert.True(fitted.EndsWith('\u2026') && first.All(show => show.Text == fitted),
                "Expected a common fitted caption with a generated overflow marker.");
            PdfPage authored = RenderCaptionFrames(angle, 2, highSide: high, height: 154,
                categoryFontSize: size, overflow: "ellipsis", categoryText: fitted);
            var second = DocxCommentsTests.ReadEmbeddedGlyphTextShows(authored, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(second.Length == first.Length && first.Zip(second).All(pair =>
                pair.First.Text == pair.Second.Text && Math.Abs(pair.First.X - pair.Second.X) < .003 &&
                Math.Abs(pair.First.Y - pair.Second.Y) < .003),
                "Returning a generated marker's advance to the plot preserves the painted text origins.");
            var font = first[0].Font.Font;
            double marker = font.GetAdvanceWidth(font.MapCodePoint('\u2026')) * (size / 100d) / font.UnitsPerEm;
            double expected = marker * Math.Sin(Math.Abs(angle) * Math.PI / 180d);
            TestAssert.True(Math.Abs(CaptionPlotHeight(generated) - CaptionPlotHeight(authored) - expected) < .003,
                "Generated markers hang outside occupied height; an authored ellipsis keeps its normal advance.");
        }
    }

    public static void PptxCategoryOverflowPreservesManualTextOrigins()
    {
        foreach (int angle in new[] { -45, 45 })
        foreach (bool high in new[] { false, true })
        {
            PdfPage generated = RenderCaptionFrames(angle, 2, highSide: high, manual: true,
                overflow: "ellipsis", categoryText: new string('W', 80));
            var first = DocxCommentsTests.ReadEmbeddedGlyphTextShows(generated, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(19, first.Length);
            TestAssert.True(first[0].Text.EndsWith('\u2026'), "Expected a fitted manual-frame caption.");
            PdfPage authored = RenderCaptionFrames(angle, 2, highSide: high, manual: true,
                overflow: "ellipsis", categoryText: first[0].Text);
            var second = DocxCommentsTests.ReadEmbeddedGlyphTextShows(authored, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(Math.Abs(CaptionPlotHeight(generated) - CaptionPlotHeight(authored)) < .003 &&
                first.Length == second.Length && first.Zip(second).All(pair => pair.First.Text == pair.Second.Text &&
                    Math.Abs(pair.First.X - pair.Second.X) < .003 && Math.Abs(pair.First.Y - pair.Second.Y) < .003),
                "Manual frames retain their source geometry and existing caption origins.");
        }
    }

    private static double CaptionPlotHeight(PdfPage page)
    {
        double[] axis = PptxColumnIntervalTests.Lines(page).Single(line =>
            line.Color == new RgbColor(255, 0, 0) && Math.Abs(line.V[0] - line.V[2]) < .001 &&
            Math.Abs(line.V[1] - line.V[3]) > 1).V;
        return Math.Abs(axis[3] - axis[1]);
    }
}
