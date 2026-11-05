using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxSourceUnitBoundsTests
{
    public static void PptxInnerCategoryEllipsisRetainsACompleteSourceCharacter()
    {
        // Independent Office controls retain a source character when a tight
        // authored inner frame cannot fit even the ellipsis. Exercise that
        // content contract with portable fonts and a supplementary scalar.
        foreach (string kind in new[] { "bar", "line", "area" })
        foreach (int angle in new[] { -45, 45, -90, 90 })
        foreach (bool high in new[] { false, true })
        foreach (string first in new[] { "L", "W", "\U0001F600" })
        foreach (string? overflow in new string?[] { null, "ellipsis" })
        {
            PdfPage page = RenderCaptionFrames(angle, 2, kind, highSide: high, manual: true,
                categoryFontSize: 1800, overflow: overflow, categoryText: first + "QuarterSummary", manualHeight: .9d);
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(19, captions.Length);
            TestAssert.True(captions.All(show => show.Text == first + "\u2026"),
                $"Tight {kind}/{angle}/{high}/{overflow} inner captions retain a whole source character: {string.Join('|', captions.Select(show => show.Text))}.");
        }
    }

    public static void PptxInnerSingleCharacterCaptionNeedsNoGeneratedEllipsis()
    {
        foreach (int angle in new[] { -45, 45 })
        foreach (bool high in new[] { false, true })
        foreach (string text in new[] { "W", "\U0001F600" })
        {
            PdfPage page = RenderCaptionFrames(angle, 2, highSide: high, manual: true,
                categoryFontSize: 1800, overflow: "ellipsis", categoryText: text, manualHeight: .9d);
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(19, captions.Length);
            TestAssert.True(captions.All(show => show.Text == text),
                "A complete single-scalar caption has no omitted tail to mark.");
        }
    }

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
            var font = first[0].Font.Font;
            double marker = font.GetAdvanceWidth(font.MapCodePoint('\u2026')) * (size / 100d) / font.UnitsPerEm;
            double expected = marker * Math.Sin(Math.Abs(angle) * Math.PI / 180d);
            bool trailingEndAtAxis = (angle < 0) != high;
            double expectedOriginShift = trailingEndAtAxis ? 0d : high ? expected : -expected;
            TestAssert.Equal(first.Length, second.Length);
            var generatedPlot = Interval(generated); var authoredPlot = Interval(authored);
            for (int index = 0; index < first.Length; index++)
            {
                double ratio = (index + .5d) / first.Length;
                double sourceShift = generatedPlot.Left - authoredPlot.Left +
                    ratio * (generatedPlot.Width - authoredPlot.Width);
                double markerShift = trailingEndAtAxis ? marker * Math.Cos(Math.Abs(angle) * Math.PI / 180d) : 0d;
                // The X identity combines separately rounded PDF axis/text
                // coordinates and rotation; allow their sub-.005pt error.
                TestAssert.True(first[index].Text == second[index].Text &&
                    Math.Abs(first[index].X - second[index].X - sourceShift - markerShift) < .005 &&
                    Math.Abs(first[index].Y - second[index].Y - expectedOriginShift) < .003,
                    $"Generated marker origins follow the admitted source interval and rotated prefix anchor: {size}/{angle}/{high}/{index}, X={first[index].X - second[index].X}, source={sourceShift}, marker={markerShift}, Y={first[index].Y - second[index].Y}, expectedY={expectedOriginShift}.");
            }
            TestAssert.True(Math.Abs(CaptionPlotHeight(generated) - CaptionPlotHeight(authored) - expected) < .003,
                "Generated markers hang outside occupied height; an authored ellipsis keeps its normal advance.");
        }

        static (double Left, double Width) Interval(PdfPage page)
        {
            var lines = PptxColumnIntervalTests.Lines(page);
            double left = lines.Single(line => line.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(line.V[0] - line.V[2]) < .001 && Math.Abs(line.V[1] - line.V[3]) > 1).V[0];
            double width = lines.Where(line => line.Color == new RgbColor(0, 0, 255) &&
                Math.Abs(line.V[1] - line.V[3]) < .001).Max(line => Math.Abs(line.V[2] - line.V[0]));
            return (left, width);
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

    public static void PptxCategoryOverflowTracksValueFontWithoutChangingCategoryTypeface()
    {
        foreach (int angle in new[] { -45, 45 })
        foreach (bool high in new[] { false, true })
        {
            var prefixes = new List<string>();
            foreach (int valueSize in new[] { 800, 1200, 1800 })
            {
                PdfPage page = RenderCaptionFrames(angle, 2, highSide: high, height: 154,
                    overflow: "ellipsis", categoryText: new string('W', 80), valueFontSize: valueSize);
                var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
                    fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
                TestAssert.Equal(19, captions.Length);
                TestAssert.True(captions.All(show => show.Size == 12 && show.Text == captions[0].Text),
                    "Changing numeric typography keeps the category typeface size and source positions.");
                TestAssert.True(Labels(page).SequenceEqual(new double[] { 0, 2, 4, 6 }),
                    "Numeric typography preserves the authored source scale.");
                prefixes.Add(captions[0].Text);
            }
            TestAssert.True(prefixes[0].Length > prefixes[^1].Length && prefixes.Zip(prefixes.Skip(1)).All(
                pair => pair.First.Length >= pair.Second.Length),
                "Larger numeric type consumes fitting space at both ends; category typography remains fixed.");
        }
    }

    public static void PptxBuiltInCategoryFittingIgnoresOverriddenRoleFonts()
    {
        foreach (int gallery in new[] { 2, 9, 48 })
        foreach (int size in new[] { 1200, 1600 })
        foreach (int angle in new[] { -45, 45 })
        foreach (bool high in new[] { false, true })
        {
            PdfPage Render(int roleSize, int? globalSize = null) => RenderCaptionFrames(angle, 2,
                highSide: high, height: 154, categoryFontSize: size, overflow: "ellipsis",
                categoryText: new string('i', 80), builtInStyle: gallery,
                categoryRoleFontSize: roleSize, chartFontSize: globalSize);
            PdfPage baseline = Render(1200);
            var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(baseline, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(19, expected.Length);
            TestAssert.True(expected.All(show => show.Text.EndsWith('\u2026')),
                "The fixture must exercise fitted captions, not unchanged full strings.");
            foreach (PdfPage variant in new[] { Render(800), Render(1800), Render(1200, 800), Render(1200, 1800) })
            {
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(variant, usePageCoordinates: true,
                    fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
                TestAssert.True(Math.Abs(CaptionPlotHeight(baseline) - CaptionPlotHeight(variant)) < .003 &&
                    actual.Length == expected.Length && expected.Zip(actual).All(pair =>
                        pair.First.Text == pair.Second.Text && pair.First.Size == pair.Second.Size &&
                        Math.Abs(pair.First.X - pair.Second.X) < .003 && Math.Abs(pair.First.Y - pair.Second.Y) < .003),
                    "Overridden role/global fonts preserve built-in fitting, effective caption fonts and occupied geometry.");
                TestAssert.True(Labels(variant).SequenceEqual(Labels(baseline)),
                    "Unused font defaults preserve the authored numeric scale.");
            }
        }
    }

    public static void PptxBuiltInAxesInheritChartFontSizeBeforeRoleDefaults()
    {
        foreach (int gallery in new[] { 2, 9, 48 })
        foreach (int? globalSize in new int?[] { null, 1800 })
        foreach (int angle in new[] { -45, 45 })
        {
            int size = globalSize ?? 1000;
            PdfPage Render(bool inherit) => RenderCaptionFrames(angle, 2, height: 240,
                categoryFontSize: size, valueFontSize: size, overflow: "ellipsis", categoryText: new string('W', 80),
                builtInStyle: gallery, categoryRoleFontSize: 1600, valueRoleFontSize: 800,
                chartFontSize: globalSize, inheritAxisFontSizes: inherit);
            PdfPage explicitFonts = Render(false), inheritedFonts = Render(true);
            foreach (string prefix in new[] { "CCA", "CVA" })
            {
                var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(explicitFonts, usePageCoordinates: true,
                    fontResourceFilter: name => name.StartsWith(prefix, StringComparison.Ordinal)).ToArray();
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(inheritedFonts, usePageCoordinates: true,
                    fontResourceFilter: name => name.StartsWith(prefix, StringComparison.Ordinal)).ToArray();
                TestAssert.True(expected.Length > 0 && expected.Length == actual.Length && expected.Zip(actual).All(pair =>
                    pair.First.Text == pair.Second.Text && pair.First.Size == pair.Second.Size &&
                    Math.Abs(pair.First.X - pair.Second.X) < .003 && Math.Abs(pair.First.Y - pair.Second.Y) < .003),
                    "Equivalent local/global axis sizes preserve typography and geometry; unused role sizes do not become effective.");
            }
            TestAssert.True(Math.Abs(CaptionPlotHeight(explicitFonts) - CaptionPlotHeight(inheritedFonts)) < .003,
                "Equivalent font declarations retain the occupied caption frame.");
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
