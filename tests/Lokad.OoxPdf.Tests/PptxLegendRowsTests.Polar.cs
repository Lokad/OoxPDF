using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxSideLegendClipsRetainTheActualGlyphInk()
    {
        foreach (string kind in new[] { "pie", "doughnut", "column", "line" })
        foreach (string position in new[] { "l", "r" })
        foreach (int size in new[] { 8, 12, 18 })
        foreach (int count in new[] { 2, 6 })
        {
            var page = Render(576, size, count, overlay: false, kind: kind,
                categoryLegend: kind != "line", legendPosition: position).Page;
            AssertSideLegendGlyphInk(page, count, $"{kind}/{position}/{size}/{count}");
        }
    }

    public static void PptxSideLegendClipsRetainInkBeyondNominalFontMetrics()
    {
        foreach (bool math in new[] { false, true })
        foreach (string kind in new[] { "pie", "doughnut" })
        foreach (string position in new[] { "l", "r" })
        foreach (int size in new[] { 12, 18 })
        foreach (int count in new[] { 2, 6 })
        {
            string[] labels = Enumerable.Range(1, count).Select(i => $"Category {i}").ToArray();
            labels[^1] = math ? "Flow ⇶" : "Flow ⨜";
            var page = Render(576, size, count, overlay: false, kind: kind,
                categoryLegend: true, categoryNames: labels, legendPosition: position,
                resolver: new PortableRenderingTests.PinnedResolver(mathFont: math)).Page;
            TestAssert.True(AssertSideLegendGlyphInk(page, count, $"{kind}/{position}/{size}/{count}/math:{math}") > 0,
                "These pinned font outlines extend beyond their nominal metric box.");
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal));
            TestAssert.True(shows.Select(show => show.Text).SequenceEqual(labels),
                "All source labels, including the tall symbol, survive in order.");
        }
    }

    private static int AssertSideLegendGlyphInk(PdfPage page, int count, string context)
    {
        int names = 0, measuredGlyphs = 0, beyondNominal = 0;
        foreach (var (block, clip, transform) in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(
            new PdfPage(page.Width, page.Height, block, page.Fonts), usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)))
        {
            names++;
            var font = show.Font.Font;
            bool typographic = (font.Os2.SelectionFlags & 0x80) != 0;
            double ascent = typographic ? font.Os2.TypographicAscender : font.Os2.WindowsAscender;
            double descent = typographic ? -(double)font.Os2.TypographicDescender : font.Os2.WindowsDescender;
            foreach (char character in show.Text)
            {
                if (char.IsWhiteSpace(character)) continue;
                ushort glyph = font.MapCodePoint(character);
                TestAssert.True(glyph != 0, "The pinned font contains the source character.");
                TestAssert.True(font.TryGetGlyphBounds(glyph, out var ink),
                    "The pinned portable font exposes actual source glyph ink.");
                if (ink.YMin < -descent || ink.YMax > ascent) beyondNominal++;
                double scale = show.Size / font.UnitsPerEm;
                double bottom = (show.Y + ink.YMin * scale) * transform.M22 + transform.M32;
                double top = (show.Y + ink.YMax * scale) * transform.M22 + transform.M32;
                measuredGlyphs++;
                TestAssert.True(bottom >= clip.Y - .01d && top <= clip.Y + clip.Height + .01d,
                    $"A complete side-legend name cannot lose its glyph ink: {context}/{show.Text}/{character}.");
            }
        }
        TestAssert.Equal(count, names);
        TestAssert.True(measuredGlyphs > count, "Every legend's source glyphs are checked independently of extraction.");
        return beyondNominal;
    }

    public static void PptxAuthoredSideLegendRetainsItsExplicitViewport()
    {
        const string layout = "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".5\"/><c:y val=\".2\"/><c:w val=\".2\"/><c:h val=\".1\"/></c:manualLayout></c:layout>";
        var page = Render(576, 18, 6, overlay: false, kind: "pie", categoryLegend: true,
            legendPosition: "r", legendMetadata: layout).Page;
        double[] clip = ReadPolarLegendClip(page, 576);
        TestAssert.True(Math.Abs(clip[0] - 360d) < .001d && Math.Abs(clip[1] - (page.Height - 108d - 108d)) < .001d &&
            Math.Abs(clip[2] - 115.2d) < .001d && Math.Abs(clip[3] - 36d) < .001d,
            "The explicit source rectangle retains its own viewport even when all names do not fit.");
    }

    public static void PptxPolarHorizontalRowsKeepEveryNameInsideTheirClip()
    {
        foreach (string kind in new[] { "pie", "doughnut" })
        foreach (string position in new[] { "t", "b" })
        foreach (bool overlay in new[] { false, true })
        foreach (int width in new[] { 420, 576 })
        foreach (int size in new[] { 12, 18 })
        foreach (int count in new[] { 6, 7 })
        {
            var result = Render(width, size, count, overlay, kind: kind,
                categoryLegend: true, legendPosition: position);
            var names = DocxCommentsTests.ReadEmbeddedGlyphTextShows(result.Page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            TestAssert.True(names.Select(show => show.Text).SequenceEqual(
                Enumerable.Range(1, count).Select(i => $"Category {i}")),
                $"Every {kind}/{position} legend name must survive in source order.");
            double[] clip = ReadPolarLegendClip(result.Page, width);
            foreach (var show in names)
            {
                var font = show.Font.Font;
                double advance = show.Text.Sum(c => font.GetAdvanceWidth(font.MapCodePoint(c))) * show.Size / font.UnitsPerEm;
                TestAssert.True(show.X >= clip[0] - .01d && show.X + advance <= clip[0] + clip[2] + .01d,
                    $"The complete {show.Text} must fit its actual {kind}/{position} legend clip at {width}pt/{size}pt.");
            }
            TestAssert.True(!result.Diagnostics.Any(item => item.Message.Contains("overflowing entries", StringComparison.Ordinal)),
                "These small legends fit complete rows without refusing any entries.");
            if (width == 420 && size == 12 && count == 6)
                TestAssert.True(names.GroupBy(show => show.Y).Select(row => row.Count()).SequenceEqual(new[] { 3, 3 }),
                    "The independent narrow Office controls wrap six ordinary names into two complete rows.");
        }
    }

    public static void PptxPolarLegendRowsStayAttachedWhilePieSpaceIsReserved()
    {
        foreach (string kind in new[] { "pie", "doughnut" })
        foreach (string position in new[] { "t", "b" })
        {
            var overlay = Render(420, 12, 6, true, kind: kind, categoryLegend: true, legendPosition: position);
            var reserved = Render(420, 12, 6, false, kind: kind, categoryLegend: true, legendPosition: position);
            var overlayNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(overlay.Page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            var reservedNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(reserved.Page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            TestAssert.True(overlayNames.Select(show => (show.Text, show.X, show.Y))
                .SequenceEqual(reservedNames.Select(show => (show.Text, show.X, show.Y))),
                "Overlay changes plot admission while both polar legends remain attached to their chart frame.");
            if (kind != "pie") continue;
            const string number = @"(-?\d+(?:\.\d+)?)";
            var centers = Regex.Matches(reserved.Page.Content,
                "(?m)^" + number + " " + number + @" l\r?\nh\r?\nf\*");
            TestAssert.Equal(6, centers.Count);
            double x = double.Parse(centers[0].Groups[1].Value, CultureInfo.InvariantCulture);
            double y = double.Parse(centers[0].Groups[2].Value, CultureInfo.InvariantCulture);
            Match start = Regex.Match(reserved.Page.Content, "(?m)^" + number + " " + number + @" m\r?$");
            TestAssert.True(start.Success, "The first sector starts on its circle.");
            double dx = double.Parse(start.Groups[1].Value, CultureInfo.InvariantCulture) - x;
            double dy = double.Parse(start.Groups[2].Value, CultureInfo.InvariantCulture) - y;
            double radius = Math.Sqrt(dx * dx + dy * dy);
            double[] clip = ReadPolarLegendClip(reserved.Page, 420);
            TestAssert.True(position == "b" ? y - radius >= clip[1] + clip[3] - .01d : y + radius <= clip[1] + .01d,
                "The automatic pie circle stays outside the measured legend footprint.");
        }
    }

    private static double[] ReadPolarLegendClip(PdfPage page, int width)
    {
        var clips = Regex.Matches(page.Content, @"(?m)^([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re W\*? n\r?$")
            .Select(match => match.Groups.Cast<Group>().Skip(1)
                .Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(rect => rect[2] <= width + .01d && rect[3] < 360d)
            .DistinctBy(rect => (rect[0], rect[1], rect[2], rect[3])).ToArray();
        TestAssert.Equal(1, clips.Length);
        return clips[0];
    }
}
