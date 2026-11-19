using System.Globalization;
using System.Numerics;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxPolarLabelsKeepTheirWedgesWhenEarlierLabelsAreDeleted()
    {
        string[] names = ["Fir", "Oak", "Elm", "Ash"];
        const string flags = "<c:showVal val=\"0\"/><c:showCatName val=\"1\"/><c:showSerName val=\"0\"/>" +
            "<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/><c:showLegendKey val=\"0\"/>";
        const string coordinate = @"(-?\d+(?:\.\d+)?)";
        var wedgePath = new Regex($@"(?m)^{coordinate} {coordinate} m\r?\n[^\r\n]+ c\r?\n(?:[^\r\n]+ [cl]\r?\n)*h\r?\nf\*");
        foreach (string kind in new[] { "pie", "doughnut" })
        {
            PdfPage RenderLabels(string pointOverride) => Render(576, 12, names.Length, overlay: true,
                kind: kind, categoryLegend: true, categoryNames: names, firstSliceAngle: 30,
                dataLabels: $"<c:dLbls>{pointOverride}{flags}</c:dLbls>").Page;

            PdfPage original = RenderLabels("");
            Match[] wedges = wedgePath.Matches(original.Content).ToArray();
            TestAssert.Equal(names.Length, wedges.Length);
            Vector2[] starts = wedges.Select(match => new Vector2(
                float.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture),
                float.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture))).ToArray();
            Vector2 center = (starts[0] + starts[2]) / 2;
            Vector2 firstRadius = Vector2.Normalize(starts[0] - center);
            // 30 degrees clockwise from the top in PDF coordinates.
            TestAssert.True(Math.Abs(firstRadius.X - .5d) < .001d && Math.Abs(firstRadius.Y - Math.Sqrt(3d) / 2) < .001d,
                $"{kind}: the painted first wedge follows the authored rotation.");

            var originalInk = new List<(string Text, PdfRectangle Ink)>();
            AssertPolarLabelGlyphInk(original, kind, originalInk);
            TestAssert.Equal(names.Length, originalInk.Count);
            for (int index = 0; index < names.Length; index++)
            {
                PdfRectangle ink = originalInk.Single(show => show.Text == names[index]).Ink;
                Vector2 inkCenter = new((float)(ink.X + ink.Width / 2), (float)(ink.Y + ink.Height / 2));
                Vector2 midpoint = (starts[index] + starts[(index + 1) % starts.Length]) / 2 - center;
                TestAssert.True(Vector2.Dot(Vector2.Normalize(inkCenter - center), Vector2.Normalize(midpoint)) > .99f,
                    $"{kind}: {names[index]} belongs near the midpoint of its own painted wedge.");
            }

            foreach (int deleted in new[] { 0, 1, 3 })
            {
                PdfPage edited = RenderLabels($"<c:dLbl><c:idx val=\"{deleted}\"/><c:delete val=\"1\"/></c:dLbl>");
                TestAssert.True(wedges.Select(match => match.Value).SequenceEqual(wedgePath.Matches(edited.Content).Select(match => match.Value)),
                    $"{kind}: deleting label {deleted} preserves all wedge paths.");
                var editedInk = new List<(string Text, PdfRectangle Ink)>();
                AssertPolarLabelGlyphInk(edited, $"{kind}/delete {deleted}", editedInk);
                TestAssert.True(originalInk.Where(show => show.Text != names[deleted]).SequenceEqual(editedInk),
                    $"{kind}: deleting label {deleted} preserves the other labels' full ink and positions.");
            }
        }
    }

    public static void PptxPolarDataLabelsRetainWholeAndRichTextInk()
    {
        const string layout = "<c:layout><c:manualLayout><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/>" +
            "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".1\"/><c:y val=\".1\"/>" +
            "<c:w val=\".6\"/><c:h val=\".7\"/></c:manualLayout></c:layout>";
        foreach (string kind in new[] { "pie", "doughnut" })
        foreach (string separator in new[] { ", ", " " })
        foreach (bool key in new[] { false, true })
        foreach (bool rich in new[] { false, true })
        foreach (bool math in new[] { false, true })
        {
            string[] names = Enumerable.Range(1, 6).Select(i => "North Eastern Area Revenue " + i).ToArray();
            string richLabel = rich ? "<c:dLbl><c:idx val=\"0\"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>" +
                "<a:r><a:t>North Eastern Area </a:t></a:r><a:r><a:t>Revenue 1" + separator + "4</a:t></a:r>" +
                "</a:p></c:rich></c:tx></c:dLbl>" : "";
            string labels = "<c:dLbls>" + richLabel + "<c:showVal val=\"1\"/><c:showCatName val=\"1\"/>" +
                "<c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/>" +
                $"<c:showLegendKey val=\"{(key ? 1 : 0)}\"/><c:separator>{separator}</c:separator></c:dLbls>";
            var page = Render(576, 12, 6, overlay: false, kind: kind, categoryLegend: true,
                categoryNames: names, plotLayout: layout, dataLabels: labels,
                seriesOptions: "<c:explosion val=\"20\"/>", resolver: new PortableRenderingTests.PinnedResolver(math)).Page;
            string sourceText = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: name => name.StartsWith("CP", StringComparison.Ordinal)).Select(show => show.Text));
            static string Compact(string text) => string.Concat(text.Where(c => !char.IsWhiteSpace(c)));
            foreach (string name in names) TestAssert.Contains(Compact(name + separator + "4"), Compact(sourceText));
            TestAssert.True(AssertPolarLabelGlyphInk(page, $"{kind}/{separator}/key:{key}/rich:{rich}/math:{math}") > 100,
                "The actual painted label glyphs are checked independently of text extraction.");
        }
    }

    private static int AssertPolarLabelGlyphInk(PdfPage page, string context, List<(string Text, PdfRectangle Ink)>? output = null)
    {
        int measured = 0;
        foreach (var (block, clip, transform) in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        {
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block, page.Fonts), false,
                fontResourceFilter: name => name.StartsWith("CP", StringComparison.Ordinal)).ToArray();
            if (shows.Length == 0) continue;
            TestAssert.Equal(1, shows.Length);
            var show = shows[0];
            var font = show.Font.Font;
            double cursor = show.X;
            PdfRectangle? inkBox = null;
            Match operation = Regex.Match(block, @"(?s)(?:\[.*?\] TJ|<[0-9A-F]+> Tj)");
            TestAssert.True(operation.Success, "The generated label uses a supported native glyph text operation.");
            Match spacing = Regex.Match(block, @"([-+\d.]+) Tc");
            double characterSpacing = double.Parse(spacing.Groups[1].Value, CultureInfo.InvariantCulture);
            foreach (Match token in Regex.Matches(operation.Value, @"<(?<hex>[0-9A-F]+)>|(?<adjust>[-+\d.]+)"))
            {
                if (token.Groups["adjust"].Success)
                {
                    cursor -= double.Parse(token.Groups["adjust"].Value, CultureInfo.InvariantCulture) * show.Size / 1000d;
                    continue;
                }
                string hex = token.Groups["hex"].Value;
                for (int offset = 0; offset < hex.Length; offset += 4)
                {
                    ushort cid = ushort.Parse(hex.AsSpan(offset, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    TestAssert.True(show.Font.UnicodeByCid.TryGetValue(cid, out int codePoint), "Every label CID has source ownership.");
                    ushort glyph = font.MapCodePoint(codePoint);
                    if (!char.IsWhiteSpace((char)codePoint))
                    {
                        TestAssert.True(font.TryGetGlyphBounds(glyph, out var ink), "The pinned face exposes label ink.");
                        double scale = show.Size / font.UnitsPerEm;
                        Vector2 first = Vector2.Transform(new((float)(cursor + ink.XMin * scale), (float)(show.Y + ink.YMin * scale)), transform);
                        Vector2 last = Vector2.Transform(new((float)(cursor + ink.XMax * scale), (float)(show.Y + ink.YMax * scale)), transform);
                        TestAssert.True(Math.Min(first.X, last.X) >= clip.X - .02d && Math.Max(first.X, last.X) <= clip.X + clip.Width + .02d &&
                            Math.Min(first.Y, last.Y) >= clip.Y - .02d && Math.Max(first.Y, last.Y) <= clip.Y + clip.Height + .02d,
                            $"Complete data-label glyph ink must survive its actual clip: {context}/{show.Text}/{(char)codePoint}.");
                        double x0 = Math.Min(first.X, last.X), y0 = Math.Min(first.Y, last.Y);
                        double x1 = Math.Max(first.X, last.X), y1 = Math.Max(first.Y, last.Y);
                        if (inkBox is { } prior)
                        {
                            x0 = Math.Min(x0, prior.X); y0 = Math.Min(y0, prior.Y);
                            x1 = Math.Max(x1, prior.X + prior.Width); y1 = Math.Max(y1, prior.Y + prior.Height);
                        }
                        inkBox = new PdfRectangle(x0, y0, x1 - x0, y1 - y0);
                        measured++;
                    }
                    cursor += Math.Round(font.GetAdvanceWidth(glyph) * 1000d / font.UnitsPerEm) * show.Size / 1000d + characterSpacing;
                }
            }
            if (inkBox is { } bounds) output?.Add((show.Text, bounds));
        }
        return measured;
    }
}
