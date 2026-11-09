using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxRevisionBarTests
{
    public static void DocxRevisionBarHorizontalFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native Word controls use installed Windows fonts.");
        CheckHorizontalFrames(null);
    }

    public static void DocxRevisionBarParagraphMarkFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native Word controls use installed Windows fonts.");
        using var reference = ReadParagraphMarkReference();
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            TestAssert.Equal(2, row.GetProperty("OfficeView").GetProperty("RevisionsMode").GetInt32());
            var page = RenderParagraphMark(row, null);
            var bar = ReadParagraphMarkBar(page.Content, page.Height);
            var expected = row.GetProperty("Bounds").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            double[] actual = [bar.Left, bar.Top, bar.Right, bar.Bottom];
            TestAssert.True(actual.Zip(expected).All(pair => Math.Abs(pair.First - pair.Second) < 1d),
                "Measured paragraph-mark bars retain the independent one-point Word edge gate: " + row.GetProperty("Id").GetString());
            TestAssert.True(Math.Abs(bar.Left - expected[0]) < .03d &&
                Math.Abs(bar.Right - bar.Left - (expected[2] - expected[0])) < .03d,
                "The printed X/thickness contract remains unchanged.");
        }
    }

    internal static void CheckParagraphMarkAfterSpacing()
    {
        using var reference = ReadParagraphMarkReference();
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        var rows = reference.RootElement.GetProperty("Rows").EnumerateArray()
            .Where(row => row.GetProperty("AfterPoints").GetInt32() is 0 or 12)
            .OrderBy(row => row.GetProperty("AfterPoints").GetInt32()).ToArray();
        TestAssert.Equal(2, rows.Length);
        var pages = rows.Select(row => RenderParagraphMark(row, resolver)).ToArray();
        var bars = pages.Select(page => ReadParagraphMarkBar(page.Content, page.Height)).ToArray();
        double nativeDelta = rows[1].GetProperty("Bounds")[3].GetDouble() - rows[0].GetProperty("Bounds")[3].GetDouble();
        TestAssert.True(Math.Abs(bars[1].Top - bars[0].Top) < .001d &&
            Math.Abs((bars[1].Bottom - bars[0].Bottom) - nativeDelta) < .05d,
            "Authored after-spacing extends the changed paragraph's bar below its unchanged occupied top, as measured independently in Word.");
        var captions = pages.Select(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
            .Single(show => show.Text.Trim() == "Formatted:")).ToArray();
        TestAssert.True(Math.Abs(captions[0].Y - captions[1].Y) < .001d,
            "A trailing paragraph gap cannot move its caption.");
    }

    private static JsonDocument ReadParagraphMarkReference() => JsonDocument.Parse(File.ReadAllText(
        ParagraphMarkCase("paragraph-mark-intrinsic-bars-office.json")));

    private static string ParagraphMarkCase(string name) => Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "Cases", "RevisionBars", name));

    private static Lokad.OoxPdf.Pdf.PdfPage RenderParagraphMark(JsonElement row, IFontResolver? resolver)
    {
        string input = ParagraphMarkCase(row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        return new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup), null, default).ToBlockingEnumerable().Single();
    }

    private static (double Left, double Top, double Right, double Bottom) ReadParagraphMarkBar(string content, double pageHeight)
    {
        var bars = Regex.Matches(content, @"(?m)^(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f\r?$")
            .Cast<Match>().Select(match => new
            {
                X = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y = double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                Width = double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                Height = double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)
            }).Where(bar => bar.X < 60d && bar.Width < 1d && bar.Height > 5d).ToArray();
        TestAssert.Equal(1, bars.Length);
        var bar = bars[0];
        return (bar.X, pageHeight - bar.Y - bar.Height, bar.X + bar.Width, pageHeight - bar.Y);
    }

    internal static void CheckHorizontalFrames(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "RevisionBars"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "revision-bars-office-geometry.json")));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(folder, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            TestAssert.Equal(1, row.GetProperty("OfficeView").GetProperty("RevisionsMode").GetInt32());
            var pages = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup), null, default).ToBlockingEnumerable().ToArray();
            var page = pages[row.GetProperty("PageNumber").GetInt32() - 1];
            double x = row.GetProperty("X").GetDouble();
            double width = row.GetProperty("Width").GetDouble();
            var rectangles = Regex.Matches(page.Content,
                @"(?m)^(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<width>[\d.]+) (?<height>[\d.]+) re f\r?$");
            int matched = rectangles.Cast<Match>().Count(match =>
                double.Parse(match.Groups["height"].Value, CultureInfo.InvariantCulture) > 8d &&
                Math.Abs(double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture) - x) < .03d &&
                Math.Abs(double.Parse(match.Groups["width"].Value, CultureInfo.InvariantCulture) - width) < .03d);
            TestAssert.True(matched == 1, "Independent Word X/thickness across margins, orientation and fitted review frame: " + id);
        }
    }
}
