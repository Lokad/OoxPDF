using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxUnitPrintOriginTests
{
    public static void DocxUnitPrintOriginsMatchIndependentWordFillAndSelectedBaseline()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent unit-scale origins use installed Windows fonts.");
        using var sources = ReadJson("sources.json");
        using var gold = ReadJson("office.json");
        foreach (var expected in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = sources.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == expected.GetProperty("Id").GetString());
            TestAssert.Equal(source.GetProperty("Sha256").GetString(), expected.GetProperty("InputSha256").GetString());
            var page = Render(source, null);
            var fill = ReadFill(page);
            double[] actual = [fill.X, page.Height - fill.Y - fill.Height, fill.X + fill.Width, page.Height - fill.Y];
            for (int index = 0; index < actual.Length; index++)
                TestAssert.True(Math.Abs(actual[index] - expected.GetProperty("Fill")[index].GetDouble()) < 1d,
                    "Unit-scale source/line geometry retains the independent Word fill edge: " + source.GetProperty("Id").GetString());
            var selected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Single(show => show.Text.Contains("Selected public words.", StringComparison.Ordinal));
            TestAssert.True(Math.Abs(page.Height - selected.Y - expected.GetProperty("SelectedOrigin")[1].GetDouble()) < 1d,
                "A measured unit-scale body must not inherit the fitted pre-layout shift.");
        }
    }

    public static void DocxUnitPrintOriginsPreservePortableContentAndCancellation()
    {
        using var sources = ReadJson("sources.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var source in sources.RootElement.EnumerateArray())
        {
            var page = Render(source, resolver);
            TestAssert.True(ReadFill(page).Height > 0d, "Unit-scale selected text keeps its source-owned fill.");
            string text = Normalize(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Select(show => show.Text)));
            foreach (var expected in source.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.True(text.Contains(Normalize(expected.GetString()!), StringComparison.Ordinal), "Unit-scale rendering preserves source prose.");
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Render(sources.RootElement[0], resolver, canceled.Token));
        static string Normalize(string text) => Regex.Replace(text, @"\s+", "");
    }

    public static void DocxUnitPrintOriginsAlreadyFittingPagesKeepPortableSizeAndPaperBounds()
    {
        using var sources = ReadJson("shrink-only-office.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var source in sources.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = Render(source, resolver);
            var lane = DocxCommentsTests.ReviewLaneBounds(page);
            TestAssert.True(Math.Abs(lane.Y) < .001d && Math.Abs(lane.Height - page.Height) < .001d,
                "An already fitting page is not magnified beyond its paper.");
            var body = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Single(show => show.Text.Trim() == source.GetProperty("BodyText").GetString());
            TestAssert.True(Math.Abs(body.Size - source.GetProperty("BodySize").GetDouble()) < .1d,
                "A wide markup margin does not magnify authored body type.");
        }
    }

    public static void DocxUnitPrintOriginsAlreadyFittingPagesMatchIndependentWordPaperAndBody()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent paper/body origins use installed Windows fonts.");
        using var sources = ReadJson("shrink-only-office.json");
        foreach (var source in sources.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = Render(source, null);
            var lane = DocxCommentsTests.ReviewLaneBounds(page);
            double[] edges = [lane.X, page.Height - lane.Y - lane.Height, lane.X + lane.Width, page.Height - lane.Y];
            for (int index = 0; index < edges.Length; index++)
                TestAssert.True(Math.Abs(edges[index] - source.GetProperty("Lane")[index].GetDouble()) < 1d,
                    "Printed lane edges retain independent Word geometry: " + source.GetProperty("Id").GetString());
            var body = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Single(show => show.Text.Trim() == source.GetProperty("BodyText").GetString());
            TestAssert.True(Math.Abs(body.X - source.GetProperty("BodyOrigin")[0].GetDouble()) < 1d &&
                Math.Abs(page.Height - body.Y - source.GetProperty("BodyOrigin")[1].GetDouble()) < 1d &&
                Math.Abs(body.Size - source.GetProperty("BodySize").GetDouble()) < .1d,
                "Already fitting pages retain independent body origins and size: " + source.GetProperty("Id").GetString());
        }
    }

    private static PdfPage Render(JsonElement source, IFontResolver? resolver, CancellationToken cancellationToken = default)
    {
        string file = Path.GetFileName(source.GetProperty("Input").GetString()!);
        using var stream = File.OpenRead(Path.Combine(CasesDirectory, file));
        TestAssert.Equal(source.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
        stream.Position = 0;
        using var package = OoxPackage.Open(stream, cancellationToken);
        var document = new DocxReader().Read(package, null, cancellationToken, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
        return new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(document, null, cancellationToken).ToBlockingEnumerable().Single();
    }

    private static PdfRectangle ReadFill(PdfPage page)
    {
        var match = Regex.Matches(page.Content, @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f")
            .Single(match => double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture) < 250d &&
                double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture) > 5d);
        double Value(string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
        return new(Value("x"), Value("y"), Value("w"), Value("h"));
    }

    private static string CasesDirectory => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "UnitPrintOrigins"));
    private static JsonDocument ReadJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.Combine(CasesDirectory, name)));
}
