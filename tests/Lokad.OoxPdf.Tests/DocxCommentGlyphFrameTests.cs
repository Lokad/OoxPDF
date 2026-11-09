using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxCommentGlyphFrameTests
{
    public static void DocxCommentGlyphFramesMatchIndependentWordPaths()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent comment glyph frames require installed Windows fonts.");
        using var sources = ReadJson("sources.json");
        using var gold = ReadJson("office.json");
        foreach (var source in sources.RootElement.EnumerateArray())
        {
            var expected = gold.RootElement.GetProperty("Rows").EnumerateArray()
                .Single(row => row.GetProperty("Id").GetString() == source.GetProperty("Id").GetString());
            TestAssert.Equal(source.GetProperty("Sha256").GetString(), expected.GetProperty("InputSha256").GetString());
            var page = Render(Read(source, default), null);
            var actual = ReadGeometry(page);
            TestAssert.True(Math.Abs(actual.StrokeWidth - expected.GetProperty("StrokeWidth").GetDouble()) < .001d,
                "Comment brackets retain the physical Office stroke across print scales.");
            for (int path = 0; path < actual.Paths.Length; path++)
                for (int axis = 0; axis < actual.Paths[path].Length; axis++)
                    TestAssert.True(Math.Abs(actual.Paths[path][axis] - expected.GetProperty("Paths")[path][axis].GetDouble()) < 1d,
                        "All bracket corners and stems retain the absolute Word threshold: " + source.GetProperty("Id").GetString() + "/" + path + "/" + axis);
            AssertContent(page, source);
        }
    }

    public static void DocxCommentGlyphFramesKeepBookmarksRunSplitsAndCancellationNeutral()
    {
        using var sources = ReadJson("sources.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var source in sources.RootElement.EnumerateArray().Where(row => row.GetProperty("PortableProbe").GetBoolean()))
        {
            var original = ReadGeometry(Render(Read(source, default), resolver));
            TestAssert.True(Math.Abs(original.StrokeWidth - .14d) < .001d, "Portable prose brackets retain a physical stroke.");
            foreach ((bool bookmark, bool split) in new[] { (true, false), (false, true), (true, true) })
            {
                var document = DocxInlineCompositionTests.ReadMutation(Fixture(source), OoxPdfDocxMarkupMode.AllMarkup, bookmark, split);
                var page = Render(document, resolver);
                var transformed = ReadGeometry(page);
                AssertPathsEqual(original.Paths, transformed.Paths);
                TestAssert.Equal(original.StrokeWidth, transformed.StrokeWidth);
                AssertContent(page, source);
            }
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Read(sources.RootElement[0], canceled.Token));
        var prepared = Read(sources.RootElement[0], default);
        TestAssert.Throws<OperationCanceledException>(() => Render(prepared, resolver, canceled.Token));
    }

    public static void DocxCommentGlyphFramesSeparateNaturalBracketsFromAuthoredSpacing()
    {
        using var sources = ReadJson("sources.json");
        var source = sources.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "arial12-auto");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        double[][]? originalShape = null;
        double originalFill = 0d;
        foreach ((string rule, int twips) in new[] { ("auto", 240), ("atLeast", 1200), ("exact", 640), ("exact", 1200) })
        {
            var document = DocxInlineCompositionTests.ReadMutation(Fixture(source), OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, _) =>
                {
                    var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(w + "commentRangeStart").Any());
                    var spacing = paragraph.Element(w + "pPr")!.Element(w + "spacing")!;
                    spacing.SetAttributeValue(w + "lineRule", rule);
                    spacing.SetAttributeValue(w + "line", twips);
                });
            var actual = ReadGeometry(Render(document, resolver));
            double top = actual.Paths[1][3];
            double[][] shape = actual.Paths.Select(path => new[] { path[0], path[1] - top, path[2], path[3] - top }).ToArray();
            if (originalShape is null) { originalShape = shape; originalFill = actual.FillHeight; }
            else AssertPathsEqual(originalShape, shape);
            if (rule == "atLeast")
                TestAssert.True(Math.Abs(actual.FillHeight - originalFill) < .002d, "Extra minimum leading does not enlarge the natural selection strut.");
            if (rule == "exact")
            {
                double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                TestAssert.True(Math.Abs(actual.FillHeight - twips / 20d * scale) < .002d,
                    "Exact spacing still owns the selection height while glyph brackets retain their natural shape.");
            }
        }
    }

    private static (double[][] Paths, double StrokeWidth, double FillHeight) ReadGeometry(PdfPage page)
    {
        var fills = Regex.Matches(page.Content, @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f");
        TestAssert.True(fills.Count > 0, "Selected prose retains its filled source rectangle.");
        var fill = fills.OrderBy(match => Number(match, "x")).First();
        string afterFill = page.Content[(fill.Index + fill.Length)..];
        var paths = Regex.Matches(afterFill, @"(?m)^(?<x>[-\d.]+) (?<y>[-\d.]+) m\s+(?<endX>[-\d.]+) (?<endY>[-\d.]+) l S")
            .Take(6).Select(match => new[] { Number(match, "x"), page.Height - Number(match, "y"), Number(match, "endX"), page.Height - Number(match, "endY") }).ToArray();
        TestAssert.Equal(6, paths.Length);
        var width = Regex.Match(afterFill, @"(?m)^(?<w>[\d.]+) w\r?$");
        TestAssert.True(width.Success, "Source brackets declare their stroke before painting.");
        return (paths, Number(width, "w"), Number(fill, "h"));
    }

    private static void AssertPathsEqual(double[][] expected, double[][] actual)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        for (int path = 0; path < expected.Length; path++)
            for (int axis = 0; axis < expected[path].Length; axis++)
                TestAssert.True(Math.Abs(expected[path][axis] - actual[path][axis]) < .002d,
                    "Non-rendering metadata, run splits and authored spacing preserve the source-owned bracket shape.");
    }

    private static void AssertContent(PdfPage page, JsonElement source)
    {
        string text = Regex.Replace(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Select(show => show.Text)), @"\s+", "");
        foreach (var expected in source.GetProperty("ExpectedText").EnumerateArray())
            TestAssert.Contains(Regex.Replace(expected.GetString()!, @"\s+", ""), text);
    }

    private static PdfPage Render(DocxDocument document, IFontResolver? resolver, CancellationToken token = default) =>
        new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(document, null, token).ToBlockingEnumerable().Single();

    private static DocxDocument Read(JsonElement source, CancellationToken token)
    {
        string path = Path.GetFullPath(Path.Combine(CasesDirectory, "..", Fixture(source) + ".docx"));
        using var stream = File.OpenRead(path);
        TestAssert.Equal(source.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
        stream.Position = 0;
        using var package = OoxPackage.Open(stream, token);
        return new DocxReader().Read(package, null, token, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
    }

    private static string Fixture(JsonElement source)
    {
        const string prefix = "tests/Lokad.OoxPdf.Tests/Cases/";
        string input = source.GetProperty("Input").GetString()!;
        TestAssert.True(input.StartsWith(prefix, StringComparison.Ordinal), "Comment frame inputs remain public.");
        return input[prefix.Length..^5];
    }

    private static double Number(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    private static string CasesDirectory => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CommentGlyphFrames"));
    private static JsonDocument ReadJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.Combine(CasesDirectory, name)));
}
