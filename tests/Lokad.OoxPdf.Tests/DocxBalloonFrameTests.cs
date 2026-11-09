using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxBalloonFrameTests
{
    public static void DocxBalloonFramesUnitScaleEdgesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent balloon frames use installed Windows fonts.");
        using var sources = ReadJson("sources.json");
        using var gold = ReadJson("office.json");
        foreach (var source in sources.RootElement.EnumerateArray().Where(row => row.GetProperty("UnitScale").GetBoolean()))
        {
            var expected = FindGold(gold, source);
            var (placement, page) = Render(Read(source, default), null);
            double[] actual = [placement.X, page.Height - placement.Y - placement.Height, placement.X + placement.Width, page.Height - placement.Y];
            for (int axis = 0; axis < actual.Length; axis++)
                TestAssert.True(Math.Abs(actual[axis] - expected.GetProperty("Bounds")[axis].GetDouble()) < 1d,
                    "Unit-scale balloons use the shared lane and measured text frame: " + source.GetProperty("Id").GetString() + "/" + axis);
        }
    }

    public static void DocxBalloonFramesVerticalBoundsAndTitleOriginsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent balloon origins use installed Windows fonts.");
        using var sources = ReadJson("sources.json");
        using var gold = ReadJson("office.json");
        foreach (var source in sources.RootElement.EnumerateArray())
        {
            var expected = FindGold(gold, source);
            var (placement, page) = Render(Read(source, default), null);
            var shows = BalloonShows(page, placement);
            var title = shows.Single(show => show.Text.StartsWith("Commented", StringComparison.Ordinal));
            TestAssert.True(Math.Abs(page.Height - placement.Y - placement.Height - expected.GetProperty("Bounds")[1].GetDouble()) < 1d &&
                Math.Abs(page.Height - placement.Y - expected.GetProperty("Bounds")[3].GetDouble()) < 1d,
                "Source tops and measured row font boxes retain independent Word vertical edges: " + source.GetProperty("Id").GetString());
            TestAssert.True(Math.Abs(title.X - expected.GetProperty("FirstOrigin")[0].GetDouble()) < 1d &&
                Math.Abs(page.Height - title.Y - expected.GetProperty("FirstOrigin")[1].GetDouble()) < 1d,
                "Balloon title origins retain the unchanged absolute Word threshold: " + source.GetProperty("Id").GetString());
            TestAssert.Equal(expected.GetProperty("Rows").GetInt32(), shows.Select(show => Math.Round(show.Y, 2)).Distinct().Count());
            TestAssert.Equal(Normalize(expected.GetProperty("Text").GetString()!), Normalize(string.Concat(shows.Select(show => show.Text))));
            if (source.TryGetProperty("ExpectedCommentFamilies", out var families))
                foreach (var family in families.EnumerateArray())
                    TestAssert.True(shows.Where(show => !show.Text.StartsWith("Commented", StringComparison.Ordinal))
                        .Any(show => show.Font.Font.FamilyName.Equals(family.GetString(), StringComparison.OrdinalIgnoreCase)),
                        "Wrapping preserves the declared comment face: " + family.GetString());
        }
    }

    public static void DocxBalloonFramesPreservePortableContentAndCancellation()
    {
        using var sources = ReadJson("sources.json");
        using var gold = ReadJson("office.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var source in sources.RootElement.EnumerateArray().Where(row => row.GetProperty("PortableProbe").GetBoolean()))
        {
            var (placement, page) = Render(Read(source, default), resolver);
            var shows = BalloonShows(page, placement);
            string allText = Normalize(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Select(show => show.Text)));
            foreach (var expected in source.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.True(allText.Contains(Normalize(expected.GetString()!), StringComparison.Ordinal), "Balloon layout preserves body prose.");
            TestAssert.Equal(Normalize(FindGold(gold, source).GetProperty("Text").GetString()!), Normalize(string.Concat(shows.Select(show => show.Text))));
            TestAssert.True(shows.All(show => show.Y > placement.Y && show.Y < placement.Y + placement.Height),
                "Painted portable rows stay inside their measured balloon.");
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Read(sources.RootElement[0], canceled.Token));
    }

    public static void DocxBalloonFramesKeepBookmarksAndRunSplitsNeutral()
    {
        using var sources = ReadJson("sources.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var source in sources.RootElement.EnumerateArray().Where(row => row.GetProperty("UnitScale").GetBoolean()))
        {
            var (original, originalPage) = Render(Read(source, default), resolver);
            foreach ((bool bookmark, bool split) in new[] { (true, false), (false, true), (true, true) })
            {
                const string prefix = "tests/Lokad.OoxPdf.Tests/Cases/";
                string fixture = source.GetProperty("Input").GetString()![prefix.Length..^5];
                var document = DocxInlineCompositionTests.ReadMutation(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, split,
                    (xml, w, _) =>
                    {
                        if (!bookmark) return;
                        var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(w + "commentRangeStart").Any());
                        var start = new XElement(w + "bookmarkStart", new XAttribute(w + "id", "8220"), new XAttribute(w + "name", "PublicBalloonFrame"));
                        if (paragraph.Element(w + "pPr") is { } properties) properties.AddAfterSelf(start); else paragraph.AddFirst(start);
                        paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "8220")));
                    });
                var (actual, page) = Render(document, resolver);
                TestAssert.True(Math.Abs(original.X - actual.X) < .01d && Math.Abs(original.Y - actual.Y) < .01d &&
                    Math.Abs(original.Width - actual.Width) < .01d && Math.Abs(original.Height - actual.Height) < .01d,
                    "Non-rendering source metadata and identical-format run splits cannot change the balloon frame.");
                TestAssert.Equal(Normalize(string.Concat(BalloonShows(originalPage, original).Select(show => show.Text))),
                    Normalize(string.Concat(BalloonShows(page, actual).Select(show => show.Text))));
            }
        }
    }

    private static (DocxMarkupBalloonPlacementSnapshot Placement, PdfPage Page) Render(DocxDocument document, IFontResolver? resolver)
    {
        var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var placement = renderer.InspectMarkupBalloons(document).Single();
        return (placement, renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single());
    }

    private static (PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] BalloonShows(PdfPage page, DocxMarkupBalloonPlacementSnapshot placement) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
            .Where(show => show.X >= placement.X && show.X < placement.X + placement.Width && !string.IsNullOrWhiteSpace(show.Text)).ToArray();

    private static DocxDocument Read(JsonElement source, CancellationToken cancellationToken)
    {
        const string prefix = "tests/Lokad.OoxPdf.Tests/Cases/";
        string input = source.GetProperty("Input").GetString()!;
        TestAssert.True(input.StartsWith(prefix, StringComparison.Ordinal), "Balloon inputs stay public.");
        string path = Path.GetFullPath(Path.Combine(CasesDirectory, "..", input[prefix.Length..]));
        using var stream = File.OpenRead(path);
        TestAssert.Equal(source.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
        stream.Position = 0;
        using var package = OoxPackage.Open(stream, cancellationToken);
        return new DocxReader().Read(package, null, cancellationToken, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
    }

    private static JsonElement FindGold(JsonDocument gold, JsonElement source)
    {
        var row = gold.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == source.GetProperty("Id").GetString());
        TestAssert.Equal(source.GetProperty("Sha256").GetString(), row.GetProperty("InputSha256").GetString());
        return row;
    }

    private static string Normalize(string text) => Regex.Replace(text, @"\s+", "");
    private static string CasesDirectory => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "BalloonFrames"));
    private static JsonDocument ReadJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.Combine(CasesDirectory, name)));
}
