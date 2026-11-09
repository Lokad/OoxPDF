using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxWrappedMinimumLeadingTests
{
    public static void DocxWrappedMinimumLeadingMatchesIndependentWordOrigins()
    {
        if (!OperatingSystem.IsWindows() || !File.Exists(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "calibri.ttf")))
            TestAssert.Skip("Independent wrapped-minimum origins require installed Windows fonts.");
        using var gold = ReadGold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var pages = Renderer(row).RenderPagesAsync(Read(row), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var chars = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .SelectMany(show => Compact(show.Text).Select(c => (Character: c, Page: index, show.X, show.Y, show.Size)))).ToArray();
            string text = new(chars.Select(c => c.Character).ToArray());
            string expected = row.GetProperty("ExpectedBodyText").GetString()!;
            int start = text.IndexOf(expected, StringComparison.Ordinal);
            TestAssert.True(start >= 0 && text.IndexOf(expected, start + 1, StringComparison.Ordinal) < 0,
                "The complete source body survives once: " + row.GetProperty("Id").GetString());
            foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
            {
                var actual = chars[start + anchor.GetProperty("Offset").GetInt32()];
                TestAssert.True(actual.Character.ToString() == anchor.GetProperty("Character").GetString() &&
                    actual.Page == anchor.GetProperty("Page").GetInt32() &&
                    Math.Abs(actual.X - anchor.GetProperty("X").GetDouble()) <= 1d &&
                    Math.Abs(actual.Y - anchor.GetProperty("Y").GetDouble()) <= 1d &&
                    Math.Abs(actual.Size - anchor.GetProperty("Size").GetDouble()) <= .1d,
                    $"Minimum slots keep their independent Word origin: {row.GetProperty("Id").GetString()}/{row.GetProperty("Mode").GetString()}/{actual}.");
            }
        }
    }

    public static void DocxWrappedMinimumLeadingKeepsPortableSlotsAndShading()
    {
        using var gold = ReadGold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray()
            .Where(row => row.GetProperty("Mode").GetString() == "final-preserve"))
        {
            var document = Read(row);
            var resolver = new PortableRenderingTests.PinnedResolver();
            var measurer = DocxFontPlanTextMeasurer.CreateAsync(DocxFontPlan.Create(document, resolver, default),
                resolver.Resolve(new FontRequest("Calibri")), default, resolver).GetAwaiter().GetResult();
            var original = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, measurer, default);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(document, original, out var placed),
                "Wrapped paragraphs retain their independently measured minimum slots.");
            TestAssert.Equal(original.Pages.Count, placed.Pages.Count);
            var before = Lines(original); var after = Lines(placed);
            TestAssert.Equal(before.Length, after.Length);
            foreach (var pair in before.Zip(after))
            {
                double excess = pair.First.LineHeight!.Value - pair.First.SingleLineHeight!.Value;
                TestAssert.True(excess > 0 && Math.Abs(pair.First.BaselineY - pair.Second.BaselineY - excess) < .00001d,
                    "Every natural baseline occupies the bottom of its measured minimum slot.");
                TestAssert.True(ReferenceEquals(pair.First.Segments, pair.Second.Segments), "Placement retains measured text and source owners.");
                TestAssert.Equal(pair.First.LineHeight, pair.Second.LineHeight);
                TestAssert.Equal(pair.First.SourceLineIndex, pair.Second.SourceLineIndex);
                if (pair.First.IsFirstParagraphLine == true)
                    TestAssert.True(Math.Abs(pair.First.BaselineY + pair.First.FirstLineInsetPoints!.Value -
                        pair.Second.BaselineY - pair.Second.FirstLineInsetPoints!.Value) < .00001d,
                        "The first source line retains its paragraph top.");
                else TestAssert.True(pair.Second.FirstLineInsetPoints is null, "Continuation lines do not become paragraph anchors.");
                if (pair.First.ParagraphShadingFrame is { } paint)
                {
                    var result = pair.Second.ParagraphShadingFrame!.Value;
                    if (pair.First.IsFirstParagraphLine == true)
                        TestAssert.True(Math.Abs(result.Height - (paint.Height - excess)) < .00001d,
                            "The first natural line leaves its minimum leading unshaded.");
                    else
                        TestAssert.True(result.Height == paint.Height &&
                            Math.Abs(pair.First.BaselineY + paint.BaselineInset - pair.Second.BaselineY - result.BaselineInset) < .00001d,
                            "Continuation backgrounds fill their original occupied slots across page breaks.");
                }
            }
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMinimumBodyLineLeading.TryCreateLayout(
                document, original, out _, cancellation.Token));
        }
    }

    public static void DocxWrappedMinimumLeadingKeepsPortableContentAndNavigation()
    {
        using var gold = ReadGold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var pages = Renderer(row, new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(Read(row), null, default)
                .ToBlockingEnumerable().ToArray();
            string text = Compact(string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page))
                .Select(show => show.Text)));
            TestAssert.Contains(row.GetProperty("ExpectedBodyText").GetString()!, text);
            if (row.GetProperty("Id").GetString() == "ln28")
                TestAssert.True(pages.Sum(page => page.Annotations.Count) >= 3, "The wrapped source link retains its clickable lines.");
            if (row.GetProperty("Id").GetString() == "nt28")
            {
                TestAssert.Contains("Boundednotetext.", text);
                TestAssert.True(pages.SelectMany(page => page.Annotations).Any(a => a.Destination is not null),
                    "A note marker retains its actual note destination.");
            }
        }
    }

    private static DocxTextLineLayout[] Lines(DocxLayout layout) => layout.Pages.SelectMany(page => page.Items)
        .OfType<DocxTextLineLayout>().Where(line => line.SourceBlockIndex == 1).ToArray();
    private static string Compact(string text) => new(text.Normalize(NormalizationForm.FormKC).Where(c => !char.IsWhiteSpace(c)).ToArray());
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "WrappedMinimumLeading"));
    private static JsonDocument ReadGold() => JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
    private static DocxDocument Read(JsonElement row)
    {
        string path = Path.Combine(Cases, row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        return DocxTests.ReadDocx(path, row.GetProperty("Mode").GetString() == "final-preserve"
            ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup);
    }
    private static DocxRenderer Renderer(JsonElement row, IFontResolver? resolver = null) =>
        row.GetProperty("Mode").GetString() == "final-preserve"
            ? new(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            : new(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
}
