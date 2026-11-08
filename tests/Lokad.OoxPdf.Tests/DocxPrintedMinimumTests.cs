using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPrintedMinimumTests
{
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ScaledMinimum"));

    public static void DocxPrintedMinimumMatchesIndependentWordParagraphOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent printed minimum origins require installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = Path.Combine(Cases, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var source = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.AllMarkup);
            var pages = Renderer().RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            // Ignore line-break whitespace, while retaining the painted origin
            // and page of every character. Full source paragraphs are unique.
            var chars = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .SelectMany(show => show.Text.Where(c => !char.IsWhiteSpace(c))
                    .Select(c => (Character: c, Page: index, show.X, show.Y, show.Size)))).ToArray();
            string text = new(chars.Select(c => c.Character).ToArray());
            foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
            {
                string key = anchor.GetProperty("Text").GetString()!;
                int offset = text.IndexOf(key, StringComparison.Ordinal);
                TestAssert.True(offset >= 0 && text.IndexOf(key, offset + 1, StringComparison.Ordinal) < 0,
                    "Each complete source paragraph must survive exactly once: " + row.GetProperty("Id").GetString());
                var actual = chars[offset];
                TestAssert.True(actual.Page == anchor.GetProperty("Page").GetInt32() &&
                    Math.Abs(actual.X - anchor.GetProperty("X").GetDouble()) <= 1d &&
                    Math.Abs(actual.Y - anchor.GetProperty("Y").GetDouble()) <= 1d &&
                    Math.Abs(actual.Size - anchor.GetProperty("Size").GetDouble()) <= .1d,
                    $"Printed minimum paragraph keeps its independent page/origin/size: {row.GetProperty("Id").GetString()}/{key}/{actual}.");
            }
        }
    }

    public static void DocxPrintedMinimumKeepsPortableContentAndKeepChains()
    {
        foreach (string name in new[] { "minimum28-wrapped", "minimum28-kept", "tahoma10-minimum26" })
        {
            var source = DocxTests.ReadDocx(Path.Combine(Cases, name + ".docx"), OoxPdfDocxMarkupMode.AllMarkup);
            var renderer = Renderer(new PortableRenderingTests.PinnedResolver());
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            string painted = string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page))
                .Select(show => string.Concat(show.Text.Where(c => !char.IsWhiteSpace(c)))));
            foreach (var paragraph in source.Paragraphs)
            {
                string text = string.Concat(paragraph.Runs.SelectMany(run => run.Text).Where(c => !char.IsWhiteSpace(c)));
                TestAssert.True(painted.Contains(text, StringComparison.Ordinal), "Wrapping and kept placement retain complete source text.");
            }
            if (name == "minimum28-kept")
            {
                var lines = renderer.InspectTextEmission(source).Lines.Where(line => line.ContainerStoryKind == "Body").ToArray();
                TestAssert.Equal(lines.First(line => line.SourceBlockIndex == 24).PageIndex,
                    lines.First(line => line.SourceBlockIndex == 26).PageIndex);
            }
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxRenderer Renderer(Lokad.OoxPdf.Fonts.IFontResolver? resolver = null) =>
        new(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
}
