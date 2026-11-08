using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxParagraphShadingTests
{
    public static void DocxParagraphShadingPrintedFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent printed background frames require installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "printed-office.json")));
        var colors = gold.RootElement.GetProperty("Palette").EnumerateArray().Select(value => value.GetString()!).ToHashSet();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = Path.GetFullPath(Path.Combine(Cases, "..", row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var source = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.AllMarkup);
            var pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var actual = pages.SelectMany((page, index) => ReadRectangles(page.Content)
                .Where(rectangle => colors.Contains(rectangle.Color)).Select(rectangle => (Page: index, Frame: rectangle))).ToArray();
            var expected = row.GetProperty("Frames").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            for (int i = 0; i < actual.Length; i++)
            {
                var a = actual[i]; var b = expected[i];
                TestAssert.True(a.Page == b.GetProperty("Page").GetInt32() && a.Frame.Color == b.GetProperty("Color").GetString() &&
                    Math.Abs(a.Frame.X - b.GetProperty("X").GetDouble()) <= 1d &&
                    Math.Abs(a.Frame.Y - b.GetProperty("Y").GetDouble()) <= 1d &&
                    Math.Abs(a.Frame.Width - b.GetProperty("Width").GetDouble()) <= 1d &&
                    Math.Abs(a.Frame.Height - b.GetProperty("Height").GetDouble()) <= 1d,
                    $"Printed background retains its independent Word page, color and rectangle: {row.GetProperty("Id").GetString()}/{i}/{a}.");
            }
        }
    }

    public static void DocxParagraphShadingPrintedContinuityKeepsPortableTextAndConsumedSpacing()
    {
        foreach (string name in new[] { "wrapped-minimum28", "joined-colors-suppressed", "joined-colors-explicit" })
        {
            var source = DocxTests.ReadDocx(Path.Combine(Cases, "printed-" + name + ".docx"), OoxPdfDocxMarkupMode.AllMarkup);
            var replacements = source.Paragraphs.ToDictionary(paragraph => paragraph, paragraph => paragraph with { Shading = null });
            var replace = new DocxParagraphReplacements(replacements, default);
            var plain = source with { BodyElements = replace.ReplaceElements(source.BodyElements),
                FallbackParagraphs = source.FallbackParagraphs.Select(replace.Replace).ToArray() };
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var plainPages = renderer.RenderPagesAsync(plain, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(plainPages.Length, pages.Length);
            var blocks = source.BodyElements.Select((element, index) => (element, index))
                .Where(item => item.element is DocxParagraphElement { Paragraph.Shading: not null }).Select(item => item.index).ToHashSet();
            var layout = renderer.InspectLayout(source);
            for (int page = 0; page < pages.Length; page++)
            {
                static IEnumerable<(string, double, double, double)> Glyphs(Lokad.OoxPdf.Pdf.PdfPage value) =>
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(value).Select(show => (show.Text, show.Size, show.X, show.Y));
                TestAssert.True(Glyphs(plainPages[page]).SequenceEqual(Glyphs(pages[page])), "Background continuity must preserve every placed glyph.");
                var frames = ReadRectangles(pages[page].Content).Where(frame => frame.Color is "D8E4F0" or "F0D8E4" or "E4F0D8").ToArray();
                var lines = layout.Pages[page].Items.Where(line => line.Kind == "TextLine" && blocks.Contains(line.SourceBlockIndex ?? -1)).ToArray();
                TestAssert.Equal(lines.Length, frames.Length);
                for (int i = 1; i < frames.Length; i++)
                    if (lines[i - 1].SourceBlockIndex == lines[i].SourceBlockIndex || lines[i - 1].SourceBlockIndex + 1 == lines[i].SourceBlockIndex)
                        TestAssert.True(Math.Abs(frames[i - 1].Y - frames[i].Y - frames[i].Height) < .002d,
                            "Connected shaded rows fill their actual leading/spacing without gaps at PDF formatting precision.");
                if (name == "joined-colors-suppressed" && frames.Length != 0)
                    TestAssert.True(Math.Abs(frames[0].Height - lines[0].SingleLineHeightPoints!.Value) < .002d,
                        "A suppressed authored after-value cannot enlarge the preceding color's background.");
            }
        }
    }
}
