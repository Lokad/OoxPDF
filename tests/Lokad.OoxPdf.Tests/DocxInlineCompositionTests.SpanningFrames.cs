using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSpanningFramesRetainCompleteInkAndIndependentWordOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent Word origins use installed Windows fonts.");
        CheckSpanningFrames(null, compareWord: true);
    }

    internal static void CheckSpanningFrames(IFontResolver? resolver, bool compareWord)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "spanning-frame-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(directory, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(input, mode);
            var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
            var layout = renderer.InspectLayout(source);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Continuation keeps supported child mathematics.");
            var anchors = row.GetProperty("WordProseAnchors").EnumerateArray().ToArray();
            var remaining = anchors.Select(anchor => anchor.GetProperty("Text").GetString()!).ToHashSet(StringComparer.Ordinal);
            for (int pageIndex = 0; pageIndex < pages.Length; pageIndex++)
            {
                var page = pages[pageIndex];
                foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => show.Text.StartsWith("Planning row ", StringComparison.Ordinal)))
                {
                    string text = show.Text.Trim();
                    TestAssert.True(remaining.Remove(text), $"{id}: every source planning paragraph is painted exactly once: {text}.");
                    var cell = layout.Pages[pageIndex].TableRows.Single().Cells.Single();
                    var font = show.Font.Font;
                    foreach (var rune in show.Text.EnumerateRunes())
                    {
                        if (Rune.IsWhiteSpace(rune)) continue;
                        TestAssert.True(font.TryGetGlyphBounds(font.MapCodePoint(rune.Value), out var bounds), "Planning glyph ink is measurable.");
                        double bottom = show.Y + bounds.YMin * show.Size / font.UnitsPerEm;
                        double top = show.Y + bounds.YMax * show.Size / font.UnitsPerEm;
                        TestAssert.True(bottom >= cell.Y - .001d && top <= cell.Y + cell.Height + .001d,
                            $"{id}, page{pageIndex + 1}: {text} must keep complete glyph ink inside its painted fragment.");
                    }
                    if (!compareWord) continue;
                    var expected = anchors.Single(anchor => anchor.GetProperty("Text").GetString() == text);
                    TestAssert.Equal(expected.GetProperty("Page").GetInt32(), pageIndex + 1);
                    TestAssert.True(Math.Abs(show.X - expected.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(page.Height - show.Y - expected.GetProperty("Y").GetDouble()) < 1d &&
                        Math.Abs(show.Size - expected.GetProperty("Size").GetDouble()) < .1d,
                        $"{id}, page{pageIndex + 1}: {text} retains independent Word origins and size.");
                }
            }
            TestAssert.Equal(0, remaining.Count);
            TestAssert.True(pages.Length >= 3, "The source row really spans pages before the following explicit break.");
            if (compareWord) TestAssert.Equal(row.GetProperty("WordPages").GetInt32(), pages.Length);
            string painted = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (var expected in row.GetProperty("ExpectedText").EnumerateArray()) TestAssert.Contains(expected.GetString()!, painted);
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(source, null, canceled.Token).ToArray());
        }
    }
}
