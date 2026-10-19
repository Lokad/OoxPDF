using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNaturalRowFragmentsKeepIndependentWordBorderCoverage()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent Word composition uses installed Windows fonts.");
        CheckNaturalFragmentBorders(null, compareWord: true);
    }

    internal static void CheckNaturalFragmentBorders(IFontResolver? resolver, bool compareWord = false)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        foreach (string goldenFile in new[] { "parallel-cell-bordered-office.json", "parallel-cell-border-variants-office.json" })
        {
            using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, goldenFile)));
            foreach (var expected in gold.RootElement.GetProperty("Rows").EnumerateArray())
            {
                string id = expected.GetProperty("Id").GetString()!;
                string input = Path.Combine(directory, id + ".docx");
                TestAssert.Equal(expected.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
                bool final = expected.GetProperty("Mode").GetString() == "final-preserve";
                var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
                var document = DocxTests.ReadDocx(input, mode);
                var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var pages = renderer.RenderBlankPages(document, null, default).ToArray();
                var remaining = expected.GetProperty("WordProseAnchors").EnumerateArray()
                    .Select(anchor => anchor.GetProperty("Text").GetString()!).ToHashSet(StringComparer.Ordinal);
                var wordBands = expected.GetProperty("WordBorderBands").EnumerateArray()
                    .Where(band => IsHorizontal(band.GetProperty("Rect")))
                    .GroupBy(band => band.GetProperty("Page").GetInt32()).ToArray();
                TestAssert.True(wordBands.Length >= 2 && wordBands.All(group => group.Count() == 4),
                    "Independent Word fragments have both top and bottom bands for each of the two source cells.");
                int fragments = 0;
                for (int pageIndex = 0; pageIndex < pages.Length; pageIndex++)
                {
                    var page = pages[pageIndex];
                    var bands = Regex.Matches(page.Content,
                        @"(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f")
                        .Select(match => (X: Number(match, "x"), Y: Number(match, "y"), Width: Number(match, "w"), Height: Number(match, "h")))
                        .Where(rect => rect.Width > 100d && rect.Height is > 0d and < 4d).ToArray();
                    var sourceRows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                        .Where(show => show.Text.StartsWith("Left source row ", StringComparison.Ordinal) ||
                            show.Text.StartsWith("Right source row ", StringComparison.Ordinal)).ToArray();
                    if (sourceRows.Length == 0) { TestAssert.Equal(0, bands.Length); continue; }
                    fragments++;
                    TestAssert.Equal(wordBands[0].Count(), bands.Length);
                    foreach (var show in sourceRows)
                    {
                        string text = show.Text.Trim();
                        TestAssert.True(remaining.Remove(text), $"{id}: border painting preserves every source row exactly once.");
                        var cellBands = bands.Where(rect => show.X >= rect.X && show.X < rect.X + rect.Width).ToArray();
                        TestAssert.Equal(2, cellBands.Length);
                        double bottom = cellBands.Min(rect => rect.Y + rect.Height), top = cellBands.Max(rect => rect.Y);
                        var font = show.Font.Font;
                        foreach (var rune in show.Text.EnumerateRunes())
                        {
                            if (System.Text.Rune.IsWhiteSpace(rune)) continue;
                            TestAssert.True(font.TryGetGlyphBounds(font.MapCodePoint(rune.Value), out var bounds), "Source glyph ink is measurable.");
                            TestAssert.True(show.Y + bounds.YMin * show.Size / font.UnitsPerEm >= bottom - .001d &&
                                show.Y + bounds.YMax * show.Size / font.UnitsPerEm <= top + .001d,
                                $"{id}: {text} must clear both painted page-edge borders.");
                        }
                        if (!compareWord) continue;
                        var anchor = expected.GetProperty("WordProseAnchors").EnumerateArray().Single(anchor => anchor.GetProperty("Text").GetString() == text);
                        TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), pageIndex + 1);
                        TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                            Math.Abs(page.Height - show.Y - anchor.GetProperty("Y").GetDouble()) < 1d &&
                            Math.Abs(show.Size - anchor.GetProperty("Size").GetDouble()) < .1d,
                            "Source pages, origins and size retain independent Word gates.");
                    }
                }
                TestAssert.True(fragments >= 2, "The row really spans physical page fragments.");
                TestAssert.Equal(0, remaining.Count);
                using var canceled = new CancellationTokenSource();
                canceled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(document, null, canceled.Token).ToArray());
            }
        }

        static double Number(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);

        static bool IsHorizontal(JsonElement rect) => rect[2].GetDouble() - rect[0].GetDouble() > 100d &&
            rect[3].GetDouble() - rect[1].GetDouble() is > 0d and < 4d;
    }
}
