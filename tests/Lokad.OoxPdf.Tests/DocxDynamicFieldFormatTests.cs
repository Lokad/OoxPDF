using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxDynamicFieldFormatTests
{
    private static readonly string[] GoldFiles = ["minimum-passive-page-fields-office.json",
        "minimum-field-code-office.json", "minimum-field-mark-office.json", "minimum-field-format-office.json"];

    public static void DocxDynamicFieldFormatUsesSourceOwnersAcrossSimpleComplexAndPreservedResults()
    {
        foreach (string file in GoldFiles)
        {
            using var gold = JsonDocument.Parse(File.ReadAllText(Case(file)));
            foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
            {
                var document = DocxTests.ReadDocx(Case(row.GetProperty("Id").GetString() + ".docx"),
                    row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup);
                var paragraph = document.HeaderParagraphs.Concat(document.FooterParagraphs).Single();
                var run = paragraph.Runs.Single(run => run.FieldKind is DocxFieldKind.Page or DocxFieldKind.NumPages);
                double expected = row.GetProperty("WordFieldGlyphs")[0].GetProperty("Size").GetDouble();
                TestAssert.True(Math.Abs(run.FontSize - expected) < .1d,
                    "Computed fields use instruction formatting; explicit MERGEFORMAT retains cached result formatting.");
                var field = paragraph.FieldReferences.Single();
                TestAssert.True(field.UsesPlaceholder && !field.RendersCachedResult && field.HasCachedResult,
                    "Formatting ownership does not replace dynamic field identity with stale result text.");
            }
        }
        foreach (string kind in new[] { "header", "footer" })
        {
            var document = DocxTests.ReadDocx(Case("math-page-fields-simple-" + kind + ".docx"), OoxPdfDocxMarkupMode.Final);
            var runs = document.HeaderParagraphs.Concat(document.FooterParagraphs).SelectMany(paragraph => paragraph.Runs)
                .Where(run => run.FieldKind is DocxFieldKind.Page or DocxFieldKind.NumPages).ToArray();
            TestAssert.Equal(2, runs.Length);
            TestAssert.True(runs.All(run => run.FontFamily == "Calibri" && run.FontSize == 12d),
                "The independent math/story controls inherit instruction Calibri instead of cached Arial.");
        }
    }

    public static void DocxDynamicFieldFormatMatchesIndependentWordGlyphOriginsAndSizes()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent field origins require installed Windows fonts.");
        CheckGlyphs(null, compareWord: true);
    }

    internal static void CheckGlyphs(IFontResolver? resolver, bool compareWord)
    {
        foreach (string file in GoldFiles)
        {
            using var gold = JsonDocument.Parse(File.ReadAllText(Case(file)));
            foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
            {
                string input = Case(row.GetProperty("Id").GetString() + ".docx");
                TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
                bool final = row.GetProperty("Mode").GetString() == "final-preserve";
                var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
                var geometry = final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var pages = new DocxRenderer(resolver, mode, geometry).RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).ToArray();
                foreach (var expected in row.GetProperty("WordFieldGlyphs").EnumerateArray())
                {
                    var page = pages[expected.GetProperty("Page").GetInt32() - 1];
                    var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                        .Single(show => show.Text.Trim() == expected.GetProperty("Text").GetString());
                    TestAssert.True(Math.Abs(actual.Size - expected.GetProperty("Size").GetDouble()) < .1d,
                        "The evaluated field retains its independently observed glyph size with native and portable fonts.");
                    if (compareWord) TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(page.Height - actual.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                        "The evaluated field retains its independent one-point Word origin.");
                }
            }
        }
    }

    private static string Case(string file) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", file));
}
