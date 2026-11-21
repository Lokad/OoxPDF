using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionListLabelFamiliesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent emitted label faces use installed Word-compatible fonts.");
        CheckListLabelFamilies(null);
    }

    internal static void CheckListLabelFamilies(IFontResolver? resolver)
    {
        string folder = TestFixtures.PublicCasePath("CompositionProbes");
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "list-families-office.json")));
        foreach (var row in RequireListLabelReferenceRows(golden, "list-families-office.json"))
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = document.FooterParagraphs.Where(p => p.ListLabel is not null).ToArray();
            var labels = row.GetProperty("Labels").EnumerateArray().ToArray();
            TestAssert.Equal(3, paragraphs.Length);
            foreach (var paragraph in paragraphs)
            {
                var label = paragraph.ListLabel!;
                var expected = labels.Single(value => value.GetProperty("Text").GetString() == label.Text);
                var run = DocxLayoutEngine.CreateListLabelRun(label, paragraph.Runs.FirstOrDefault(), 11d);
                TestAssert.Equal(expected.GetProperty("Family").GetString(),
                    DocxFontResolver.ResolveTypeface(run, document.FontCatalog).Primary);
            }
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            var paints = ReadListLabelPaintFonts(page, labels, id);
            if (resolver is null)
                for (int i = 0; i < labels.Length; i++)
                    TestAssert.Equal(labels[i].GetProperty("Family").GetString(), paints[i].Font.FamilyName);
        }
    }

    private static JsonElement.ArrayEnumerator RequireListLabelReferenceRows(JsonDocument golden, string name)
    {
        JsonElement rows = golden.RootElement.GetProperty("Rows");
        TestAssert.True(rows.GetArrayLength() != 0, $"The list label reference {name} has no rows.");
        return rows.EnumerateArray();
    }

    private static PdfEmbeddedFont[] ReadListLabelPaintFonts(PdfPage page, JsonElement[] labels, string id)
    {
        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
            includeTextTransforms: true, honorActualText: false).ToArray();
        string[] paragraphStarts = ["Summary:", "The two relationships", "The trend marker"];
        TestAssert.Equal(paragraphStarts.Length, labels.Length);
        var fonts = new PdfEmbeddedFont[labels.Length];
        for (int i = 0; i < labels.Length; i++)
        {
            string label = labels[i].GetProperty("Text").GetString()!;
            var paints = shows.Where(show => show.Text.Trim() == label).OrderBy(show => show.X).ToArray();
            // The portable face draws synthetic bold a second time on the same baseline.
            bool syntheticBold = paints.Length == 2 && labels[i].TryGetProperty("Bold", out var bold) &&
                bold.GetBoolean() && ReferenceEquals(paints[0].Font, paints[1].Font) &&
                paints[0].Size == paints[1].Size && paints[0].Y == paints[1].Y &&
                paints[1].X - paints[0].X is > 0d and <= .351d;
            TestAssert.True(paints.Length == 1 || syntheticBold,
                $"{id}: expected one painted '{label}', found {paints.Length} text shows.");
            var paint = paints[0];
            TestAssert.True(paint.Y > 0d && paint.Y < page.Height, $"{id}: label '{label}' leaves the page.");
            string followingText = string.Concat(shows.Where(show => show.X > paints[^1].X && Math.Abs(show.Y - paint.Y) < 2d)
                .OrderBy(show => show.X).Select(show => show.Text));
            TestAssert.True(followingText.TrimStart().StartsWith(paragraphStarts[i], StringComparison.Ordinal),
                $"{id}: label '{label}' must precede '{paragraphStarts[i]}' on its printed line. Found '{followingText}'.");
            fonts[i] = paint.Font;
        }
        return fonts;
    }
}
