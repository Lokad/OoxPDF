using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNumberingMarkDecorationsKeepActualSourceOwners() =>
        CheckNumberingMarkDecorations(null, render: false);

    public static void DocxInlineCompositionNumberingMarkDecorationsRenderAcrossStories()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native geometry uses installed Word-compatible fonts.");
        CheckNumberingMarkDecorations(null, render: true);
    }

    internal static void CheckNumberingMarkDecorations(IFontResolver? resolver, bool render)
    {
        string folder = Path.GetDirectoryName(NumberingMarkCase("numbering-mark-decorations-office.json"))!;
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "numbering-mark-decorations-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(document).Where(p => p.ListLabel?.Text == "1.")
                .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).Single();
            var first = paragraph.Runs[0];
            var mark = paragraph.ParagraphMarkRun!;
            var insertion = mark.Revisions.Single();
            var placementRun = first with { CharacterSpacingPoints = .25d, PairKerningEnabled = false };
            var label = DocxLayoutEngine.CreateListLabelRun(paragraph.ListLabel!, placementRun, 11d);
            TestAssert.True(ReferenceEquals(label.Revision, insertion) && label.Revisions.Any(r => ReferenceEquals(r, insertion)),
                "The actual numbered mark owns prefix decoration, independently of the first text revision: " + id);
            TestAssert.Equal(.25d, label.CharacterSpacingPoints);
            TestAssert.True(!label.PairKerningEnabled, "Marker ink cannot discard placement-time kerning decisions.");
            int underlines = row.GetProperty("UnderlineLines").GetInt32();
            TestAssert.Equal(underlines != 0, label.Underline);
            if (underlines != 0) TestAssert.Equal(underlines == 2 ? "double" : "single", label.UnderlineValue);
            if (final) TestAssert.Equal(row.GetProperty("FirstMarkerColor").GetString(), label.ColorHex ?? "000000");
            TestAssert.True(ReferenceEquals(first, paragraph.Runs[0]), "Marker ownership preserves source text objects.");
            if (!id.Contains("math", StringComparison.Ordinal))
                TestAssert.True(first.ColorHex == "1964AE" && first.UnderlineValue == "double",
                    "Authored first-run blue/double underline must stay on its source.");
            if (!render) continue;
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            var paint = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                .First(show => show.Text.Trim() == "1.");
            TestAssert.True(paint.Y > 0d && paint.Y < page.Height, "The decorated marker must remain visible: " + id);
            if (resolver is null)
                TestAssert.True(Math.Abs(paint.X - row.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(paint.Y - row.GetProperty("Y").GetDouble()) < 1d,
                    "Marker decoration preserves the independently qualified one-point geometry tolerance: " + id);
        }
    }
}
