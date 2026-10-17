using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionUnchangedListDecorationsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native paint checks use installed Word-compatible fonts.");
        CheckUnchangedListDecorations(null);
    }

    internal static void CheckUnchangedListDecorations(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "list-decorations-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            // Those independent facts remain in the golden as a separate known
            // revision/renumbering gap; this slice qualifies unchanged markers.
            if (row.GetProperty("HasParagraphMarkRevision").GetBoolean()) continue;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = document.FooterParagraphs.Where(p => p.ListLabel is not null).ToArray();
            var owners = paragraphs.SelectMany(p => p.Runs).ToArray();
            TestAssert.Equal(3, paragraphs.Length);
            var labels = row.GetProperty("Labels").EnumerateArray().ToArray();
            foreach (var paragraph in paragraphs)
            {
                var label = paragraph.ListLabel!;
                var expected = labels.Single(value => value.GetProperty("Text").GetString() == label.Text);
                var run = DocxLayoutEngine.CreateListLabelRun(label, paragraph.Runs.FirstOrDefault(), 11d);
                TestAssert.Equal(expected.GetProperty("ColorHex").GetString(), run.ColorHex ?? "000000");
                int underlines = expected.GetProperty("UnderlineLines").GetInt32();
                TestAssert.Equal(underlines != 0, run.Underline);
                if (underlines == 2) TestAssert.Equal("double", run.UnderlineValue);
                TestAssert.True(run.Revision is null && run.Revisions.Count == 0,
                    "Text-scoped revisions cannot acquire an otherwise unchanged marker.");
            }
            var first = paragraphs[0].Runs[0];
            TestAssert.True(first.ColorHex == "1964AE" && first.Underline && first.UnderlineValue == "double",
                "First-run authored color/underline stay on their original owner.");
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
            foreach (var expected in labels)
                TestAssert.True(shows.Any(show => show.Text.Trim() == expected.GetProperty("Text").GetString() && show.Y > 0d && show.Y < page.Height),
                    "Decoration ownership cannot remove a visible numbering marker.");
            TestAssert.True(owners.Zip(paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Resolving marker decoration cannot rewrite source runs.");
        }
    }
}
