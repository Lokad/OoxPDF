using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMinimumBookmarksMatchIndependentWordBaselinesAndPageFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent minimum/bookmark controls use installed Windows fonts.");
        foreach (string name in new[] { "minimum-bookmark-office.json", "minimum-bookmark-static-office.json" })
        {
            using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath(name)));
            foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
            {
                string path = StaticNoteAdmissionCasePath(row.GetProperty("Input").GetString()!);
                TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
                bool final = row.GetProperty("Mode").GetString() == "final-preserve";
                var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
                var source = DocxTests.ReadDocx(path, mode);
                var pages = new DocxRenderer(null, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, default).ToArray();
                TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                var shows = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)).ToArray();
                foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
                {
                    string text = anchor.GetProperty("Text").GetString()!;
                    var show = shows.Single(show => show.Text.Contains(text, StringComparison.Ordinal));
                    TestAssert.True(Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < .5d,
                        $"Neutral metadata and requested Final/Word modes retain the qualified source frame: {row.GetProperty("Id").GetString()}/{mode}/{text}/{show.Y}.");
                }
            }
        }
    }

    public static void DocxMinimumBookmarksKeepSourceMetadataOutOfVisibleFrames()
    {
        var pairs = new[]
        {
            ("ordinary-minimum-line-20.docx", "CompositionProbes/minimum-bookmark-ordinary12-bookmarked.docx"),
            ("CompositionProbes/minimum-bookmark-ordinary11-control.docx", "CompositionProbes/minimum-bookmark-ordinary11-bookmarked.docx"),
            ("note-marked-minimum-line-20.docx", "CompositionProbes/minimum-bookmark-noted12-bookmarked.docx"),
            ("CompositionProbes/plain-automatic-batch-after12.docx", "CompositionProbes/plain-batch-footer-bookmark.docx")
        };
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (control, bookmarked) in pairs)
        {
            var source = DocxTests.ReadDocx(StaticNoteAdmissionCasePath(control), mode);
            var annotated = DocxTests.ReadDocx(StaticNoteAdmissionCasePath(bookmarked), mode);
            var sourceParagraphs = DocxBlockTraversal.EnumerateBodyParagraphs(source).ToArray();
            var annotatedParagraphs = DocxBlockTraversal.EnumerateBodyParagraphs(annotated).ToArray();
            TestAssert.Equal(string.Concat(sourceParagraphs.SelectMany(p => p.Runs).Select(r => r.Text)),
                string.Concat(annotatedParagraphs.SelectMany(p => p.Runs).Select(r => r.Text)));
            var renderer = new DocxRenderer(new TestFaceFontResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderBlankPages(source, null, default).ToArray();
            var after = renderer.RenderBlankPages(annotated, null, default).ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(before[page], usePageCoordinates: true)
                    .Select(show => (show.Text, show.X, show.Y)).ToArray();
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(after[page], usePageCoordinates: true)
                    .Select(show => (show.Text, show.X, show.Y)).ToArray();
                TestAssert.True(expected.SequenceEqual(actual),
                    "A bookmark in an automatic-height neighbor or footer cannot alter supported visible content/geometry elsewhere.");
            }
        }
    }
}
