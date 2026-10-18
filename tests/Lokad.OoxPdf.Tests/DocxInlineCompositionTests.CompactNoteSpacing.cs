using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCompactNoteSpacingMatchesIndependentWordPaginationAndBaselines()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note-spacing controls use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("compact-note-spacing-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = StaticNoteAdmissionCasePath(row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var source = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.Final);
            var pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(source, null, default).ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Select(show => (Page: index + 1, Show: show))).ToArray();
            var note = shows.Single(item => item.Show.Text.Contains("All quantities", StringComparison.Ordinal));
            var marker = shows.Single(item => item.Show.Text.Contains("Review conclusion:", StringComparison.Ordinal));
            TestAssert.Equal(marker.Page, note.Page);
            TestAssert.True(Math.Abs(note.Show.Y - row.GetProperty("NoteY").GetDouble()) < .5d,
                $"The final note's optional spacing consumes only the available body/separator capacity: {id}/{note.Show.Y}.");
            TestAssert.True(Math.Abs(marker.Show.Y - row.GetProperty("MarkerParagraphY").GetDouble()) < .5d,
                $"Required body spacing keeps the reference on Word's independent page and baseline: {id}.");
        }
    }

    public static void DocxInlineCompositionCompactNoteSpacingRetainsOwnedContentAcrossPageCapacities()
    {
        foreach (string context in new[] { "full", "no-header" })
        foreach (int pageHeight in new[] { 8000, 10400, 13200, 15840 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            int? expectedPages = null;
            foreach (int after in new[] { 0, 6, 12 })
            {
                var source = ReadMutation($"CompositionProbes/compact-note-{context}-note-after{after}", mode,
                    false, false, (xml, w, _) => xml.Descendants(w + "pgSz").Single()
                        .SetAttributeValue(w + "h", pageHeight.ToString(CultureInfo.InvariantCulture)));
                var equations = DocxBlockTraversal.EnumerateBodyParagraphs(source).Concat(
                        DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(source, default).Select(owner => owner.Paragraph))
                    .SelectMany(owner => owner.Equations).Distinct().ToDictionary(equation => equation,
                        equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
                var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
                TestAssert.True(layout.UsesSourceOwnedFootnoteReserve, "Measured source owners retain complete-note admission.");
                expectedPages ??= layout.Pages.Count;
                TestAssert.Equal(expectedPages.Value, layout.Pages.Count);
                var placed = layout.Pages.SelectMany(page => page.PlacedRelatedStories
                    .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.IsNormalStoryType)
                    .Select(story => (Page: page, Note: story))).Single();
                var referenceParagraph = source.Paragraphs.Single(paragraph => paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Footnote));
                TestAssert.True(placed.Page.Items.OfType<DocxTextLineLayout>().Any(line => ReferenceEquals(line.SourceParagraph, referenceParagraph)),
                    "The note and its actual body reference share a physical page.");
                TestAssert.True(placed.Note.Height >= placed.Note.StoryLayout.TextLines.Single().LineHeight!.Value - .001d &&
                    placed.Note.TopY - placed.Note.Height >= placed.Page.MarginBottom - .001d,
                    "Only trailing paragraph space may collapse; the complete note line remains inside the page.");
                TestAssert.Equal((double)after, source.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Footnote &&
                    story.IsNormalStoryType).Paragraphs.Single().SpacingAfterPoints);
            }
        }
    }
}
