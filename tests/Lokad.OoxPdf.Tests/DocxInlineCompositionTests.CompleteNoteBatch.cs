using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCompleteNoteBatchMatchesIndependentWordPagesAndTargetBaselines()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent shared-note controls use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("complete-note-batch-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = StaticNoteAdmissionCasePath(row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var source = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.Final);
            var pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(source, null, default).ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Select(show => (Page: index, Show: show))).ToArray();
            var marker = shows.Single(item => item.Show.Text.Contains(row.GetProperty("MarkerText").GetString()!, StringComparison.Ordinal));
            TestAssert.True(Math.Abs(marker.Show.Y - row.GetProperty("MarkerY").GetDouble()) < .5d,
                "Complete source-owned note heads clear the body's required paragraph boundary.");
            foreach (var expected in row.GetProperty("Notes").EnumerateArray())
            {
                var note = shows.Single(item => item.Show.Text.Contains(expected.GetProperty("Text").GetString()!, StringComparison.Ordinal));
                TestAssert.Equal(marker.Page, note.Page);
                TestAssert.True(Math.Abs(note.Show.Y - expected.GetProperty("Y").GetDouble()) < .5d,
                    $"A complete shared note retains its independent Word baseline: {row.GetProperty("Id").GetString()}.");
            }
        }
    }

    public static void DocxInlineCompositionCompleteNoteBatchPreservesCanonicalLinesAndInterNoteSpace()
    {
        foreach (string fixture in new[] { "plain-automatic-note-after12", "plain-automatic-batch-after12" })
        foreach (int pageHeight in new[] { 400, 520, 660, 800 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            int? expectedPages = null;
            foreach (int after in new[] { 0, 12, 24 })
            {
                var source = ReadMutation("CompositionProbes/" + fixture, mode, false, false, (xml, w, _) =>
                    xml.Descendants(w + "footnote").Last(note => note.Attribute(w + "type") is null)
                        .Element(w + "p")!.Element(w + "pPr")!.Element(w + "spacing")!
                        .SetAttributeValue(w + "after", (after * 20).ToString(CultureInfo.InvariantCulture)), partOverride: "word/footnotes.xml");
                var settings = source.PageSettings with { HeightValue = (pageHeight * 20).ToString(CultureInfo.InvariantCulture) };
                source = source with { PageHeightPoints = pageHeight, PageSettings = settings,
                    FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings } };
                var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
                TestAssert.True(layout.UsesSourceOwnedFootnoteReserve, "Complete owners retain their minimum shared head across static frames.");
                expectedPages ??= layout.Pages.Count;
                TestAssert.Equal(expectedPages.Value, layout.Pages.Count);
                var placed = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                    .Where(note => note.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && note.StoryLayout.Story.IsNormalStoryType)
                    .Select(note => (Page: index, Note: note))).ToArray();
                TestAssert.Equal(source.RelatedStories.Count(story => story.Kind == DocxRelatedStoryKind.Footnote && story.IsNormalStoryType), placed.Length);
                TestAssert.Equal(1, placed.Select(item => item.Page).Distinct().Count());
                TestAssert.True(placed.All(item => item.Note.UsesPlainFootnoteFrame), "The control exercises the shared static-frame placement path.");
                var owner = source.Paragraphs.Single(paragraph => paragraph.InlineReferences.Count != 0);
                TestAssert.True(layout.Pages[placed[0].Page].Items.OfType<DocxTextLineLayout>().Any(line => ReferenceEquals(line.SourceParagraph, owner)),
                    "The entire note batch shares its actual source-reference page.");
                foreach (var item in placed)
                {
                    var canonical = layout.RelatedStories.Single(story => ReferenceEquals(story.Story, item.Note.StoryLayout.Story));
                    var line = canonical.TextLines.Single();
                    TestAssert.Equal(line.Text, item.Note.StoryLayout.TextLines.Single().Text);
                    TestAssert.True(item.Note.Height >= line.LineHeight!.Value - .001d,
                        "Only optional trailing space can collapse; no admitted source line can be deferred or clipped away.");
                    TestAssert.Equal(canonical.Story.Paragraphs.Single().SpacingAfterPoints, line.ParagraphAfterSpacing!.Value);
                    if (!ReferenceEquals(item.Note, placed[^1].Note))
                        TestAssert.Equal(canonical.ContentHeight, item.Note.Height);
                }
            }
        }
    }
}
