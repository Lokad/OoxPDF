using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionPlainNoteBatchKeepsDeferredOwnersOrderedAndDisjoint()
    {
        var source = DocxTests.ReadDocx(NestedMathCasePath("two-plain-notes-report.docx"), OoxPdfDocxMarkupMode.Final);
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
        var placed = layout.Pages.SelectMany(page => page.PlacedRelatedStories.Where(story =>
            story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.IsNormalStoryType)).ToArray();
        var expected = layout.RelatedStories.Where(story => story.Story.Kind == DocxRelatedStoryKind.Footnote && story.Story.IsNormalStoryType).ToArray();
        TestAssert.Equal(2, expected.Length);
        var actualOwners = placed.Select(story => story.StoryLayout.Story.Id).Distinct().ToArray();
        TestAssert.True(actualOwners.SequenceEqual(expected.Select(story => story.Story.Id)),
            "Deferred heads retain their original source order across the drain and postpass.");
        bool secondStarted = false;
        foreach (var story in placed)
        {
            if (story.StoryLayout.Story.Id == expected[1].Story.Id) secondStarted = true;
            TestAssert.True(!secondStarted || story.StoryLayout.Story.Id == expected[1].Story.Id,
                "A preceding note tail cannot appear after the next source note has started.");
        }
        foreach (var story in expected)
        {
            var lines = placed.Where(part => part.StoryLayout.Story.Id == story.Story.Id)
                .SelectMany(part => part.StoryLayout.TextLines).ToArray();
            TestAssert.True(story.TextLines.Select(line => (line.Text, line.SourceParagraph, line.SourceLineIndex))
                .SequenceEqual(lines.Select(line => (line.Text, line.SourceParagraph, line.SourceLineIndex))),
                "Each canonical source line and paragraph owner is placed exactly once.");
        }
        foreach (var page in layout.Pages)
        {
            var notes = page.PlacedRelatedStories.Where(story => story.StoryLayout.Story.IsNormalStoryType &&
                story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote).ToArray();
            if (notes.Length == 0) continue;
            TestAssert.Equal(1, page.PlacedRelatedStories.Count(story => story.SeparatorY is not null));
            for (int index = 1; index < notes.Length; index++)
                TestAssert.True(notes[index].TopY <= notes[index - 1].TopY - notes[index - 1].Height + .000001d,
                    "A shared note compartment gives each source owner a disjoint placed container.");
        }
    }
}
