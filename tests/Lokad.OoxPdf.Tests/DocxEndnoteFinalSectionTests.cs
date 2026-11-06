using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxEndnoteFinalSectionTests
{
    public static void DocxEmptyFinalNextPageSectionOwnsDocumentEndnoteFrame()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var (source, expectedPages, expectedFirstNotePage) = Read(mode);
            TestAssert.True(source.BodyElements.Last() is DocxSectionBreakElement { TypeValue: DocxSectionBreakType.NextPage } &&
                source.FinalSectionBreak is { TypeValue: DocxSectionBreakType.NextPage },
                "The unchanged native control has a parsed break followed by an empty next-page section.");
            var layout = Layout(source);
            TestAssert.Equal(expectedPages, layout.Pages.Count);
            TestAssert.Equal(expectedFirstNotePage, FirstEndnotePage(layout));
            TestAssert.True(layout.Pages[1].Items.Count == 0 && layout.Pages[1].PlacedRelatedStories.Count != 0,
                "The section's first page owns note content even though it has no body items.");
            TestAssert.Equal(81, layout.Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.IsNormalStoryType)
                .Select(note => note.StoryLayout.Story).Distinct().Count());
        }
    }

    public static void DocxEmptyFinalSectionDoesNotCreateUnreferencedNotePages()
    {
        var source = Read(OoxPdfDocxMarkupMode.Final).Source;
        var unresolved = source with { RelatedStories = source.RelatedStories.Where(note => !note.IsNormalStoryType).ToArray() };
        TestAssert.Equal(1, Layout(unresolved).Pages.Count);
        var unreferenced = source with { BodyElements = source.BodyElements.Select(element => element is DocxParagraphElement paragraph
            ? paragraph with { Paragraph = paragraph.Paragraph with { InlineReferences = [] } } : element).ToArray() };
        TestAssert.Equal(1, Layout(unreferenced).Pages.Count);
    }

    public static void DocxEmptyFinalSectionRetainsContinuousAndParityFallbacks()
    {
        var source = Read(OoxPdfDocxMarkupMode.Final).Source;
        foreach (var type in new[] { DocxSectionBreakType.Continuous, DocxSectionBreakType.EvenPage, DocxSectionBreakType.OddPage })
        {
            var layout = Layout(source with { FinalSectionBreak = source.FinalSectionBreak! with { TypeValue = type } });
            TestAssert.Equal(2, layout.Pages.Count);
            TestAssert.Equal(1, FirstEndnotePage(layout));
        }
    }

    public static void DocxSectionEndNotesDoNotClaimEmptyFinalDocumentFrame()
    {
        var source = Read(OoxPdfDocxMarkupMode.Final).Source;
        source = source with { BodyElements = source.BodyElements.Select(element => element is DocxSectionBreakElement section
            ? section with { PageSettings = section.PageSettings with { EndnoteReferenceSettings =
                section.PageSettings.EndnoteReferenceSettings with { PositionValue = "sectEnd" } } } : element).ToArray() };
        var layout = Layout(source);
        // Section-end notes retain the existing generic boundary charges;
        // the parent also needs three pages, starting in the owning section.
        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(1, FirstEndnotePage(layout));
    }

    private static (DocxDocument Source, int Pages, int FirstNotePage) Read(OoxPdfDocxMarkupMode mode)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "endnote-empty-final-section-office.json")));
        string input = Path.Combine(folder, reference.RootElement.GetProperty("CaseRelativePath").GetString()!);
        TestAssert.Equal(reference.RootElement.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        string name = mode == OoxPdfDocxMarkupMode.Final ? "final-preserve" : "all-word-compatible";
        var row = reference.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Mode").GetString() == name);
        return (DocxTests.ReadDocx(input, mode), row.GetProperty("ExpectedPages").GetInt32(), row.GetProperty("FirstEndnotePage").GetInt32());
    }

    private static DocxLayout Layout(DocxDocument source) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
        .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);

    private static int FirstEndnotePage(DocxLayout layout) => layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
        .Where(note => note.StoryLayout.Story.Id == "37").Select(_ => index + 1)).Min();
}
