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
            TestAssert.True(layout.Pages[1].Items is [DocxTextLineLayout { SourceParagraph: not null } mark] &&
                string.IsNullOrWhiteSpace(mark.Text) && layout.Pages[1].PlacedRelatedStories.Count != 0,
                "The implicit paragraph uses the ordinary source-owned body flow before the notes.");
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

    public static void DocxImplicitFinalSectionParagraphUsesOrdinaryFlow()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (family, size) in new[] { ("Calibri", 12d), ("Calibri", 14d), ("Arial", 12d) })
        {
            var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/endnote-empty-final-section", mode,
                false, false, (xml, w, _) =>
                {
                    var defaults = xml.Root!.Element(w + "docDefaults")!.Element(w + "rPrDefault")!.Element(w + "rPr")!;
                    defaults.Element(w + "sz")!.SetAttributeValue(w + "val", size * 2d);
                    defaults.Element(w + "rFonts")!.SetAttributeValue(w + "ascii", family);
                    defaults.Element(w + "rFonts")!.SetAttributeValue(w + "hAnsi", family);
                }, partOverride: "word/styles.xml");
            var blank = source.EmptyFinalSectionParagraph ?? throw new InvalidOperationException("Resolve the source paragraph defaults.");
            TestAssert.True(blank.Runs.Count != 0 && blank.Runs.All(run => run.FontSize == size && run.FontFamily == family),
                "The implicit paragraph inherits the source default font and size.");
            var implicitFlow = Layout(source);
            var explicitFlow = Layout(source with { BodyElements = [.. source.BodyElements, new DocxParagraphElement(blank)] });
            TestAssert.Equal(explicitFlow.Pages.Count, implicitFlow.Pages.Count);
            var left = NoteLines(implicitFlow);
            var right = NoteLines(explicitFlow);
            TestAssert.Equal(81, left.Length);
            TestAssert.Equal(right.Length, left.Length);
            for (int i = 0; i < left.Length; i++)
            {
                TestAssert.Equal(right[i].Id, left[i].Id);
                TestAssert.Equal(right[i].Page, left[i].Page);
                TestAssert.True(Math.Abs(right[i].Baseline - left[i].Baseline) < .001d,
                    $"Implicit and ordinary paragraph flow agree for {family} {size}, {mode}, note {left[i].Id}.");
            }
        }
    }

    public static void DocxMissingFinalParagraphDefaultsRetainTheBoundedFrame()
    {
        var source = Read(OoxPdfDocxMarkupMode.Final).Source with { EmptyFinalSectionParagraph = null };
        var layout = Layout(source);
        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(2, FirstEndnotePage(layout));
        TestAssert.Equal(0, layout.Pages[1].Items.Count);
    }

    public static void DocxImplicitFinalParagraphRetainsStaticAndScaledFallbacks()
    {
        var source = Read(OoxPdfDocxMarkupMode.AllMarkup).Source;
        var defaults = source.EmptyFinalSectionParagraph ?? throw new InvalidOperationException("Resolve paragraph defaults.");
        foreach (string variant in new[] { "static", "scaled", "mirror", "alignment" })
        {
            var document = variant switch
            {
                "static" => source with { HeaderParagraphs = [DocxTests.CreateDocxLayoutParagraph("Public header", 12d, 14d)] },
                "mirror" => source with { Settings = source.Settings with { MirrorMargins = true } },
                "alignment" => source with { EmptyFinalSectionParagraph = defaults with { Alignment = DocxTextAlignment.Center } },
                _ => source
            };
            var geometry = variant == "scaled" ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            var expected = new DocxLayoutEngine(geometry, wordCompatiblePrintScale: .8d)
                .Create(document with { EmptyFinalSectionParagraph = null }, new DocxTests.FamilyWidthTextMeasurer(), default);
            var actual = new DocxLayoutEngine(geometry, wordCompatiblePrintScale: .8d)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.Equal(expected.Pages.Count, actual.Pages.Count);
            TestAssert.Equal(0, actual.Pages[1].Items.Count);
            var left = NoteLines(expected);
            var right = NoteLines(actual);
            TestAssert.Equal(left.Length, right.Length);
            for (int i = 0; i < left.Length; i++)
            {
                TestAssert.Equal(left[i].Id, right[i].Id);
                TestAssert.Equal(left[i].Page, right[i].Page);
                TestAssert.True(Math.Abs(left[i].Baseline - right[i].Baseline) < .001d,
                    "Unqualified paragraph/frame defaults retain the previous placement: " + variant);
            }
        }
    }

    public static void DocxEndnoteOriginConsumesTheFinalAutomaticLeadingOnce()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (double factor in new[] { 1d, 1.5d, 2d })
        {
            var source = Read(mode).Source;
            var paragraph = source.EmptyFinalSectionParagraph!;
            var single = source with { EmptyFinalSectionParagraph = paragraph with { LineSpacingFactor = 1d } };
            var multiple = source with { EmptyFinalSectionParagraph = paragraph with { LineSpacingFactor = factor } };
            var baseline = Layout(single);
            var actual = Layout(multiple);
            var mark = actual.Pages[1].Items.OfType<DocxTextLineLayout>().Single();
            var before = NoteLines(baseline).First();
            var after = NoteLines(actual).First();
            double leading = mark.LineHeight!.Value * (1d - 1d / factor);
            TestAssert.Equal(before.Page, after.Page);
            TestAssert.True(Math.Abs(before.Baseline - after.Baseline - leading) < .001d,
                "The note boundary consumes only the final automatic slot's added leading, independently of paragraph-after spacing.");
            TestAssert.Equal(81, NoteLines(actual).Length);
        }
    }

    public static void DocxEndnoteAutomaticLeadingRetainsTheScaledOriginFallback()
    {
        var source = Read(OoxPdfDocxMarkupMode.AllMarkup).Source;
        var paragraph = source.EmptyFinalSectionParagraph!;
        DocxLayout Create(double factor) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup,
            wordCompatiblePrintScale: .8d).Create(source with { BodyElements = [.. source.BodyElements,
                new DocxParagraphElement(paragraph with { LineSpacingFactor = factor })] }, new DocxTests.FamilyWidthTextMeasurer(), default);
        var single = NoteLines(Create(1d));
        var doubled = NoteLines(Create(2d));
        TestAssert.Equal(single.Length, doubled.Length);
        for (int i = 0; i < single.Length; i++)
        {
            TestAssert.Equal(single[i].Page, doubled[i].Page);
            TestAssert.True(Math.Abs(single[i].Baseline - doubled[i].Baseline) < .001d,
                "The unqualified scaled frame keeps its previous baseline origin.");
        }
    }

    private static (string? Id, int Page, double Baseline)[] NoteLines(DocxLayout layout) => layout.Pages
        .SelectMany((page, index) => page.PlacedRelatedStories.Where(note => note.StoryLayout.Story.IsNormalStoryType)
            .SelectMany(note => note.TextLines.Select(line => (note.StoryLayout.Story.Id, index + 1, line.BaselineY)))).ToArray();

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
