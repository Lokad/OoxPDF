using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCellPageNotesPublishOneOwnedParagraphAcrossViews()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/cell-page-notes-" + typography, mode, false, false);
            var before = DocxBlockTraversal.EnumerateBodyParagraphs(source).ToArray();
            TestAssert.Equal(4, before.Length);
            TestAssert.True(!DocxPageFootnoteNumbering.IsEligibleDocument(source), "Typography retains its plain-body scope.");
            var equations = before.SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var first = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out var numbered, out var owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "Complete measured cell marker/note owners can restart independently of typography.");
            TestAssert.Equal("1,1,2,2", string.Join(',', new[] { "37", "4", "73", "8" }.Select(id => owners[id])));
            var after = DocxBlockTraversal.EnumerateBodyParagraphs(numbered).ToArray();
            TestAssert.Equal("09,10,09,10", string.Join(',', after.SelectMany(paragraph => paragraph.InlineReferences).Select(reference => reference.DisplayText)));
            TestAssert.Equal("105,106,107,108", string.Join(',', before.SelectMany(paragraph => paragraph.InlineReferences).Select(reference => reference.DisplayText)));
            foreach (var table in numbered.Tables)
            {
                var cell = table.Rows.Single().Cells.Single();
                var paragraph = cell.Paragraphs.Single();
                TestAssert.True(ReferenceEquals(paragraph, cell.BodyElements.OfType<DocxParagraphElement>().Single().Paragraph),
                    "Ordered and fallback cell views publish the same replacement paragraph.");
                var reference = paragraph.InlineReferences.Single();
                var run = TestAssert.NotNull(reference.DisplayRun);
                TestAssert.True(paragraph.Runs.Any(item => ReferenceEquals(item, run)) &&
                    paragraph.InlineContent!.OfType<DocxInlineText>().Any(item => ReferenceEquals(item.Run, run) && item.Length == run.Text.Length),
                    "References, source runs and ordered text share one new marker owner.");
                TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
                TestAssert.Equal(1, paragraph.Images.Count);
                TestAssert.True(paragraph.BookmarkAnchors.Count != 0 && cell.Text.Contains(reference.DisplayText!, StringComparison.Ordinal),
                    "Publication retains composition metadata and refreshes derived cell text.");
            }
            for (int index = 0; index < before.Length; index++)
            {
                int marker = before[index].InlineReferences.Single().SourceRunIndex;
                TestAssert.True(before[index].Runs.Where(run => run.SourceRunIndex != marker).All(run =>
                    after[index].Runs.Any(retained => ReferenceEquals(run, retained))), "Authored runs keep their original identities.");
                TestAssert.True(ReferenceEquals(before[index].Images.Single(), after[index].Images.Single()), "Pictures keep their source identity.");
            }
            var final = engine.Create(numbered, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
            TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Cell source/note pages remain stable after the bounded retry.");
            TestAssert.True(!DocxPageFootnoteNumbering.HasStablePages(numbered, final, owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Font purposes do not gain cell admission.");
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(numbered, final, out var unchanged, out _,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(numbered, unchanged),
                "Already correct labels do not trigger another font/layout pass.");
        }
    }

    public static void DocxInlineCompositionCellPageNotesRejectMissingAmbiguousAndClippedGlyphs()
    {
        var source = ReadMutation("CompositionProbes/cell-page-notes-calibri-12-zero", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(source).ToArray();
        var equations = paragraphs.SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
        var first = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
        TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "The unmodified measured cell control qualifies.");
        var row = first.Pages[0].Items.OfType<DocxTableRowLayout>().First();
        var cell = row.Cells.Single();
        var reference = paragraphs[0].InlineReferences.Single();
        var line = cell.TextLines.Single(line => line.Segments.Any(segment => segment.SourceTextRunIndex == reference.SourceRunIndex));
        var marker = line.Segments.Single(segment => segment.SourceTextRunIndex == reference.SourceRunIndex);
        foreach (string variant in new[] { "missing", "duplicate", "horizontal", "vertical", "unmeasured", "merge" })
        {
            var updated = variant switch {
                "missing" => line with { Segments = line.Segments.Where(segment => !ReferenceEquals(segment, marker)).ToArray() },
                "duplicate" => line with { Segments = [.. line.Segments, marker] },
                "horizontal" => line with { Segments = line.Segments.Select(segment => ReferenceEquals(segment, marker)
                    ? marker with { X = cell.X + cell.Width + 1d } : segment).ToArray() },
                "vertical" => line with { BaselineY = cell.Y - 1d },
                "unmeasured" => line with { InlineFrame = null, BodyLineBoxBaselineInsetPoints = null, BodyLineBoxHeightPoints = null },
                _ => line
            };
            var changedCell = cell with { TextLines = cell.TextLines.Select(item => ReferenceEquals(item, line) ? updated : item).ToArray(),
                VisualOwnership = variant == "merge" ? DocxTableCellVisualOwnership.VerticalMergeOwner : cell.VisualOwnership };
            var changedRow = row with { Cells = [changedCell] };
            var pages = first.Pages.ToArray();
            pages[0] = pages[0] with { Items = pages[0].Items.Select(item => ReferenceEquals(item, row) ? changedRow : item).ToArray() };
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first with { Pages = pages }, out var fallback, out _,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(source, fallback),
                "Incomplete or unpainted cell glyphs reject numbering atomically: " + variant);
        }
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellation.Token));
    }

    public static void DocxInlineCompositionCellPageNotesKeepIndependentLabelsMathAndImages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent cell-note controls use installed Windows fonts.");
        CheckPageNoteMath(null, "cell-page-notes-office.json");
    }

    public static void DocxInlineCompositionCellPageNotesRetainCustomSourceAndTargetIdentities()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        {
            var source = ReadMutation("CompositionProbes/cell-page-notes-mixed-custom-" + typography,
                OoxPdfDocxMarkupMode.Final, false, false);
            var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(source).ToArray();
            var custom = paragraphs.SelectMany(paragraph => paragraph.InlineReferences).Single(reference => reference.CustomMarkText is not null);
            var manualTable = source.Tables.Single(table => DocxBlockTraversal.EnumerateTableParagraphs(table)
                .Any(paragraph => paragraph.InlineReferences.Contains(custom)));
            var target = source.RelatedStories.Single(story => story.Id == custom.Id && story.IsNormalStoryType);
            var equations = paragraphs.SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var first = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out var numbered, out var owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "Complete cell markers retain manual ownership while automatic references consume page counters.");
            TestAssert.Equal("*,09,09,10", string.Join(',', DocxBlockTraversal.EnumerateBodyParagraphs(numbered)
                .SelectMany(paragraph => paragraph.InlineReferences).Select(reference => reference.DisplayText ?? reference.CustomMarkText)));
            TestAssert.True(numbered.Tables.Any(table => ReferenceEquals(table, manualTable)),
                "A table containing only an unchanged custom reference retains its container identity.");
            TestAssert.True(ReferenceEquals(target.Paragraphs[0], numbered.RelatedStories.Single(story => story.Id == custom.Id && story.IsNormalStoryType).Paragraphs[0]),
                "An authored target mark is not replaced with automatic text.");
            var final = engine.Create(numbered, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
            TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Manual and automatic cell glyph pages remain stable together.");
        }
    }
}
