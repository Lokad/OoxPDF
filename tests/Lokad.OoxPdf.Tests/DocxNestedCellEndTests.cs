using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNestedCellEndTests
{
    public static void DocxNestedCellEndMarkHasNoOccupiedHeightAndRetainsItsSource()
    {
        foreach (var mode in Modes)
        foreach (double size in new[] { 12d, 36d })
        {
            var mark = Paragraph(string.Empty, size) with
            { BookmarkAnchors = [new("5", "CellEnd", 0, 0, 0)] };
            var nested = Nested();
            var table = Outer([new DocxTableElement(nested), new DocxParagraphElement(mark)]);
            var placed = Layout(table, mode).Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(20d, placed.Height);
            var carrier = placed.Cells[0].TextLines.Single();
            TestAssert.True(ReferenceEquals(mark, carrier.SourceParagraph),
                "A structural mark keeps the actual source paragraph and bookmark owner.");
            TestAssert.Equal("CellEnd", carrier.SourceParagraph!.BookmarkAnchors.Single().Name);
            TestAssert.Equal(new DocxOccupiedLineBox(0d, 0d), carrier.OccupiedLineBox);
            TestAssert.Equal(0d, carrier.AppliedBeforeSpacing);
            TestAssert.Equal(0d, carrier.PendingAfterSpacing);
            TestAssert.True(carrier.Segments.Count > 0 && placed.Cells[0].NestedRows.Count == 1,
                "Keep nominal navigation geometry and the real child table.");
            TestAssert.True(DocxTableCellClipGeometry.HasCompleteMeasuredRows([placed], default),
                "A retained nonpainting source owner cannot make a complete cell frame appear clipped or unknown.");
        }
    }

    public static void DocxNestedCellEndMarkKeepsVisibleLeadingAndReferenceFrames()
    {
        var empty = Paragraph(string.Empty);
        foreach (var mode in Modes)
        {
            var nested = new DocxTableElement(Nested());
            foreach (var (elements, expected) in new (DocxBodyElement[], double)[]
            {
                ([nested, new DocxParagraphElement(empty)], 20d),
                ([nested, new DocxParagraphElement(Paragraph("Tail"))], 40d),
                ([nested, new DocxParagraphElement(Paragraph(" "))], 40d),
                ([new DocxParagraphElement(empty), nested, new DocxParagraphElement(empty)], 40d),
                ([nested, new DocxParagraphElement(empty), new DocxParagraphElement(Paragraph("End"))], 60d),
                ([new DocxParagraphElement(empty)], 20d),
                ([nested, new DocxParagraphElement(empty with
                    { InlineReferences = [new(DocxRelatedStoryKind.Comment, "0", null, null, 0, 0, 0)] })], 40d),
            })
            {
                var row = Layout(Outer(elements), mode).Pages[0].Items.OfType<DocxTableRowLayout>().Single();
                TestAssert.Equal(expected, row.Height);
            }
        }
    }

    public static void DocxNestedCellEndMarkUsesTheSameFrameForPaginationAndPlacement()
    {
        foreach (var mode in Modes)
        {
            var mark = Paragraph(string.Empty);
            var table = Outer([new DocxTableElement(Nested()), new DocxParagraphElement(mark)]);
            var rows = new List<DocxTableRow> { table.Rows[0] with { IsHeader = true, CantSplit = true } };
            for (int i = 0; i < 7; i++) rows.Add(new([Cell(Paragraph("ROW" + i))], 40d) { CantSplit = true });
            table = table with { Rows = rows };
            var after = Paragraph("After");
            var document = DocxTests.CreateLayoutTestDocument(
                [new DocxTableElement(table), new DocxParagraphElement(after)], [table]);
            var layout = Engine(mode).Create(document, new Measurer(), default);
            TestAssert.Equal(2, layout.Pages.Count);
            TestAssert.True(layout.Pages[^1].Items.OfType<DocxTextLineLayout>().Any(line => ReferenceEquals(after, line.SourceParagraph)),
                "The real following paragraph remains; no final-page removal is involved.");
            TestAssert.Equal(7, layout.Pages.Sum(p => p.Items.OfType<DocxTableRowLayout>().Count(r => !r.IsHeader)));
            TestAssert.True(layout.Pages.All(p => p.Items.OfType<DocxTableRowLayout>().First().Height == 20d),
                "Repeated admission and actual placement share the child frame.");
        }
    }

    public static void DocxNestedCellEndMarkKeepsItsPlacedInsertionFrame()
    {
        foreach (var mode in Modes)
        foreach (double size in new[] { 12d, 36d })
        foreach (double before in new[] { 0d, 20d })
        {
            var mark = Paragraph(string.Empty, size) with { SpacingBeforePoints = before, SpacingAfterPoints = 20d };
            var table = Outer([new DocxTableElement(Nested()), new DocxParagraphElement(mark)]);
            var placed = Layout(table, mode).Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            var carrier = placed.Cells[0].TextLines.Single();
            TestAssert.True(carrier.FirstLineInsetPoints is { } inset && double.IsFinite(inset),
                "The zero-occupancy source owner publishes its placed insertion frame independently of glyph metrics.");
            TestAssert.Equal(170d, carrier.BaselineY + carrier.FirstLineInsetPoints!.Value);
            TestAssert.Equal(20d, placed.Height);
            TestAssert.Equal(new DocxOccupiedLineBox(0d, 0d), carrier.OccupiedLineBox);
        }
    }

    public static void DocxNestedCellEndMarkRetainsItsSourceFrameThroughReviewProjection()
    {
        foreach (double size in new[] { 12d, 36d })
        {
            var mark = Paragraph(string.Empty, size) with { LineSpacingPoints = null };
            var childParagraph = Paragraph("Child") with { LineSpacingPoints = null };
            var child = new DocxTable("fixed", [80d], [new([Cell(childParagraph)], null)]);
            var table = Outer([new DocxTableElement(child), new DocxParagraphElement(mark)]);
            var source = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            var measurer = new Measurer(naturalMetrics: true);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, .75d)
                .Create(source, measurer, default);
            var placed = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(.75d, placed.ReviewBorderPrintScale);
            var carrier = placed.Cells[0].TextLines.Single();
            TestAssert.True(carrier.FirstLineInsetPoints is not null,
                "Review projection retains the source frame even though the mark paints nothing.");
            TestAssert.True(Math.Abs(carrier.BaselineY + carrier.FirstLineInsetPoints!.Value - placed.Y) < .000001d,
                "The insertion frame follows the containing row; an additional glyph-baseline projection must not move it.");
        }
    }

    public static void DocxNestedCellEndBookmarkUsesThePlacedFrameWithoutGlyphAscent()
    {
        foreach (var mode in Modes)
        foreach (double size in new[] { 12d, 36d })
        {
            var mark = Paragraph(string.Empty, size);
            var run = mark.Runs[0] with { SourceRunIndex = 0 };
            mark = mark with { Runs = [run], InlineContent = [],
                BookmarkAnchors = [new("5", "CellEnd", 0, 0, 0) { InlineItemIndex = 0 }] };
            var table = Outer([new DocxTableElement(Nested()), new DocxParagraphElement(mark)]);
            var jump = Paragraph("Jump");
            jump = jump with { Runs = [jump.Runs[0] with { SourceRunIndex = 0 }],
                Hyperlinks = [new(null, "CellEnd", null, null, null, null, null, 0, 1, 0, 1, 4)] };
            var source = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table), new DocxParagraphElement(jump)], [table]);
            var pages = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, mode)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var destination = pages.SelectMany(page => page.Annotations).Single(annotation => annotation.Destination is not null)
                .Destination!.Value;
            TestAssert.Equal(0, destination.PageIndex);
            TestAssert.True(Math.Abs(destination.Top!.Value - 170d) < .001d,
                "A bookmark on a source-only cell mark uses the placed cell end, independent of the empty run's font size.");
        }
    }

    private static readonly OoxPdfDocxMarkupGeometryMode[] Modes =
        [OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup];
    private static DocxLayoutEngine Engine(OoxPdfDocxMarkupGeometryMode mode) => new(mode, wordCompatiblePrintScale: 1d);
    private static DocxLayout Layout(DocxTable table, OoxPdfDocxMarkupGeometryMode mode) =>
        Engine(mode).Create(DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]), new Measurer(), default);
    private sealed class Measurer(bool naturalMetrics = false) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double size) => text.Length * 5d;
        public double MeasureWindowsAscender(DocxTextRun? run, double size) => naturalMetrics ? size : 8d;
        public double MeasureWindowsDescender(DocxTextRun? run, double size) => naturalMetrics ? size / 4d : 2d;
    }
    private static DocxParagraph Paragraph(string text, double size = 12d) => DocxTests.CreateDocxLayoutParagraph(text, size, 20d);
    private static DocxTableCell Cell(DocxParagraph paragraph) => new(paragraph.Runs[0].Text, [paragraph], null, null, null, null, [],
        DocxTableCellMargins.Empty with { TopPoints = 0d, BottomPoints = 0d, LeftPoints = 0d, RightPoints = 0d });
    private static DocxTable Nested() => new("fixed", [80d], [new([Cell(Paragraph("Child"))], 20d)]);
    private static DocxTable Outer(IReadOnlyList<DocxBodyElement> elements) => new("fixed", [80d],
        [new([Cell(Paragraph(string.Empty)) with { BodyElements = elements }], null)]);
}
