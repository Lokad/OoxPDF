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

    private static readonly OoxPdfDocxMarkupGeometryMode[] Modes =
        [OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup];
    private static DocxLayoutEngine Engine(OoxPdfDocxMarkupGeometryMode mode) => new(mode, wordCompatiblePrintScale: 1d);
    private static DocxLayout Layout(DocxTable table, OoxPdfDocxMarkupGeometryMode mode) =>
        Engine(mode).Create(DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]), new Measurer(), default);
    private sealed class Measurer : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double size) => text.Length * 5d;
        public double MeasureWindowsAscender(DocxTextRun? run, double size) => 8d;
        public double MeasureWindowsDescender(DocxTextRun? run, double size) => 2d;
    }
    private static DocxParagraph Paragraph(string text, double size = 12d) => DocxTests.CreateDocxLayoutParagraph(text, size, 20d);
    private static DocxTableCell Cell(DocxParagraph paragraph) => new(paragraph.Runs[0].Text, [paragraph], null, null, null, null, [],
        DocxTableCellMargins.Empty with { TopPoints = 0d, BottomPoints = 0d, LeftPoints = 0d, RightPoints = 0d });
    private static DocxTable Nested() => new("fixed", [80d], [new([Cell(Paragraph("Child"))], 20d)]);
    private static DocxTable Outer(IReadOnlyList<DocxBodyElement> elements) => new("fixed", [80d],
        [new([Cell(Paragraph(string.Empty)) with { BodyElements = elements }], null)]);
}
