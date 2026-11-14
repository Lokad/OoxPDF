using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxAutomaticReviewRowTests
{
    public static void DocxExplicitAutomaticRowsSharePrintedFragmentFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,
                     OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        foreach (int count in new[] { 24, 40 })
        foreach (bool prefix in new[] { false, true })
        foreach (bool? widow in new bool?[] { null, false, true })
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph(
                string.Join('\n', Enumerable.Range(0, count).Select(index => $"ROW{index:000}")), 10d, 12d,
                DocxParagraphKeepRules.Empty with { WidowControl = widow }) with { LineSpacingPoints = null };
            var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [],
                DocxTableCellMargins.Empty with { LeftPoints = 0d, RightPoints = 0d, TopPoints = 0d, BottomPoints = 0d });
            var implicitLayout = Layout(null, null);
            var expected = Lines(implicitLayout);
            TestAssert.Equal(count, expected.Length);
            foreach (double value in new[] { 0d, 320d })
            {
                var explicitLayout = Layout(value, "auto");
                TestAssert.Equal(implicitLayout.Pages.Count, explicitLayout.Pages.Count);
                TestAssert.True(expected.SequenceEqual(Lines(explicitLayout)),
                    "Explicit automatic height must preserve each source line's printed page, frame and baseline, including default/explicit widow policy.");
            }

            DocxLayout Layout(double? height, string? rule)
            {
                var row = new DocxTableRow([cell], height) { HeightRuleValue = rule };
                var table = new DocxTable("fixed", [140d], [row]);
                IReadOnlyList<DocxBodyElement> elements = prefix
                    ? [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d)), new DocxTableElement(table)]
                    : [new DocxTableElement(table)];
                var document = DocxTests.CreateLayoutTestDocument(elements, [table]);
                return new DocxLayoutEngine(mode, wordCompatiblePrintScale: .75d)
                    .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
            }

            static (int Page, string Text, double X, double Y, double Width, double Height)[] Lines(DocxLayout layout) =>
                layout.Pages.SelectMany((page, index) => page.Items.OfType<DocxTableRowLayout>()
                    .SelectMany(row => row.Cells).SelectMany(c => c.TextLines)
                    .Select(line => (index, line.Text, line.X, line.BaselineY, line.Width, line.LineHeight))).ToArray();
        }
    }
}
