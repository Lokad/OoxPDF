using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStaticTableAuthoredMarginsKeepContainerGrid()
    {
        foreach (double margin in new[] { 0d, 5.4d, 10d, 18d })
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph("Public cell", 11d, 16d);
            var cell = new DocxTableCell("Public cell", [paragraph], null, null, null, null, [],
                DocxTableCellMargins.Empty with { LeftPoints = margin, RightPoints = margin });
            var table = new DocxTable("fixed", [200d], [new DocxTableRow([cell], null)]);
            var page = LayoutSharedTable(table);
            var flowing = page.Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            var stationary = page.StaticTableRows.Single().Cells.Single();
            TestAssert.True(Math.Abs(flowing.X - (page.MarginLeft - margin)) < 1e-9,
                "Flowing tables anchor the first text edge using their authored margin.");
            TestAssert.True(Math.Abs(stationary.X - page.MarginLeft) < 1e-9,
                "A header/footer table anchors its shared grid at the container origin.");
            TestAssert.True(Math.Abs(stationary.TextLines.Single().X - flowing.TextLines.Single().X - margin) < 1e-9,
                "Cell content and its frame carry the same authored inset exactly once.");
            TestAssert.Equal(flowing.Width, stationary.Width);
        }
    }

    public static void DocxStaticTableOriginKeepsIndentStyleAndCompatibilityFallbacks()
    {
        var paragraph = DocxTests.CreateDocxLayoutParagraph("Public cell", 11d, 16d);
        var cell = new DocxTableCell("Public cell", [paragraph], null, null, null, null, [],
            DocxTableCellMargins.Empty with { LeftPoints = 10d, RightPoints = 10d });
        var basis = new DocxTable("fixed", [200d], [new DocxTableRow([cell], null)]);
        foreach (var table in new[]
        {
            basis with { IndentPoints = 18d },
            basis with { IndentPoints = -18d },
            basis with { UseLegacyTableGrid = true },
            basis with { Rows = [new DocxTableRow([cell with
                { StyleMargins = DocxTableCellMargins.Empty with { LeftPoints = 12d } }], null)] }
        })
        {
            var page = LayoutSharedTable(table);
            var flowing = page.Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            var stationary = page.StaticTableRows.Single().Cells.Single();
            TestAssert.Equal(flowing.X, stationary.X);
            TestAssert.Equal(flowing.Width, stationary.Width);
            TestAssert.Equal(flowing.TextLines.Single().X, stationary.TextLines.Single().X);
        }
    }

    public static void DocxStaticTableOriginsMatchIndependentWordContentAndGeometry()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent static-table origins use installed Windows fonts.");
        const string golden = "static-table-origins-office.json";
        CheckInlineBoxMath(null, golden);
        CheckInlineBoxWordOrigins(golden);
        CheckInlineBoxProseOrigins(golden);
        CheckInlineBoxFollowingOrigins(golden);
    }

    private static DocxLayoutPage LayoutSharedTable(DocxTable table)
    {
        var body = DocxTests.CreateDocxLayoutParagraph("Public body", 11d, 16d);
        var settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>
                { ["default"] = [new DocxTableElement(table)] }
        };
        var document = new DocxDocument(500d, 700d, 36d, 36d, 100d, 36d, settings, [], [], [],
            [new DocxParagraphElement(body), new DocxTableElement(table)], [body], [table]);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default).Pages.Single();
    }
}
