using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxBorderPlanTests
{
    public static void SharedBorderPlanMatchesNestedEnumerationAligned()
    {
        foreach (int cols in new[] { 8, 16, 32 })
        {
            (DocxTableRowLayout row, DocxTableRowLayout next) = LayoutAlignedPair(cols);
            DocxRenderer.RowPairBorderPlan? plan = DocxRenderer.RowPairBorderPlan.TryBuild(row, next, CancellationToken.None);
            TestAssert.True(plan is not null, "Aligned consecutive rows must produce a plan.");
            TestAssert.Equal(cols, plan.Overlaps.Length);
            (double X, double Right)[] expected = BruteForcePairs(row, next);
            TestAssert.Equal(expected.Length, plan.Overlaps.Length);
            for (int k = 0; k < expected.Length; k++)
            {
                TestAssert.Equal(expected[k].X, plan.Overlaps[k].X);
                TestAssert.Equal(expected[k].Right, plan.Overlaps[k].Right);
            }
        }
    }

    public static void SharedBorderPlanMatchesNestedEnumerationMerged()
    {
        (DocxTableRowLayout row, DocxTableRowLayout next) = LayoutMergedPair();
        TestAssert.Equal(1, row.Cells.Count);
        TestAssert.Equal(2, next.Cells.Count);
        DocxRenderer.RowPairBorderPlan? plan = DocxRenderer.RowPairBorderPlan.TryBuild(row, next, CancellationToken.None);
        TestAssert.True(plan is not null, "Merged consecutive rows must produce a plan.");
        (double X, double Right)[] expected = BruteForcePairs(row, next);
        TestAssert.Equal(2, expected.Length);
        TestAssert.Equal(expected.Length, plan.Overlaps.Length);
        for (int k = 0; k < expected.Length; k++)
        {
            TestAssert.Equal(expected[k].X, plan.Overlaps[k].X);
            TestAssert.Equal(expected[k].Right, plan.Overlaps[k].Right);
        }
    }

    public static void WideAlignedTableRendersOneStripPerOverlap()
    {
        string input = WriteBorderTable(columns: 4, rows: 2);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(4, pdf.Split("0 0 1 rg", StringSplitOptions.None).Length - 1);
    }

    private static (DocxTableRowLayout Row, DocxTableRowLayout Next) LayoutAlignedPair(int cols)
    {
        double[] widths = Enumerable.Repeat(20d, cols).ToArray();
        DocxTableRow top = MakeRow(cols, "top");
        DocxTableRow bottom = MakeRow(cols, "bottom");
        var table = new DocxTable(null, widths, [top, bottom]);
        DocxLayout layout = LayoutTable(table);
        DocxTableRowLayout[] rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Where(r => r.RowIndex is 0 or 1).ToArray();
        TestAssert.Equal(2, rows.Length);
        DocxTableRowLayout row = rows.Single(r => r.RowIndex == 0);
        DocxTableRowLayout next = rows.Single(r => r.RowIndex == 1);
        TestAssert.Equal(cols, row.Cells.Count);
        TestAssert.Equal(cols, next.Cells.Count);
        AssertAllOwnCells(row);
        AssertAllOwnCells(next);
        return (row, next);
    }

    private static (DocxTableRowLayout Row, DocxTableRowLayout Next) LayoutMergedPair()
    {
        string input = WriteBorderTable(columns: 2, rows: 2, spanFirstCell: true);
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Where(r => r.RowIndex is 0 or 1).ToArray();
        TestAssert.Equal(2, rows.Length);
        DocxTableRowLayout row = rows.Single(r => r.RowIndex == 0);
        DocxTableRowLayout next = rows.Single(r => r.RowIndex == 1);
        TestAssert.Equal(1, row.Cells.Count);
        TestAssert.Equal(2, next.Cells.Count);
        AssertAllOwnCells(row);
        AssertAllOwnCells(next);
        return (row, next);
    }

    private static (double X, double Right)[] BruteForcePairs(DocxTableRowLayout row, DocxTableRowLayout next)
    {
        var pairs = new List<(double X, double Right)>();
        foreach (DocxTableCellLayout current in row.Cells)
        {
            foreach (DocxTableCellLayout other in next.Cells)
            {
                double x = Math.Max(current.X, other.X);
                double right = Math.Min(current.X + current.Width, other.X + other.Width);
                if (right > x)
                {
                    pairs.Add((x, right));
                }
            }
        }

        return pairs.ToArray();
    }

    private static void AssertAllOwnCells(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            TestAssert.True(cell.VisualOwnership == DocxTableCellVisualOwnership.OwnCell, "Plan equivalence fixtures must avoid vertical merges so visibility is trivially true.");
        }
    }

    private static DocxTableRow MakeRow(int cols, string prefix)
    {
        var cells = new DocxTableCell[cols];
        for (int i = 0; i < cols; i++)
        {
            DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph(prefix + i.ToString(System.Globalization.CultureInfo.InvariantCulture), 10d, 10d);
            cells[i] = new DocxTableCell("c", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        }

        return new DocxTableRow(cells, 60d);
    }

    private static DocxLayout LayoutTable(DocxTable table)
    {
        var document = new DocxDocument(
            2000d,
            2000d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table]);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
    }

    private static string WriteBorderTable(int columns, int rows, bool spanFirstCell = false)
    {
        var grid = new StringBuilder();
        var firstRow = new StringBuilder();
        var otherRow = new StringBuilder();
        for (int i = 0; i < columns; i++)
        {
            grid.Append("<w:gridCol w:w=\"1440\"/>");
            otherRow.Append("<w:tc><w:p/></w:tc>");
            if (spanFirstCell && i == 0)
            {
                firstRow.Append("<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p/></w:tc>");
                i++;
                otherRow.Append("<w:tc><w:p/></w:tc>");
            }
            else
            {
                firstRow.Append("<w:tc><w:p/></w:tc>");
            }
        }

        var body = new StringBuilder();
        body.Append("<w:tbl><w:tblPr><w:tblBorders><w:insideH w:val=\"single\" w:color=\"0000FF\" w:sz=\"8\"/></w:tblBorders></w:tblPr>");
        body.Append("<w:tblGrid>").Append(grid).Append("</w:tblGrid>");
        body.Append("<w:tr>").Append(firstRow).Append("</w:tr>");
        for (int r = 1; r < rows; r++)
        {
            body.Append("<w:tr>").Append(otherRow).Append("</w:tr>");
        }

        body.Append("</w:tbl><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>");
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                """ + body + """
                  </w:body>
                </w:document>
                """,
        });
    }
}
