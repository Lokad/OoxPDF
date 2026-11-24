using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
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
            AssertOverlapPlan(row, next, cols);
        }
    }

    public static void SharedBorderPlanMatchesNestedEnumerationMerged()
    {
        (DocxTableRowLayout row, DocxTableRowLayout next) = LayoutMergedPair();
        AssertOverlapPlan(row, next, 2);
    }

    private static void AssertOverlapPlan(DocxTableRowLayout row, DocxTableRowLayout next, int count)
    {
        DocxRenderer.RowPairBorderPlan plan = TestAssert.NotNull(
            DocxRenderer.RowPairBorderPlan.TryBuild(row, next, CancellationToken.None), "Consecutive rows must produce a plan.");
        (double X, double Right)[] expected = BruteForcePairs(row, next);
        TestAssert.Equal(count, expected.Length);
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
        var strips = ReadFills(pdf).Where(fill => fill.Color == "0 0 1").ToArray();
        TestAssert.Equal(4, strips.Length);
        for (int index = 0; index < strips.Length; index++)
        {
            var strip = strips[index];
            TestAssert.Equal(36d + index * 72d, strip.X);
            TestAssert.Equal(72d, strip.Width);
            TestAssert.True(Math.Abs(strip.Y + strip.Height - 708d) < .1d && strip.Height is > 0d and < 2d,
                "The filled strips meet the authored boundary one 48pt row below the top margin.");
        }
    }

    public static void DocxMergedSharedBordersRetainEachNeighborRow()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string?[] colors = ["FF0000", "00FF00", "0000FF"];
            var plain = ReadFills(RenderBorderNeighbors(mode, merge: false, colors));
            var merged = ReadFills(RenderBorderNeighbors(mode, merge: true, colors));
            TestAssert.Equal(3, plain.Length);
            TestAssert.Equal(3, merged.Length);
            for (int index = 0; index < merged.Length; index++)
            {
                TestAssert.Equal(plain[index], merged[index]);
                TestAssert.Equal(48d, merged[index].Height);
            }
            TestAssert.Equal("1 0 0", merged[0].Color);
            TestAssert.Equal("0 1 0", merged[1].Color);
            TestAssert.Equal("0 0 1", merged[2].Color);
            TestAssert.Equal(merged[0].Y, merged[1].Y + merged[1].Height);
            TestAssert.Equal(merged[1].Y, merged[2].Y + merged[2].Height);
            TestAssert.Equal(plain[2], ReadFills(RenderBorderNeighbors(mode, merge: true, [null, null, "0000FF"])).Single());

            var aligned = ReadFills(RenderBorderNeighbors(mode, merge: true, ["FF0000", null, null], rightMergeStart: 0)).Single();
            TestAssert.Equal(("1 0 0", plain[0].X, plain[2].Y, plain[0].Width, 144d), aligned);
            var staggered = ReadFills(RenderBorderNeighbors(mode, merge: true, ["FF0000", "00FF00", null], rightMergeStart: 1));
            TestAssert.Equal(3, staggered.Length);
            TestAssert.Equal(plain[0], staggered[0]);
            TestAssert.Equal(plain[1], staggered[1]);
            TestAssert.Equal(("0 1 0", plain[2].X, plain[2].Y, plain[2].Width, 48d), staggered[2]);
        }
    }

    private static string RenderBorderNeighbors(OoxPdfDocxMarkupMode mode, bool merge, string?[] colors, int? rightMergeStart = null)
    {
        string rows = string.Concat(colors.Select((color, index) => $$"""
            <w:tr>
              <w:trPr><w:trHeight w:val="960" w:hRule="exact"/></w:trPr>
              <w:tc><w:tcPr>{{Merge(index, merge ? 0 : null)}}</w:tcPr><w:p/></w:tc>
              <w:tc><w:tcPr>{{Merge(index, rightMergeStart)}}{{(color is null ? "" : $"<w:tcBorders><w:left w:val=\"single\" w:sz=\"24\" w:color=\"{color}\"/></w:tcBorders>")}}</w:tcPr><w:p/></w:tc>
            </w:tr>
            """));
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = TestFixtures.DocxContentTypes,
            ["_rels/.rels"] = TestFixtures.DocxPackageRelationships,
            ["word/document.xml"] = $$"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
                  <w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
                    <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>{{rows}}</w:tbl>
                  <w:sectPr><w:pgSz w:w="8640" w:h="6000"/>
                    <w:pgMar w:top="480" w:bottom="480" w:left="720" w:right="720"/></w:sectPr>
                </w:body></w:document>
                """
        });
        using var package = OoxPackage.Open(input, default);
        var document = new DocxReader().Read(package, null, default, markupMode: mode);
        var geometry = mode == OoxPdfDocxMarkupMode.Final
            ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        return new DocxRenderer(new TestFaceFontResolver(), mode, geometry)
            .RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single().Content;

        static string Merge(int row, int? start) => row == start ? "<w:vMerge w:val=\"restart\"/>" :
            row > start ? "<w:vMerge/>" : string.Empty;
    }

    private static (string Color, double X, double Y, double Width, double Height)[] ReadFills(string content)
    {
        return Regex.Matches(content, @"(?m)^(?<rgb>[01] [01] [01]) rg\r?\n(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f\r?$")
            .Select(m => (m.Groups["rgb"].Value, Number(m, "x"), Number(m, "y"), Number(m, "w"), Number(m, "h")))
            .ToArray();

        static double Number(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
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
            cells[i] = new DocxTableCell([paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
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
            [new DocxTableElement(table)]);
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
        const string rowOpen = "<w:tr><w:trPr><w:trHeight w:val=\"960\" w:hRule=\"exact\"/></w:trPr>";
        body.Append(rowOpen).Append(firstRow).Append("</w:tr>");
        for (int r = 1; r < rows; r++)
        {
            body.Append(rowOpen).Append(otherRow).Append("</w:tr>");
        }

        body.Append("</w:tbl><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"720\" w:bottom=\"720\" w:left=\"720\" w:right=\"720\"/></w:sectPr>");
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
