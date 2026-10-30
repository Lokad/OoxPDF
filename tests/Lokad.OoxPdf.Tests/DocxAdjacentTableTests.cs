using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxAdjacentTableTests
{
    private static readonly string Cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "AdjacentTables"));
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static void DocxAdjacentTablesMatchIndependentWordTopologyAndProperties()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(path, mode);
                IEnumerable<DocxBodyElement> elements = control.GetProperty("Container").GetString() switch
                {
                    "header" => document.HeaderBodyElementsByType["default"],
                    "cell" => document.Tables.Single().Rows.Single().Cells.Single().BodyElements,
                    _ => document.BodyElements
                };
                var tables = elements.OfType<DocxTableElement>().Select(element => element.Table).ToArray();
                var expectedTables = control.GetProperty("Tables").EnumerateArray().ToArray();
                TestAssert.Equal(expectedTables.Length, tables.Length);
                for (int tableIndex = 0; tableIndex < tables.Length; tableIndex++)
                {
                    var table = tables[tableIndex];
                    var expected = expectedTables[tableIndex];
                    TestAssert.Equal(expected.GetProperty("GridColumns").GetInt32(), table.ColumnWidthsPoints.Count);
                    var rows = expected.GetProperty("Rows").EnumerateArray().ToArray();
                    TestAssert.Equal(rows.Length, table.Rows.Count);
                    for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
                    {
                        var row = table.Rows[rowIndex];
                        var rowExpected = rows[rowIndex];
                        TestAssert.Equal(rowExpected.GetProperty("GridBefore").GetInt32(), row.GridBefore);
                        var spans = rowExpected.GetProperty("Spans").EnumerateArray().Select(value => value.GetInt32()).ToArray();
                        TestAssert.Equal(spans.Length, row.Cells.Count);
                        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                        {
                            var cell = row.Cells[cellIndex];
                            TestAssert.Equal(spans[cellIndex], cell.GridSpan);
                            TestAssert.Equal(rowExpected.GetProperty("LeftPadding").GetDouble(), cell.Margins.LeftPoints);
                            string? fill = control.GetProperty("Fills").GetProperty(cell.Text).GetString();
                            TestAssert.True(fill == cell.FillHex,
                                $"Word joined table {control.GetProperty("Id").GetString()}/{cell.Text}: expected {fill}, got {cell.FillHex}.");
                            string? color = rowExpected.GetProperty("BorderColor").GetString();
                            var borders = cell.Borders.Where(border => border.Value is not ("none" or "nil")).ToArray();
                            TestAssert.Equal(color is null ? 0 : 4, borders.Length);
                            foreach (var border in borders)
                            {
                                TestAssert.Equal(color, border.Color);
                                TestAssert.Equal(rowExpected.GetProperty("BorderSize").GetString(), border.SizeValue);
                            }
                        }
                    }
                }
            }
        }
    }

    public static void DocxAdjacentTablesRestoreOriginalLegacyMembership()
    {
        string folder = Path.GetFullPath(Path.Combine(Cases, "..", "TableLook"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "legacy-adjacent-office.json")));
        var control = oracle.RootElement.GetProperty("Controls")[0];
        string input = Path.Combine(folder, control.GetProperty("Input").GetString()!);
        TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var table = DocxTests.ReadDocx(input, mode).Tables.Single();
            TestAssert.Equal(13, table.Rows.Count);
            var cells = table.Rows.SelectMany(row => row.Cells).ToArray();
            TestAssert.Equal(22, cells.Length);
            foreach (var cell in cells)
                TestAssert.Equal(control.GetProperty("Fills").GetProperty(cell.Text).GetString(), cell.FillHex);
        }
    }

    public static void DocxAdjacentTablesKeepCachedSourceAndRevisionOwners()
    {
        using var stream = File.OpenRead(Path.Combine(Cases, "header.docx"));
        var package = OoxPackage.Open(stream, default);
        var headerPart = TestAssert.NotNull(package.GetPart("/word/header1.xml"));
        var header = package.LoadXml(headerPart, default);
        string before = header.ToString(SaveOptions.DisableFormatting);
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = new DocxReader().Read(package, null, default, mode);
            TestAssert.Equal(4, document.HeaderBodyElementsByType["default"].OfType<DocxTableElement>().Single().Table.Rows.Count);
            TestAssert.Equal(before, header.ToString(SaveOptions.DisableFormatting));
        }
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        var exception = TestAssert.Throws<OperationCanceledException>(() =>
            new DocxReader().Read(package, null, canceled.Token, OoxPdfDocxMarkupMode.Final));
        TestAssert.Equal(canceled.Token, exception.CancellationToken);
        TestAssert.Equal(before, header.ToString(SaveOptions.DisableFormatting));

        string changed = Transform("adjacent.docx", document =>
        {
            var tables = document.Descendants(W + "body").Single().Elements(W + "tbl").ToArray();
            tables[0].ReplaceWith(new XElement(W + "ins", new XAttribute(W + "id", "1"),
                new XAttribute(W + "author", "Table control"), new XElement(tables[0])));
            tables[1].ReplaceWith(new XElement(W + "del", new XAttribute(W + "id", "2"),
                new XAttribute(W + "author", "Table control"), new XElement(tables[1])));
        });
        try
        {
            TestAssert.Equal(2, DocxTests.ReadDocx(changed, OoxPdfDocxMarkupMode.Final).Tables.Single().Rows.Count);
            var joined = DocxTests.ReadDocx(changed, OoxPdfDocxMarkupMode.AllMarkup).Tables.Single();
            TestAssert.Equal(4, joined.Rows.Count);
            TestAssert.Equal(0, joined.Revisions.Count);
            TestAssert.True(joined.Rows.Take(2).All(row => row.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Insertion)),
                "Inserted fragment rows retain their own revision.");
            TestAssert.True(joined.Rows.Skip(2).All(row => row.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion)),
                "Deleted fragment rows retain their own revision.");
        }
        finally { File.Delete(changed); }
    }

    public static void DocxAdjacentTablesKeepDerivedGridBoundsAndCleanFailure()
    {
        string changed = Transform("adjacent.docx", document =>
        {
            int index = 0;
            foreach (var grid in document.Descendants(W + "tblGrid"))
            {
                grid.ReplaceNodes(Enumerable.Range(0, DocxReader.MaxInferredGridColumns).Select(_ =>
                    new XElement(W + "gridCol", new XAttribute(W + "w", index == 0 ? "20" : "30"))));
                index++;
            }
        });
        try
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxTests.ReadDocx(changed, OoxPdfDocxMarkupMode.Final));
        }
        finally { File.Delete(changed); }
    }

    private static string Transform(string name, Action<XDocument> transform)
    {
        var parts = new Dictionary<string, string>();
        using (var archive = ZipFile.OpenRead(Path.Combine(Cases, name)))
            foreach (var entry in archive.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                parts[entry.FullName] = reader.ReadToEnd();
            }
        var document = XDocument.Parse(parts["word/document.xml"]);
        transform(document);
        parts["word/document.xml"] = document.ToString();
        return TestFixtures.WriteTempPackage(".docx", parts);
    }
}
