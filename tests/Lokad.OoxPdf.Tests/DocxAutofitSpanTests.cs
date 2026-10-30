using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxAutofitSpanTests
{
    private static DocxTableCell Cell(string text, int span) => new(
        text, [DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d)], null, null, null,
        null, [], DocxTableCellMargins.Empty) { GridSpan = span };

    private static DocxTableRowLayout[] Layout(DocxTable table)
    {
        var document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default).Pages
            .SelectMany(page => page.Items).OfType<DocxTableRowLayout>().ToArray();
    }

    public static void DocxAutofitSpansMeasureIndependentIntervalsAndProjectSubdivisions()
    {
        foreach (bool longFirst in new[] { false, true })
        foreach (bool reverse in new[] { false, true })
        foreach (bool longInSpannedCell in new[] { false, true })
        {
            string shortText = "AAAAAA", longText = new('A', 20);
            var a = new DocxTableRow([
                Cell(longFirst && longInSpannedCell ? longText : shortText, 2), Cell(shortText, 2)], null);
            var b = new DocxTableRow([
                Cell(longFirst && !longInSpannedCell ? longText : shortText, 1), Cell(shortText, 2), Cell(shortText, 1)], null);
            var table = new DocxTable(null, [64d, 32d, 32d, 64d], reverse ? [b, a] : [a, b]);
            var rows = Layout(table);
            var first = rows[reverse ? 1 : 0];
            var second = rows[reverse ? 0 : 1];
            double unit = 30d + .96d, left = (longFirst ? 100d : 30d) + .96d;
            TestAssert.True(Math.Abs(second.Cells[0].Width - left) < .000001d,
                "A partially overlapping cell measures the independent column it contains.");
            TestAssert.True(Math.Abs(second.Cells[1].Width - unit) < .000001d, "The middle interval keeps its measured width.");
            TestAssert.True(Math.Abs(first.Cells[0].Width - left - unit / 2d) < .000001d,
                "Placement retains the shared subdivision without charging the neighboring content twice.");
            TestAssert.True(rows.All(row => row.Cells.All(cell => cell.TextLines.Count == 1)), "All source text retains its measured line.");
        }

        var asymmetric = new DocxTable(null, [48d, 16d, 64d, 64d], [
            new DocxTableRow([Cell("AAAAAA", 1), Cell("AAAAAA", 3)], null),
            new DocxTableRow([Cell("AAAAAA", 2), Cell("AAAAAA", 1), Cell("AAAAAA", 1)], null)]);
        var actual = Layout(asymmetric);
        TestAssert.True(Math.Abs(actual[0].Cells[0].Width - actual[1].Cells[0].Width) < .000001d,
            "An unoccupied subdivision collapses rather than shifting the next content column.");
        TestAssert.True(Math.Abs(actual[0].Cells[0].Width - 30.96d) < .000001d, "The first column follows content rather than its old grid hint.");

        var bridge = new DocxTable(null, [64d, 32d, 32d, 64d], [
            new DocxTableRow([Cell(new string('A', 20), 4)], null),
            new DocxTableRow([Cell("AAAAAA", 2), Cell("AAAAAA", 2)], null),
            new DocxTableRow([Cell("AAAAAA", 1), Cell("AAAAAA", 2), Cell("AAAAAA", 1)], null)]);
        var bridgeRows = Layout(bridge);
        TestAssert.True(Math.Abs(bridgeRows[0].Cells[0].Width - (100.96d + 30.96d)) < .000001d,
            "The independent Word bridge enlarges its single-grid columns and retains the middle spanned interval.");
    }

    public static void DocxAutofitSpansMergedConstraintsPayInsetsOnceAndKeepAuthoredWidths()
    {
        foreach (bool authoredMargins in new[] { false, true })
        {
            var margins = authoredMargins
                ? DocxTableCellMargins.Empty with { LeftPoints = 2d, RightPoints = 3d }
                : DocxTableCellMargins.Empty;
            var cells = new[] { Cell(new string('A', 20), 2), Cell("AAAAAA", 1), Cell("AAAAAA", 1) }
                .Select(cell => cell with { Margins = margins }).ToArray();
            var table = new DocxTable(null, [80d, 80d], [
                new DocxTableRow([cells[0]], null), new DocxTableRow([cells[1], cells[2]], null)]);
            var rows = Layout(table);
            double expected = 100d + (authoredMargins ? 5d : .96d);
            TestAssert.True(Math.Abs(rows[0].Cells[0].Width - expected) < .000001d,
                "A fully merged constraint pays its own insets once across measured columns.");
            TestAssert.True(Math.Abs(rows[1].Cells.Sum(cell => cell.Width) - expected) < .000001d, "Merged and ordinary rows share the same width.");
            TestAssert.True(rows.All(row => row.Cells.All(cell => cell.TextLines.Count == 1)),
                "Natural sizing and wrapping must agree at the merged cell's exact content width.");
            TestAssert.Equal(160d, Layout(table with { LayoutValue = "fixed" })[0].Cells[0].Width);
            TestAssert.True(Math.Abs(Layout(table with {
                PreferredWidthType = "dxa", PreferredWidthPoints = 140d })[0].Cells[0].Width - 140d) < .000001d, "An authored preferred table width retains its floor.");
        }

        // Neither of these crossing middle intervals owns an independent partition.
        var crossing = new DocxTable(null, [10d, 10d, 10d, 10d, 10d, 10d], [
            new DocxTableRow([Cell("A", 2), Cell("B", 3), Cell("C", 1)], null),
            new DocxTableRow([Cell("D", 1), Cell("E", 3), Cell("F", 2)], null)]);
        var fallback = Layout(crossing);
        TestAssert.Equal(20d, fallback[0].Cells[0].Width);
        TestAssert.Equal(30d, fallback[0].Cells[1].Width);
        TestAssert.Equal(10d, fallback[1].Cells[0].Width);
        TestAssert.Equal(30d, fallback[1].Cells[1].Width);
    }

    public static void DocxAutofitSpansMatchIndependentWordCells()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial)) TestAssert.Skip("Environmental precondition not met: Arial is unavailable.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "Cases", "AdjacentTables"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "autofit-spans-office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(folder, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            string text = string.Join(" ", control.GetProperty("Text").EnumerateArray().Select(value => value.GetString()));
            var font = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)),
                text.EnumerateRunes().Select(rune => rune.Value), default);
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(path, mode);
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var rows = new DocxLayoutEngine(geometry).Create(document, font, default).Pages
                    .SelectMany(page => page.Items).OfType<DocxTableRowLayout>().ToArray();
                var expectedRows = control.GetProperty("CellWidths").EnumerateArray().ToArray();
                TestAssert.Equal(expectedRows.Length, rows.Length);
                for (int index = 0; index < rows.Length; index++)
                {
                    var widths = expectedRows[index].EnumerateArray().Select(value => value.GetDouble()).ToArray();
                    TestAssert.Equal(widths.Length, rows[index].Cells.Count);
                    for (int cell = 0; cell < widths.Length; cell++)
                        TestAssert.True(Math.Abs(widths[cell] - rows[index].Cells[cell].Width) <= 1d,
                            $"Independent Word cell width differs: {control.GetProperty("Input").GetString()} / {index}/{cell}: {widths[cell]} vs {rows[index].Cells[cell].Width}.");
                    TestAssert.True(rows[index].Cells.All(cell => cell.TextLines.Count == 1), "Supported content must survive on its measured line.");
                }
            }
        }
    }
}
