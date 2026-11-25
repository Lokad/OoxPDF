using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableFlowTests
{
    internal static string CasePath(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "MixedTables", name));

    public static void DocxCellFlowWithoutFontMeasurementKeepsAuthoredPictureAdvance()
    {
        var first = new DocxInlineImage(18d, 12d, "image/png", [1, 2, 3], "/word/media/first.png") { SourceRunIndex = 0 };
        var second = new DocxInlineImage(24d, 18d, "image/png", [1, 2, 3], "/word/media/second.png") { SourceRunIndex = 0 };
        DocxParagraph Picture(DocxInlineImage image) => DocxTests.CreateDocxLayoutParagraph("", 10d, 10d) with
        { Runs = [], Images = [image], SpacingBeforePoints = 0d, SpacingAfterPoints = 0d };
        var cell = new DocxTableCell([], null, null, null, null, [], DocxTableCellMargins.Empty)
        { BodyElements = [new DocxParagraphElement(Picture(first)), new DocxParagraphElement(Picture(second))] };
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 80d)]);
        var document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)]);
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var placed = engine.Create(document, embedded: null, default).Pages.Single().Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
        TestAssert.Equal(0, placed.TextLines.Count);
        TestAssert.Equal(2, placed.InlineImages.Count);
        TestAssert.True(ReferenceEquals(first, placed.InlineImages[0].Image) && ReferenceEquals(second, placed.InlineImages[1].Image),
            "Body elements retain each authored picture and its source order without font measurement.");
        TestAssert.True(Math.Abs(placed.InlineImages[0].Y - placed.InlineImages[1].Y - second.HeightPoints) < .000001d,
            "One geometry cursor consumes each source-owned picture height exactly once.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => engine.Create(document, embedded: null, cancellation.Token));
    }

    public static void DocxCellFlowNestedTablesFollowMeasuredDisplayMath()
    {
        // Change only the prepared display frame's advance. The following
        // nested table must consume that same advance, rather than rewrapping
        // the equation paragraph as empty ordinary text.
        foreach (string name in new[] { "maintenance-matrix", "capacity-schedule" })
        {
            var document = DocxTests.ReadDocx(CasePath(name + ".docx"), OoxPdfDocxMarkupMode.Final);
            var source = document.Tables[0].Rows[2].Cells[1].Paragraphs.Single(p => p.Equations.Count != 0);
            var equation = source.Equations.Single();
            (double MathY, double NestedTop, double AfterY) Place(double advance)
            {
                var math = new Dictionary<DocxEquation, DocxPreparedMathParagraph>
                { [equation] = new(equation, new(18d, 20d, 8d, []), 28d, advance) };
                var page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math).Pages.Single();
                var cell = page.Items.OfType<DocxTableRowLayout>().Single(row => row.RowIndex == 2).Cells[1];
                return (cell.TextLines.Single(line => ReferenceEquals(line.SourceParagraph, source)).BaselineY,
                    cell.NestedRows.Single().Y + cell.NestedRows.Single().Height,
                    cell.TextLines.Single(line => line.Text.StartsWith("After nested", StringComparison.Ordinal)).BaselineY);
            }
            var first = Place(40d); var taller = Place(60d);
            TestAssert.True(Math.Abs(first.MathY - taller.MathY) < .000001d,
                "Changing a following block's advance cannot move its own first baseline.");
            TestAssert.True(Math.Abs(first.NestedTop - taller.NestedTop - 20d) < .000001d &&
                Math.Abs(first.AfterY - taller.AfterY - 20d) < .000001d,
                $"The nested table and subsequent text share the measured display advance: {name}, first={first}, taller={taller}.");
        }
    }

    public static void DocxCellFlowIndependentTablesRetainVisibleContent()
    {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CasePath("expected.json")));
        foreach (var row in manifest.RootElement.EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string path = CasePath(row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var document = DocxTests.ReadDocx(path, mode);
            TestAssert.Equal(row.GetProperty("Rows").GetInt32(), document.Tables[0].Rows.Count);
            TestAssert.Equal(row.GetProperty("Columns").GetInt32(), document.Tables[0].ColumnWidthsPoints.Count);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            string painted = Regex.Replace(string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false)).Select(show => show.Text)), @"\s+", "");
            foreach (var text in row.GetProperty("Text").EnumerateArray())
                TestAssert.Contains(Regex.Replace(text.GetString()!, @"\s+", ""), painted);
            string math = string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                honorActualText: false, fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))).Select(show => show.Text));
            foreach (var expected in row.GetProperty("Math").EnumerateArray()) TestAssert.Contains(expected.GetString()!, math);
            TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "A surrounding larger table cannot discard supported mathematics.");
            foreach (var page in pages) PdfContentValidator.ValidatePage(page, 0, default);
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null, cancelled.Token).ToBlockingEnumerable().ToArray());
        }
    }
}
