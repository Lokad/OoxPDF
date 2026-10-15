using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using System.Text;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableMathTests
{
    public static void DocxTableMathCandidatesShareDisplayPolicyAndKeepUnsupportedOwners()
    {
        var (paragraph, _) = Prepared(20d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var row = new DocxTableRow([cell], null);
        DocxDocument Document(DocxTableRow candidate) => TableDocument(candidate);
        TestAssert.True(ReferenceEquals(paragraph, DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(row), CancellationToken.None).Single()),
            "Automatic cell display uses the same source paragraph identity as prepared math.");
        foreach (DocxTableRow unsupported in new[]
        {
            row with { HeightPoints = 12d, HeightRuleValue = "exact" },
            row with { Cells = [cell with { HasVerticalMerge = true }] },
            row with { Cells = [cell with { FitText = true }] },
            row with { Cells = [cell with { TextDirectionValue = "tbRl" }] },
            row with { Cells = [cell with { Paragraphs = [paragraph with { Equations = [paragraph.Equations[0] with { Display = false }] }] }] },
            row with { Cells = [cell with { BodyElements = [new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, null, null)] }] }
        })
            TestAssert.Equal(0, DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(unsupported), CancellationToken.None).Count());
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(row), cancelled.Token).ToArray());
    }

    public static void DocxTableMathMeasurementMemoAndFollowingFlowUseOneBoxAdvance()
    {
        var (paragraph, math) = Prepared(20d);
        DocxParagraph following = DocxTests.CreateDocxLayoutParagraph("Following cell.", 10d, 10d);
        var cell = new DocxTableCell(string.Empty, [paragraph, following], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxDocument document = TableDocument(new DocxTableRow([cell], null));
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxLayout Layout(double height) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None, mathParagraphs:
                new Dictionary<DocxParagraph, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                { [paragraph] = math with { LineHeight = height } });
        DocxLayout first = Layout(20d), expanded = Layout(30d);
        var a = first.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        var b = expanded.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        TestAssert.True(Math.Abs(b.Height - a.Height - 10d) < .000001,
            "The row must reserve the same extra math advance that cell placement consumes.");
        var originalLines = a.Cells.Single().TextLines;
        var expandedLines = b.Cells.Single().TextLines;
        TestAssert.True(ReferenceEquals(originalLines[0].Math, math) && originalLines[0].Segments.Count == 0,
            "A math line carries the prepared box, not a duplicate prose encoding.");
        TestAssert.True(Math.Abs(originalLines[0].BaselineY - expandedLines[0].BaselineY) < .000001 &&
            Math.Abs(originalLines[1].BaselineY - expandedLines[1].BaselineY - 10d) < .000001,
            "Changing only occupied math height preserves its baseline and shifts following cell prose.");
        double afterA = first.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        double afterB = expanded.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        TestAssert.True(Math.Abs(afterA - afterB - 10d) < .000001,
            "Following body flow must observe the same measured row extent, including a fresh memo per layout.");
    }

    public static void DocxTableMathKeepEstimateUsesThePreparedFirstRowExtent()
    {
        var (paragraph, math) = Prepared(15d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxDocument source = TableDocument(new DocxTableRow([cell], null));
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Fill", 10d, 30d);
        DocxParagraph keep = DocxTests.CreateDocxLayoutParagraph("Keep", 10d, 10d,
            new DocxParagraphKeepRules(true, null, null, null, null, null));
        var table = source.BodyElements.OfType<DocxTableElement>().Single();
        DocxDocument document = source with { PageHeightPoints = 80d,
            BodyElements = [new DocxParagraphElement(filler), new DocxParagraphElement(keep), table] };
        foreach ((double height, int pages) in new[] { (15d, 1), (25d, 2) })
        {
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None,
                    mathParagraphs: new Dictionary<DocxParagraph, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                    { [paragraph] = math with { LineHeight = height } });
            TestAssert.Equal(pages, layout.Pages.Count);
            var tablePage = layout.Pages.Single(page => page.Items.OfType<DocxTableRowLayout>().Any());
            TestAssert.True(tablePage.Items.OfType<DocxTextLineLayout>().Any(line => line.Text == "Keep"),
                "Keep-with-next and math row placement must move as the same measured block.");
        }
    }

    private static (DocxParagraph Paragraph, DocxPreparedMathParagraph Math) Prepared(double height)
    {
        var run = new DocxTextRun("A+B", 10d, null, false, false, false, null, "TestFont");
        var row = new DocxMathRow([new DocxMathRun(run, "p", null)]);
        var equation = new DocxEquation(row, 0, 0, true, null, null, null) { ParentFontSize = 10d };
        var font = new PdfFontResource("Eq1", PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(),
            run.Text.EnumerateRunes().Select(rune => rune.Value), CancellationToken.None));
        TestAssert.True(DocxMathRowLayout.TryMeasure(row, _ => font, CancellationToken.None, out var box, out _), "Expected portable math glyphs.");
        var paragraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with
            { Runs = [], LineSpacingPoints = null, Equations = [equation] };
        return (paragraph, new(equation, TestAssert.NotNull(box), 12d, height));
    }

    private static DocxDocument TableDocument(DocxTableRow row)
    {
        var table = new DocxTable(null, [100d], [row]) { PreferredWidthPoints = 100d };
        return DocxTests.CreateLayoutTestDocument([new DocxTableElement(table),
            new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Following body.", 10d, 10d))], [table]);
    }
}
