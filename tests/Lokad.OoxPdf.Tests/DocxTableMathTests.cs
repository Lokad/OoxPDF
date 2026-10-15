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
        foreach (DocxTableRow supported in new[] { row,
            row with { HeightPoints = 12d, HeightRuleValue = "exact" },
            row with { HeightPoints = 50d, HeightRuleValue = "atLeast" },
            row with { HeightPoints = 50d, HeightValue = "1000" },
            row with { HeightPoints = 0d, HeightRuleValue = "exact" } })
            TestAssert.True(ReferenceEquals(paragraph, DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(supported), CancellationToken.None)
                .Single(source => source.Equations.Count != 0)),
                "Resolved cell frames retain the same source paragraph identity as prepared math.");
        foreach (DocxTableRow unsupported in new[]
        {
            row with { HeightPoints = 12d, HeightRuleValue = "unknown" },
            row with { HeightPoints = -12d, HeightRuleValue = "exact" },
            row with { HeightPoints = double.PositiveInfinity, HeightRuleValue = "exact" },
            row with { HeightValue = "bad", HeightRuleValue = "exact" },
            row with { HeightRuleValue = "exact" },
            row with { Cells = [cell with { HasVerticalMerge = true }] },
            row with { Cells = [cell with { FitText = true }] },
            row with { Cells = [cell with { TextDirectionValue = "tbRl" }] },
            row with { Cells = [cell with { Paragraphs = [paragraph with { Equations = [paragraph.Equations[0] with { Display = false }] }] }] },
            row with { Cells = [cell with { BodyElements = [new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, null, null)] }] }
        })
            TestAssert.Equal(0, DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(unsupported), CancellationToken.None)
                .Count(source => source.Equations.Count != 0));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(row), cancelled.Token).ToArray());
    }

    public static void DocxTableZeroExactHeightRetainsTheNaturalContentFrame()
    {
        foreach (bool bordered in new[] { false, true })
        {
            DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Public cell.", 10d, 12d);
            var cell = new DocxTableCell("Public cell.", [paragraph], null, null, null, null,
                bordered ? [new DocxTableCellBorder("bottom", "single", "243042", "4")] : [], DocxTableCellMargins.Empty);
            DocxLayout Layout(string rule)
            {
                DocxDocument source = TableDocument(new DocxTableRow([cell], 0d) { HeightRuleValue = rule });
                return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(source, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            }
            DocxLayout exact = Layout("exact"), minimum = Layout("atLeast");
            var exactRow = exact.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
            var naturalRow = minimum.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.True(exactRow.Height >= 12d && Math.Abs(exactRow.Height - naturalRow.Height) < .000001,
                "Zero exact height uses natural content, including its visible border, rather than a one-point clip.");
            double afterExact = exact.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            double afterNatural = minimum.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            TestAssert.True(Math.Abs(afterExact - afterNatural) < .000001,
                "Following flow must use the same natural zero-height frame.");
        }
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
        TestAssert.True(originalLines[0].Math is { } prepared && ReferenceEquals(prepared.Box, math.Box) &&
            prepared.LineHeight == 20d && originalLines[0].Segments.Count == 0,
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

    public static void DocxTableMathInlineMeasurementAndWrappingRetainSourceOwnership()
    {
        var (display, prepared) = Prepared(24d);
        DocxTextRun before = new("Before ", 10d, null, false, false, false, null, "TestFont");
        DocxTextRun after = before with { Text = " after public equation with more prose." };
        DocxEquation equation = prepared.Equation with { Display = false, SourceRunIndex = 1 };
        DocxParagraph paragraph = display with { Runs = [before, after], Equations = [equation] };
        var math = prepared with { Equation = equation };
        var cell = new DocxTableCell(string.Concat(paragraph.Runs.Select(run => run.Text)), [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var document = TableDocument(new DocxTableRow([cell], null));
        TestAssert.True(ReferenceEquals(paragraph, DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
            .Single(source => source.Equations.Count != 0)), "Inline cells share the body semantic owner policy.");
        DocxLayout Layout(double height) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathParagraphs:
                new Dictionary<DocxParagraph, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                { [paragraph] = math with { LineHeight = height } });
        DocxLayout initial = Layout(24d), expanded = Layout(34d);
        var a = initial.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        var b = expanded.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        var lines = a.Cells.Single().TextLines;
        TestAssert.True(lines.Count > 1 && lines.Count(line => line.Math is not null) == 1 &&
            lines[0].IsInlineMath && lines[0].InlineMathOffsetX > 0d,
            "A complete box stays at its source boundary while its following prose wraps.");
        string prose = string.Concat(lines.Select(line => line.Text));
        TestAssert.True(!prose.Contains('\u200B') && new[] { "Before", "after", "public", "equation", "with", "more", "prose." }
            .All(word => prose.Contains(word, StringComparison.Ordinal)), "Every public prose word survives without layout markers.");
        TestAssert.True(Math.Abs(b.Height - a.Height - 10d) < .000001,
            "Cell measurement reserves exactly the extra descent consumed by inline placement.");
        double afterA = initial.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        double afterB = expanded.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        TestAssert.True(Math.Abs(afterA - afterB - 10d) < .000001, "Following body flow shares the measured row advance.");
        TestAssert.Equal(before.Text + after.Text, string.Concat(paragraph.Runs.Select(run => run.Text)));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), cancelled.Token,
                mathParagraphs: new Dictionary<DocxParagraph, DocxPreparedMathParagraph> { [paragraph] = math }));
    }

    public static void DocxTableMathInlineUnfitBoxesPreserveOrdinaryRowSplitting()
    {
        var (display, prepared) = Prepared(24d);
        DocxTextRun run = new(string.Join(" ", Enumerable.Repeat("Public prose", 20)), 10d, null, false, false, false, null, "TestFont");
        DocxEquation equation = prepared.Equation with { Display = false, SourceRunIndex = 1 };
        DocxParagraph paragraph = display with { Runs = [run], Equations = [equation] };
        var cell = new DocxTableCell(string.Concat(paragraph.Runs.Select(run => run.Text)), [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxDocument source = TableDocument(new DocxTableRow([cell], 150d) { HeightRuleValue = "exact" }) with { PageHeightPoints = 100d };
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxLayout plain = engine.Create(source, measurer, default);
        DocxLayout fallback = engine.Create(source, measurer, default, mathParagraphs:
            new Dictionary<DocxParagraph, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
            { [paragraph] = prepared with { Equation = equation, Box = prepared.Box with { Width = 1000d } } });
        TestAssert.Equal(plain.Pages.Count, fallback.Pages.Count);
        var ordinaryRows = plain.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
        var fallbackRows = fallback.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
        TestAssert.True(ordinaryRows.Length > 1, "This control must exercise actual row fragments across pages.");
        TestAssert.Equal(ordinaryRows.Length, fallbackRows.Length);
        for (int index = 0; index < ordinaryRows.Length; index++)
        {
            TestAssert.True(Math.Abs(ordinaryRows[index].Height - fallbackRows[index].Height) < .000001,
                "An unpainted prepared box cannot make an otherwise splittable prose row atomic.");
            TestAssert.True(fallbackRows[index].Cells.SelectMany(c => c.TextLines).All(line => line.Math is null),
                "An unfit equation retains the whole fallback.");
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
