using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionReviewNestedPlacementMatchesIndependentWordY()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent nested table origins use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("nested-review-placement-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = NestedMathCasePath(id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(DocxTests.ReadDocx(path, mode), null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true);
            var math = shows.Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                ReferenceEquals(font.Font, show.Font))).ToArray();
            var words = row.GetProperty("Math").EnumerateArray().ToArray();
            TestAssert.Equal(words.Length * 4, math.Length);
            for (int i = 0; i < words.Length; i++)
            {
                TestAssert.Equal(words[i].GetProperty("Text").GetString(), string.Concat(math.Skip(i * 4).Take(4).Select(show => show.Text)));
                TestAssert.True(Math.Abs(math[i * 4].Y - words[i].GetProperty("Y").GetDouble()) < .5d,
                    $"Mathematical frames retain independent Word Y across source order and nesting: {id}/{mode}/word={i}.");
            }
            foreach (var prose in row.GetProperty("Prose").EnumerateArray())
            {
                var actual = shows.Single(show => show.Text == prose.GetProperty("Text").GetString());
                TestAssert.True(Math.Abs(actual.Y - prose.GetProperty("Y").GetDouble()) < .5d,
                    $"Surrounding prose retains independent Word Y: {id}/{mode}/{actual.Text}.");
            }
        }
    }

    public static void DocxInlineCompositionReviewNestedPlacementSeparatesPhysicalSlotsAcrossMetrics()
    {
        foreach (double size in new[] { 9d, 12d, 17d })
        foreach (double scale in new[] { .5d, .75d, .9d })
        foreach (int depth in new[] { 1, 2 })
        {
            var seed = ReadMutation("CompositionProbes/mixed-nested-cell-inside", OoxPdfDocxMarkupMode.Final, false, false);
            var math = DocxMathParagraphPolicy.EnumeratePreparationCandidates(seed, default)
                .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                    equation => new DocxPreparedMathParagraph(equation, new(18d, 20d, 8d, []), 28d, 36d));
            var head = DocxTests.CreateDocxLayoutParagraph("Public head", size, size) with { LineSpacingPoints = null };
            var tail = DocxTests.CreateDocxLayoutParagraph("Public tail", size, size) with { LineSpacingPoints = null };
            var table = seed.Tables[0].Rows[0].Cells[0].BodyElements.OfType<DocxTableElement>().Single().Table;
            for (int index = 0; index < depth; index++)
            {
                var cell = new DocxTableCell(string.Empty, [head, tail], null, null, null, null, [], DocxTableCellMargins.Empty)
                    { BodyElements = [new DocxParagraphElement(head), new DocxTableElement(table), new DocxParagraphElement(tail)] };
                table = new DocxTable(null, [468d], [new DocxTableRow([cell], null)]);
            }
            var source = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            var natural = new DocxTests.FamilyWidthTextMeasurer();
            var measured = new DocxTests.ScaledLayoutTextMeasurer(natural, scale, scale);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: scale)
                .Create(source, measured, default, unscaledTextMeasurer: natural, mathEquations: math);
            var rows = NestedMathRows(layout.Pages.Single().Items.OfType<DocxTableRowLayout>()).ToArray();
            TestAssert.Equal(depth + 1, rows.Length);
            TestAssert.Equal(2, rows.SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines).Sum(line => line.InlineMath.Count));
            foreach (var row in rows.Where(row => row.Cells[0].NestedRows.Count != 0))
            {
                var cell = row.Cells.Single();
                var headLine = cell.TextLines.First();
                TestAssert.True(DocxInlineLineGeometry.TryGetOccupiedBounds(headLine, out double bottom, out _),
                    "A preceding ordinary line exposes its physical occupied frame.");
                TestAssert.True(cell.NestedRows[0].Y + cell.NestedRows[0].Height <= bottom + .000001d,
                    $"The nested block follows the physical ordinary slot, once at each depth: size={size}/scale={scale}/depth={depth}.");
            }
        }
    }

    public static void DocxInlineCompositionReviewNestedPlacementKeepsWholeTreeFallbacks()
    {
        var source = ReadMutation("CompositionProbes/mixed-nested-cell-inside", OoxPdfDocxMarkupMode.Final, false, false);
        var math = DocxMathParagraphPolicy.EnumeratePreparationCandidates(source, default)
            .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 20d, 8d, []), 28d, 36d));
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, measurer, default, mathEquations: math);
        var row = layout.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        var cell = row.Cells.Single(); var child = cell.NestedRows.Single(); var leaf = child.Cells.Single();
        foreach (var bad in new[] {
            child with { FragmentCount = 2, FragmentReason = "PageBoundary" },
            child with { HeightRuleValue = "unknown", DeclaredHeightPoints = 20d },
            child with { HeightRuleValue = "exact", DeclaredHeightPoints = 1d, Height = 1d },
            child with { Height = .1d, Cells = [leaf with { Height = .1d }] },
            child with { Cells = [leaf with { Cell = leaf.Cell with { FitText = true, FitTextValue = "1" } }] },
            child with { Cells = [leaf with { Cell = leaf.Cell with { VerticalAlignmentValue = "center" } }] },
            child with { Cells = [leaf with { InlineImages = leaf.InlineImages.Select(image => image with { SourceLineIndex = null }).ToArray() }] }
        })
        {
            var guarded = row with { Cells = [cell with { NestedTableRows = [bad] }] };
            var items = new List<DocxLayoutItem> { row, guarded }; double cursor = 17d;
            DocxLayoutEngine.ApplyReviewTableBaselineGeometry(items, 0, 12d, measurer, .75d, ref cursor, default, allowPageBoundaryFragments: true);
            TestAssert.True(ReferenceEquals(row, items[0]) && ReferenceEquals(guarded, items[1]) && cursor == 17d,
                "An unsupported descendant retains every earlier row and following flow before publication.");
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        var canceledItems = new List<DocxLayoutItem> { row }; double canceledCursor = 17d;
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.ApplyReviewTableBaselineGeometry(
            canceledItems, 0, 12d, measurer, .75d, ref canceledCursor, canceled.Token));
        TestAssert.True(ReferenceEquals(row, canceledItems[0]) && canceledCursor == 17d,
            "Canceled nested placement publishes no partial table.");
    }
}
