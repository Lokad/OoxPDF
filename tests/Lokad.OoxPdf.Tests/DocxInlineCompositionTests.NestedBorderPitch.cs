using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNestedBorderPitchMatchesIndependentWordHeightAndFollowingFlow()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent nested border controls use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("nested-border-pitch-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = NestedMathCasePath(id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(path, mode);
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var root = renderer.InspectLayout(source).Pages.Single().TableRows.Single();
            TestAssert.True(Math.Abs(root.Height - row.GetProperty("RootRowHeight").GetDouble()) < .5d,
                $"The complete nested row consumes independent Word's printed height: {id}/{mode}.");
            var page = renderer.RenderBlankPages(source, null, default).Single();
            var following = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Single(show => show.Text == "Following prose.");
            TestAssert.True(Math.Abs(following.Y - row.GetProperty("FollowingY").GetDouble()) < .5d,
                $"Following body text consumes all printed nested advances: {id}/{mode}.");
        }
    }

    public static void DocxInlineCompositionNestedBorderPitchScalesRowsAndSourceFlowAcrossCompetingCells()
    {
        foreach (double size in new[] { 9d, 12d, 17d })
        foreach (double scale in new[] { .5d, .75d, .9d })
        foreach (bool minimum in new[] { false, true })
        foreach (bool tallNeighbor in new[] { false, true })
        foreach (string borderSize in new[] { "4", "16" })
        {
            var seed = ReadMutation("CompositionProbes/mixed-nested-cell-inside", OoxPdfDocxMarkupMode.Final, false, false);
            var mathCell = seed.Tables[0].Rows[0].Cells[0].BodyElements.OfType<DocxTableElement>().Single().Table.Rows[0].Cells[0];
            var prepared = DocxMathParagraphPolicy.EnumeratePreparationCandidates(seed, default)
                .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                    equation => new DocxPreparedMathParagraph(equation, new(18d, 20d, 8d, []), 28d, 36d));
            var borders = new[] { "top", "bottom", "left", "right" }.Select(edge =>
                new DocxTableCellBorder(edge, "single", "000000", borderSize)).ToArray();
            DocxParagraph Prose(string text) => DocxTests.CreateDocxLayoutParagraph(text, size, size) with { LineSpacingPoints = null };
            DocxTableCell Plain(string text) => new(text, [Prose(text)], null, null, null, null, borders, DocxTableCellMargins.Empty);
            var childA = new DocxTable(null, [220d], [new DocxTableRow([mathCell with { Borders = borders }], null), new DocxTableRow([Plain("Public inner tail")], null)]);
            var childB = new DocxTable(null, [220d], [new DocxTableRow([Plain("Public second block")], null)]);
            var head = Prose("Public head"); var middle = Prose("Public middle"); var tail = Prose("Public tail");
            var left = new DocxTableCell(string.Empty, [head, middle, tail], null, null, null, null, borders, DocxTableCellMargins.Empty)
            { BodyElements = [new DocxParagraphElement(head), new DocxTableElement(childA), new DocxParagraphElement(middle),
                new DocxTableElement(childB), new DocxParagraphElement(tail)] };
            var right = Plain("Public neighbor") with { Paragraphs = Enumerable.Range(0, tallNeighbor ? 12 : 1)
                .Select(index => Prose("Public neighbor " + index)).ToArray() };
            var outer = new DocxTable(null, [234d, 234d], [new DocxTableRow([left, right], minimum ? 240d : null)
                { HeightRuleValue = minimum ? "atLeast" : null }]);
            // Keep the authored grid inside the natural page width: otherwise
            // different width clamping changes wrapping and invalidates a pure
            // vertical scaling comparison.
            var source = DocxTests.CreateLayoutTestDocument([new DocxTableElement(outer)], [outer]) with
            { PageWidthPoints = 612d, PageHeightPoints = 792d, MarginLeftPoints = 72d, MarginRightPoints = 72d,
                MarginTopPoints = 72d, MarginBottomPoints = 72d };
            var natural = new DocxTests.FamilyWidthTextMeasurer();
            var final = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, natural, default, mathEquations: prepared);
            var review = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: scale)
                .Create(source, new DocxTests.ScaledLayoutTextMeasurer(natural, scale, scale), default,
                    unscaledTextMeasurer: natural, mathEquations: prepared);
            var before = NestedMathRows(final.Pages.Single().Items.OfType<DocxTableRowLayout>()).ToArray();
            var after = NestedMathRows(review.Pages.Single().Items.OfType<DocxTableRowLayout>()).ToArray();
            TestAssert.Equal(4, before.Length); TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
                TestAssert.True(Math.Abs(after[index].Height - before[index].Height * scale) < .000001d,
                    $"Printed nested height honors all rows, competing cell maxima and declared minima: size={size}/scale={scale}/minimum={minimum}/neighbor={tallNeighbor}/border={borderSize}/row={index}/actual={after[index].Height}/expected={before[index].Height * scale}/borderScale={after[index].ReviewBorderPrintScale}/allAfter={string.Join(',', after.Select(row => row.Height))}/allBefore={string.Join(',', before.Select(row => row.Height))}.");
            var beforeLines = before[0].Cells[0].TextLines;
            var afterLines = after[0].Cells[0].TextLines;
            TestAssert.Equal(3, beforeLines.Count); TestAssert.Equal(3, afterLines.Count);
            for (int index = 1; index < beforeLines.Count; index++)
                TestAssert.True(Math.Abs((afterLines[0].BaselineY - afterLines[index].BaselineY) -
                    (beforeLines[0].BaselineY - beforeLines[index].BaselineY) * scale) < .000001d,
                    "Each source paragraph consumes the reductions of preceding nested blocks, independently of the cell that controls row height.");
        }
    }
}
