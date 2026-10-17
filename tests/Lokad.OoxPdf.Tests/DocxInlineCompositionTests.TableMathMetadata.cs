using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionTableMathMetadataKeepsIndependentWordContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent table metadata uses installed Windows fonts.");
        // Cell insertion/deletion markers are unchanged counterfactuals. Their
        // existing missing review bars remain in the complete raw comparisons.
        CheckEquationRevisions(null, false, "table-math-metadata-office.json", splitMathRuns: true,
            includeCase: id => !id.EndsWith("-cell-ins", StringComparison.Ordinal) && !id.EndsWith("-cell-del", StringComparison.Ordinal));
    }

    public static void DocxInlineCompositionTableMathMetadataKeepsChildOwnersAndDeletingFallback()
    {
        foreach (string role in new[] { "row-ins", "row-format", "cell-format", "cell-ins", "cell-del",
            "header-row-ins", "header-cell-format", "footer-row-ins", "footer-cell-format" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/table-math-revision-" + role, mode, false, false);
            var table = document.EquationParts.SelectMany(part => part.BodyElements.OfType<DocxTableElement>())
                .Select(element => element.Table).First(table => table.Rows.Any(row => row.Cells.Any(cell =>
                    cell.Paragraphs.Any(paragraph => paragraph.Equations.Count != 0))));
            var row = table.Rows.Single();
            var cell = row.Cells.Single();
            TestAssert.True(DocxMathParagraphPolicy.IsSupportedSimpleRow(row) &&
                DocxMathParagraphPolicy.IsSupportedSimpleCell(cell) &&
                cell.Paragraphs.Sum(paragraph => paragraph.Equations.Count) == 4 &&
                cell.Paragraphs.All(paragraph => paragraph.Equations.All(equation =>
                    DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true)
                        .Any(item => ReferenceEquals(item, equation)))),
                "Non-deleting table metadata retains each actual child's equation and preparation owner.");
        }
        var deleted = ReadMutation("CompositionProbes/table-math-revision-row-del", OoxPdfDocxMarkupMode.AllMarkup, false, false);
        TestAssert.True(!DocxMathParagraphPolicy.IsSupportedSimpleRow(deleted.Tables.Single().Rows.Single()),
            "Deleting table owners retain the staged fallback until row visibility is independently handled.");
        var inserted = ReadMutation("CompositionProbes/table-math-revision-row-ins", OoxPdfDocxMarkupMode.Final, false, false);
        var source = inserted.Tables.Single().Rows.Single();
        var unknown = source with { Revisions = [source.Revisions.Single() with { Kind = (DocxRevisionKind)int.MaxValue }] };
        TestAssert.True(!DocxMathParagraphPolicy.IsSupportedSimpleRow(unknown), "Unknown owner kinds retain complete fallback.");
    }

    public static void DocxInlineCompositionTableMathMetadataSharesPrintedLaneAcrossContainers()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var body = ReadMutation("CompositionProbes/table-math-revision-cell-format", OoxPdfDocxMarkupMode.AllMarkup, false, false);
        double bodyScale = DocxRenderer.ResolveWordCompatiblePrintScale(body, context);
        TestAssert.True(bodyScale < 1d, "A nonvoid formatting balloon reserves the printed page lane.");
        foreach (string story in new[] { "header", "footer" })
        {
            var formatted = ReadMutation($"CompositionProbes/table-math-revision-{story}-cell-format",
                OoxPdfDocxMarkupMode.AllMarkup, false, false);
            TestAssert.Equal(bodyScale, DocxRenderer.ResolveWordCompatiblePrintScale(formatted, context));
            var inserted = ReadMutation($"CompositionProbes/table-math-revision-{story}-row-ins",
                OoxPdfDocxMarkupMode.AllMarkup, false, false);
            TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(inserted, context));
        }
    }
}
