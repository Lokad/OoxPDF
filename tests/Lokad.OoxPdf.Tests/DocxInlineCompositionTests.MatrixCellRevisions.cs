using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMatrixCellRevisionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent matrix cell revision frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "matrix-cell-revisions-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionMatrixCellRevisionsKeepUnrevisedMetricOwners()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string? part = story is "header" or "footer" ? "word/" + story + "1.xml" : null;
            var source = ReadMutation("CompositionProbes/matrix-cell-revision-" + story, mode, false, false, partOverride: part);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(source).First(paragraph =>
                paragraph.Equations.Any(equation => equation.Content.Children.Count == 1 && equation.Content.Children[0] is DocxMathMatrix));
            var equation = paragraph.Equations.Single();
            TestAssert.True(DocxMathLayout.GetStandaloneControl(equation, default)?.ControlText is { Revision: null, Revisions.Count: 0 } &&
                equation.Revision is null && DocxMathLayout.TryGetRuns(equation, default, out var runs) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true).Any(item => ReferenceEquals(item, equation)),
                "Literal cell revisions keep the actual unrevised metric owner across supported paragraph containers.");
            if (mode == OoxPdfDocxMarkupMode.Final)
                TestAssert.True(!equation.HasRunRevisions && DocxMathLayout.IsEmptyMatrix(equation, default),
                    "Final removes revised literal ink while retaining actual matrix slots.");
            else
                TestAssert.True(equation.HasRunRevisions && DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) &&
                    owners.Count == 4 && owners.All(owner => owner.Kind == DocxRevisionKind.Deletion) &&
                    !DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph),
                    "Four actual deleted cell sources own Review ink; the unchanged control cannot invent or erase those owners.");
        }
    }
}
