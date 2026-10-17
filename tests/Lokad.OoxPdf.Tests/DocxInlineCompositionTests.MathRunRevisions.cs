using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMathRunRevisionsRequireActualOwners()
    {
        var source = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = source.Paragraphs.First();
        var equation = paragraph.Equations.Single();
        var revision = paragraph.Revisions.Single();
        TestAssert.True(equation.Revision is null && equation.HasRunRevisions && paragraph.InlineContent is not null &&
            DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
            "The actual numerator run owns its revision without revising the denominator or whole equation.");
        foreach (var invalid in new[] { paragraph with { Revisions = [revision with { }] },
            paragraph with { Equations = [equation with { HasRunRevisions = false }] },
            paragraph with { Equations = [equation with { SourceRunIndex = -1 }] } })
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(invalid, default),
                "Copied aggregate metadata and a discovery hint cannot invent source ownership.");
        foreach (string kind in new[] { "moveTo", "rPrChange" })
        {
            var unknown = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, m) => xml.Descendants(w + "ins").Single().Name = w + kind);
            TestAssert.True(!DocxMathLayout.TryGetRuns(unknown.Paragraphs.First().Equations.Single(), default, out _),
                "Unknown structural revisions retain the complete equation fallback.");
        }
        var empty = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
            (xml, w, m) => xml.Descendants(w + "ins").Single().Descendants(m + "t").Single().Value = string.Empty);
        TestAssert.True(!DocxMathLayout.TryGetRuns(empty.Paragraphs.First().Equations.Single(), default, out _),
            "Unqualified empty inner revision wrappers retain their existing fallback.");
        var deleted = ReadMutation("CompositionProbes/math-run-revision-body-del", OoxPdfDocxMarkupMode.Final, false, false);
        var visible = deleted.Paragraphs.First();
        TestAssert.True(visible.Revisions.Count == 0 && !visible.Equations.Single().HasRunRevisions &&
            DocxMathLayout.TryGetRuns(visible.Equations.Single(), default, out var runs) &&
            runs.Count == 1 && runs[0].Run.Text.Text == "y",
            "Final retains the denominator and empty numerator structure without deleted glyph ownership.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
    }

    public static void DocxInlineCompositionMathRunRevisionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent inner-math revision frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "math-run-revisions-office.json", splitMathRuns: true);
    }
}
