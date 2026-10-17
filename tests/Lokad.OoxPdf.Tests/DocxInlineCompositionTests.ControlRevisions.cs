using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionControlRevisionsKeepWordContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent control revision ink uses installed Windows fonts.");
        // Source-backed content and ink, including bookmark/run-split transforms.
        // Full one-point placement and strict visual gates are recorded separately.
        CheckEquationRevisions(null, false, "math-control-revision-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionControlRevisionsRetainActualOwnersAndRoleFallbacks()
    {
        foreach (string role in new[] { "script", "function", "matrix" })
        {
            var source = ReadMutation("CompositionProbes/math-control-revision-" + role,
                OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var paragraph = source.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            TestAssert.True(equation.Revision is null && equation.HasRunRevisions &&
                DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) && owners.Count == 1 &&
                paragraph.Revisions.All(owner => ReferenceEquals(owner, owners[0])) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "A supported control retains its actual literal revision owner through the shared grammar.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.TryGetEquationRevisionOwners(equation, canceled.Token, out _));
        }
        foreach (string orphan in new[] { "e", "sub", "sup", "num", "den", "deg", "fName", "mr" })
        {
            var source = ReadMutation("CompositionProbes/math-control-revision-script", OoxPdfDocxMarkupMode.AllMarkup,
                false, false, (xml, w, m) => xml.Descendants(m + "sSubSup").First().Name = m + orphan);
            TestAssert.True(!DocxMathLayout.TryGetRuns(source.Paragraphs.First().Equations.Single(), default, out _),
                "Role and matrix rows cannot gain a source parent from revision syntax.");
        }
    }
}
