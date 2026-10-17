using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEmptyRootDegreesKeepVisibleRootBodyAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent empty degree geometry uses installed Windows fonts.");
        // Independent geometry retains inherited large-control failures. This
        // observer qualifies restored content/source ink, not those failed gates.
        CheckEquationRevisions(null, false, "math-degree-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionEmptyRootDegreesKeepKnownRolesAndHiddenFallback()
    {
        var document = ReadMutation("CompositionProbes/math-degree-display-cube-empty", OoxPdfDocxMarkupMode.Final, false, false);
        var equation = document.Paragraphs.Single(p => p.Equations.Count != 0).Equations.Single();
        TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var sources) && sources.Count != 0 &&
            sources.All(source => source.Run.Text.Text.Length != 0), "The empty degree does not invent source runs.");
        var unknown = ReadMutation("CompositionProbes/math-degree-display-cube-empty", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, m) => xml.Descendants(m + "deg").Single().ReplaceNodes(new XElement(m + "future")))
            .Paragraphs.Single(p => p.Equations.Count != 0).Equations.Single();
        TestAssert.True(!DocxMathLayout.TryGetRuns(unknown, default, out _), "Unknown degree children retain whole-equation fallback.");
        var storedHidden = ReadMutation("docx-omml-root-degree-display-cube", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, m) => xml.Descendants(m + "degHide").Single().SetAttributeValue(m + "val", "1"))
            .Paragraphs.Single(p => p.Equations.Count != 0).Equations.Single();
        TestAssert.True(!DocxMathLayout.TryGetRuns(storedHidden, default, out _), "Stored hidden degrees retain the independently qualified fallback.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, canceled.Token, out _));
    }
}
