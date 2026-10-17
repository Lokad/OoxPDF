using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStructuralPaintMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent structural paint geometry uses installed Windows fonts.");
        CheckEquationRevisions(null, true, "math-source-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionStructuralPaintDiscoveryKeepsActualSourceOwners()
    {
        foreach (string role in new[] { "function-name", "root-body", "delimiter-argument", "nary-only" })
        {
            string fixture = "CompositionProbes/math-source-" + role + "-empty";
            var document = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false);
            var equation = document.Paragraphs.Single(p => p.Equations.Count != 0).Equations.Single();
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var runs), "Known empty literal roles retain bounded discovery.");
            TestAssert.Equal(role == "function-name" ? 1 : 0, runs.Count);
            var control = DocxMathLayout.GetStandalonePaintControl(equation, default);
            if (role == "function-name")
                TestAssert.True(control is null, "An invisible function name retains its actual argument source.");
            else
                TestAssert.True(control is not null && equation.Content.Children.Any(node => ReferenceEquals(node, control)),
                    "Structural paint retains the actual parsed control, without invented source runs.");
            var unknown = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
            {
                var owner = xml.Descendants(m + "oMath").Single().Descendants(m + (role == "function-name" ? "fName" : "e")).First();
                owner.ReplaceNodes(new XElement(m + "future"));
            }).Paragraphs.Single(p => p.Equations.Count != 0).Equations.Single();
            TestAssert.True(!DocxMathLayout.TryGetRuns(unknown, default, out _), "Unknown children cannot masquerade as an empty argument.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetStandalonePaintControl(equation, canceled.Token));
        }
    }

    internal static void CheckStructuralPaintBounds(IFontResolver resolver)
    {
        foreach (string role in new[] { "root-body", "delimiter-argument", "nary-only" })
        {
            var document = ReadMutation("CompositionProbes/math-source-" + role + "-empty", OoxPdfDocxMarkupMode.Final, false, false);
            using (OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 1 }))
                TestAssert.Throws<OoxPdfLimitExceededException>(() => new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderBlankPages(document, null, default).ToList());
        }
    }
}
