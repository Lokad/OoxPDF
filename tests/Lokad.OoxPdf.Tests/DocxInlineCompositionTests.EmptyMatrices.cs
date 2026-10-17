using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEmptyMatricesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent empty matrix frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "empty-matrix-office.json", splitMathRuns: true,
            includeCase: id => !id.Contains("-deleted-", StringComparison.Ordinal));
        // Counterfactual whole deletions keep their existing output. The 18pt
        // Review control's inherited baseline failure stays in raw full gates.
        CheckEquationRevisions(null, false, "empty-matrix-office.json", splitMathRuns: true,
            includeCase: id => id.Contains("-deleted-", StringComparison.Ordinal));
        CheckEmptyMatrixFrames(null);
    }

    public static void DocxInlineCompositionEmptyMatricesKeepActualSourceOwners()
    {
        var source = ReadMutation("CompositionProbes/math-matrix-authored-empty-11", OoxPdfDocxMarkupMode.AllMarkup,
            false, false, (xml, w, m) =>
            {
                var matrix = xml.Descendants(m + "m").First();
                matrix.ReplaceWith(new XElement(w + "ins", new XAttribute(w + "id", "982"),
                    new XAttribute(w + "author", "Empty Grid Reviewer"), new XElement(matrix)));
            });
        var paragraph = source.Paragraphs.First();
        var equation = paragraph.Equations.Single();
        TestAssert.True(equation.HasRunRevisions && DocxMathLayout.TryGetRuns(equation, default, out var runs) && runs.Count == 0 &&
            DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) && owners.Count == 1 &&
            DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
            DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph),
            "An empty grid retains its actual control owner while remaining neutral without revised literal ink.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.IsEmptyMatrix(equation, canceled.Token));
    }

    internal static void CheckEmptyMatrixFrames(IFontResolver? resolver)
    {
        foreach (string profile in new[] { "authored-empty-18", "implicit-empty-11" })
        {
            var source = ReadMutation("CompositionProbes/math-matrix-" + profile, OoxPdfDocxMarkupMode.Final,
                false, false, (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    body.ReplaceNodes(new XElement(body.Element(w + "p")!), new XElement(body.Element(w + "sectPr")!));
                });
            var equation = source.Paragraphs.First().Equations.Single();
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var runs) && runs.Count == 0 &&
                equation.Content.Children.Single() is DocxMathMatrix { Rows.Count: > 0 } matrix &&
                matrix.Rows.All(row => row.Cells.All(cell => cell.Children.Count == 0)),
                "Actual empty slots remain valid source structure.");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(source, diagnostics.Add, default).Single();
            TestAssert.True(page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
                "Metric-only grids publish no placeholder glyph, mathematical subset or dummy source run.");
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == "/word/document.xml"),
                "A supported empty grid retains its measured inline frame instead of an ignored-equation fallback.");
        }
    }
}
