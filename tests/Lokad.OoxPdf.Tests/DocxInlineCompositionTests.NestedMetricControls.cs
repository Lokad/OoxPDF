using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNestedMetricControlsKeepIndependentWordContentAndRules()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent nested metric controls use installed Windows fonts.");
        CheckEquationRevisions(null, false, "nested-metric-controls-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionNestedMetricControlsKeepOrderedActualOwners()
    {
        foreach (string style in new[] { "inline", "display" })
        foreach (string role in new[] { "numerator", "denominator" })
        {
            var document = ReadMutation("CompositionProbes/nested-metric-" + style + "-" + role,
                OoxPdfDocxMarkupMode.Final, false, false);
            var equation = document.Paragraphs.First().Equations.Single();
            bool supported = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            var controls = DocxMathLayout.GetControlSources(equation, default);
            var metrics = DocxMathLayout.GetMetricOnlyControls(equation, runs, default);
            TestAssert.True(supported && runs.Count == 1 && controls.Count == 2 &&
                controls[0] is DocxMathFraction && controls[1] is DocxMathScripts &&
                metrics.Count == 1 && ReferenceEquals(metrics[0], controls[1]) &&
                controls.All(control => control.ControlText?.SourceRunIndex == equation.SourceRunIndex),
                "A literal sibling cannot replace the nested script's actual metric owner.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetControlSources(equation, canceled.Token));
        }
    }

    public static void DocxInlineCompositionNestedMetricControlsKeepMetricOnlyPublication()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native nested metric publication uses installed Windows fonts.");
        CheckNestedMetricPublication(null);
    }

    internal static void CheckNestedMetricPublication(IFontResolver? resolver)
    {
        foreach (string style in new[] { "inline", "display" })
        foreach (string role in new[] { "numerator", "denominator" })
        {
            var document = ReadMutation("CompositionProbes/nested-metric-" + style + "-" + role,
                OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    body.ReplaceNodes(new XElement(body.Element(w + "p")!), new XElement(body.Element(w + "sectPr")!));
                });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(document, diagnostics.Add, default).Single();
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
            var embedded = fonts.Select(font => font.Font).ToHashSet();
            string actual = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .Where(show => embedded.Contains(show.Font)).Select(show => show.Text))
                .Normalize(System.Text.NormalizationForm.FormKC);
            TestAssert.True(fonts.Length == 1 && actual == (role == "numerator" ? "y" : "x+2") &&
                !diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == "/word/document.xml"),
                "Only the visible sibling publishes mathematical glyphs; empty script slots add no dummy glyph or subset.");
        }
    }
}
