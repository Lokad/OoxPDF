using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionGlyphlessFractionArgumentsKeepIndependentWordRulesAndContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent glyphless argument rules use installed Windows fonts.");
        CheckEquationRevisions(null, false, "glyphless-fraction-arguments-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionGlyphlessFractionArgumentsKeepActualMetricOwners()
    {
        foreach (string style in new[] { "inline", "display" })
        foreach (string role in new[] { "numerator", "denominator" })
        {
            var document = ReadMutation("CompositionProbes/glyphless-fraction-" + style + "-" + role,
                OoxPdfDocxMarkupMode.Final, false, false);
            var equation = document.Paragraphs.First().Equations.Single();
            bool supported = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            var controls = DocxMathLayout.GetControlSources(equation, default);
            var metrics = DocxMathLayout.GetMetricOnlyControls(equation, runs, default);
            TestAssert.True(supported && runs.All(source => source.Run.Text.Text.Length == 0) && controls.Count == 2 &&
                controls[0] is DocxMathFraction && controls[1] is DocxMathScripts &&
                metrics.SequenceEqual(controls),
                "The parent fraction and nested script retain their actual metric owners without literal sources.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetMetricOnlyControls(equation, runs, canceled.Token));
        }
    }

    public static void DocxInlineCompositionGlyphlessFractionArgumentsKeepMetricOnlyPublicationAndEmptyRuns()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native glyphless argument publication uses installed Windows fonts.");
        CheckGlyphlessFractionPublication(null);
    }

    internal static void CheckGlyphlessFractionPublication(IFontResolver? resolver)
    {
        foreach (string style in new[] { "inline", "display" })
        foreach (string role in new[] { "numerator", "denominator" })
        {
            string? original = null;
            foreach (bool emptyRuns in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/glyphless-fraction-" + style + "-" + role,
                    OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
                    {
                        var body = xml.Root!.Element(w + "body")!;
                        var paragraph = body.Element(w + "p")!;
                        if (emptyRuns)
                            foreach (var argument in paragraph.Descendants(m + "f").First().Elements()
                                .Where(child => child.Name == m + "num" || child.Name == m + "den"))
                            {
                                var target = argument.Element(m + "sSubSup")?.Element(m + "e") ?? argument;
                                target.Add(new XElement(m + "r", new XElement(m + "t", string.Empty)));
                            }
                        body.ReplaceNodes(new XElement(paragraph), new XElement(body.Element(w + "sectPr")!));
                    });
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderBlankPages(document, diagnostics.Add, default).Single();
                TestAssert.True(page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)) &&
                    Regex.IsMatch(page.Content, @"(?m)^0 g\r?\n[-\d.]+ [-\d.]+ [\d.]+ [\d.]+ re f\r?$") &&
                    !diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == "/word/document.xml"),
                    "The visible bar and occupied frames use metric fonts without dummy glyphs or math subsets.");
                if (original is not null) TestAssert.Equal(original, page.Content);
                original = page.Content;
            }
        }
    }
}
