using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionGlyphlessCompoundsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent compound metric frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "glyphless-compounds-office.json", splitMathRuns: true,
            includeRow: row => !UnchangedVisibleDeletion(row));
        // Both visible Review PDFs are exact parent identities. Their inherited
        // 0.108pt rule-width failures remain in the raw full geometry report.
        CheckEquationRevisions(null, false, "glyphless-compounds-office.json", splitMathRuns: true,
            includeRow: UnchangedVisibleDeletion);

        static bool UnchangedVisibleDeletion(System.Text.Json.JsonElement row) =>
            row.GetProperty("Id").GetString()!.EndsWith("-deleted", StringComparison.Ordinal) &&
            row.GetProperty("Mode").GetString() == "all-word-compatible";
    }

    public static void DocxInlineCompositionGlyphlessCompoundsKeepOrderedActualOwners()
    {
        foreach (string order in new[] { "script-first", "fraction-first" })
        foreach (string kind in new[] { "deleted", "authored-empty", "empty-ins" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/glyphless-compound-" + order + "-" + kind, mode, false, false);
            var paragraph = source.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            var controls = DocxMathLayout.GetControlSources(equation, default);
            bool visible = kind == "deleted" && mode == OoxPdfDocxMarkupMode.AllMarkup;
            bool hasRuns = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            TestAssert.True(controls.Count == 2 && hasRuns && runs.Count == (visible ? 6 : 0) &&
                controls.All(control => control.ControlText?.SourceRunIndex == equation.SourceRunIndex) &&
                (order == "script-first" ? controls[0] is DocxMathScripts : controls[0] is DocxMathFraction) &&
                DocxMathLayout.GetMetricOnlyControls(equation, runs, default).Count == (visible ? 0 : 2) &&
                DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) &&
                owners.Count == (visible || kind == "empty-ins" ? 1 : 0) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
                DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph) != visible,
                "Ordered actual controls retain metric and revision ownership independently of literal run presence.");
            if (kind == "empty-ins")
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
                    { Equations = [equation with { Content = new([]) }] }, default),
                    "Aggregate revision metadata cannot replace removed actual controls.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetControlSources(equation, canceled.Token));
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetMetricOnlyControls(equation, runs, canceled.Token));
        }
    }

    public static void DocxInlineCompositionGlyphlessCompoundsKeepMetricOnlyPublication()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native compound metric publication uses installed Windows fonts.");
        CheckGlyphlessCompoundPublication(null);
    }

    internal static void CheckGlyphlessCompoundPublication(IFontResolver? resolver)
    {
        foreach (string order in new[] { "script-first", "fraction-first" })
        foreach (string kind in new[] { "deleted", "authored-empty", "empty-ins" })
        {
            var mode = kind == "empty-ins" ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var source = ReadMutation("CompositionProbes/glyphless-compound-" + order + "-" + kind, mode, false, false,
                (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    body.ReplaceNodes(new XElement(body.Element(w + "p")!), new XElement(body.Element(w + "sectPr")!));
                });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(source, diagnostics.Add, default).Single();
            TestAssert.True(page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)) &&
                !diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == "/word/document.xml"),
                "Multiple glyphless controls retain occupied frames and structural paint without dummy glyphs or math subsets.");
        }
    }
}
