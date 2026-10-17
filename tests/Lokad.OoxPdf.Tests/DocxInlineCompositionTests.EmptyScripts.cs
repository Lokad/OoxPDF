using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEmptyScriptsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent empty script frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "empty-scripts-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionEmptyScriptsKeepActualOwnersAndCancellation()
    {
        foreach (string profile in EmptyScriptProfiles)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/empty-script-" + profile, mode, false, false);
            var paragraph = source.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            bool visible = profile == "combined-literal-del" && mode == OoxPdfDocxMarkupMode.AllMarkup;
            TestAssert.True(DocxMathLayout.IsEmptyScripts(equation, default) != visible &&
                ReferenceEquals(DocxMathLayout.GetStandaloneControl(equation, default), equation.Content.Children.Single()) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
                DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph) != visible &&
                DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) &&
                owners.Count == (profile == "combined-empty-ins" ? 1 : visible ? 3 : 0),
                "Actual implicit controls retain empty slots while visible deleted literals keep their own Review owners.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.IsEmptyScripts(equation, canceled.Token));
        }
    }

    public static void DocxInlineCompositionEmptyScriptsKeepMetricOnlyPublication()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native empty script frames use installed Windows fonts.");
        CheckEmptyScriptPublication(null);
    }

    private static readonly string[] EmptyScriptProfiles =
        ["sub-empty", "sup-empty", "combined-empty", "combined-literal-del", "combined-empty-ins"];

    internal static void CheckEmptyScriptPublication(IFontResolver? resolver)
    {
        foreach (string profile in EmptyScriptProfiles)
        {
            var mode = profile == "combined-empty-ins" ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var source = ReadMutation("CompositionProbes/empty-script-" + profile, mode, false, false,
                (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    body.ReplaceNodes(new XElement(body.Element(w + "p")!), new XElement(body.Element(w + "sectPr")!));
                });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(source, diagnostics.Add, default).Single();
            TestAssert.True(page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
                "Empty scripts use only metric owners, without placeholder glyphs, dummy source runs or math subsets.");
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == "/word/document.xml"),
                "A supported empty script retains its occupied frame instead of an ignored-equation fallback.");
        }
    }
}
