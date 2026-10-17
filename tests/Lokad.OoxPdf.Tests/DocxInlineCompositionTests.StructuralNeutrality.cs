using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStructuralOnlyRevisionsRetainNeutralityAndFallbacks()
    {
        foreach (string role in new[] { "root", "delimiter", "operator" })
        foreach (string kind in new[] { "ins", "del" })
        {
            var document = ReadMutation("CompositionProbes/whole-math-empty-" + role + "-" + kind,
                OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var paragraph = document.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            TestAssert.True(DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "Actual structural ownership remains valid without revised literal ink.");
            foreach (var invalid in new[] { equation with { Revision = equation.Revision! with { } },
                equation with { SourceRunIndex = -1 }, equation with { Content = new([]) } })
                TestAssert.True(!DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph with { Equations = [invalid] }),
                    "Unowned, value-equal and unknown-position metadata cannot become neutral.");
            var visible = ReadMutation("CompositionProbes/whole-math-" + role + "-" + kind,
                OoxPdfDocxMarkupMode.AllMarkup, false, false).Paragraphs.First();
            TestAssert.True(!DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(visible),
                "An actual revised literal source still contributes a visible revision band.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph, token: canceled.Token));
        }
    }

    public static void DocxInlineCompositionStructuralOnlyRevisionsMatchIndependentWordNeutrality()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent structural revision controls use installed Windows fonts.");
        CheckStructuralOnlyRevisionNeutrality(null);
    }

    internal static void CheckStructuralOnlyRevisionNeutrality(IFontResolver? resolver)
    {
        foreach (string role in new[] { "root", "delimiter", "operator" })
        foreach (string kind in new[] { "ins", "del" })
        {
            string fixture = "CompositionProbes/whole-math-empty-" + role + "-" + kind;
            var source = ReadMutation(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var neutral = ReadMutation(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
            {
                var wrapper = xml.Descendants(m + "oMath").First().Parent!;
                wrapper.ReplaceWith(wrapper.Nodes());
            });
            var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderBlankPages(source, null, default).Single();
            var after = renderer.RenderBlankPages(neutral, null, default).Single();
            TestAssert.Equal(after.Content, before.Content);
        }
    }
}
