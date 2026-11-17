using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathCenterGroupTests
{
    public static void DocxMathDocumentDefaultsReachStoryAndTextboxSources()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer", "header-cell", "footnote", "endnote", "inlinebox", "floatingbox" })
        {
            var document = Read(null, "right", container: container);
            var paragraph = Paragraph(document);
            TestAssert.True(paragraph.Equations.All(e => e.AlignmentValue is null && e.DefaultAlignmentValue == "right"),
                "Document alignment reaches its source owner without manufacturing local properties: " + container);
            TestAssert.True(paragraph.Equations.All(e => ReferenceEquals(e.LimitSettings, document.Settings.MathLimitSettings)),
                "Every source owner borrows the same immutable document math defaults: " + container);
        }
    }

    public static void DocxMathInheritedAlignmentMatchesLocalPaintAcrossContainers()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string container in new[] { "body", "cell", "header", "footer", "header-cell", "inlinebox" })
        foreach (string alignment in new[] { "left", "right" })
        {
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final ?
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var inherited = renderer.RenderPagesAsync(Read(null, alignment, container: container, markupMode: mode), null, default)
                .ToBlockingEnumerable().Single();
            var local = renderer.RenderPagesAsync(Read(alignment, container: container, markupMode: mode), null, default)
                .ToBlockingEnumerable().Single();
            TestAssert.Equal(local.Content, inherited.Content);
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(inherited, true, includeTextTransforms: true).Select(show => show.Text))
                .Normalize(System.Text.NormalizationForm.FormKC);
            TestAssert.True(text.Contains('a') && text.Contains('b') && text.Contains('c'),
                "Equivalent placement must retain all three equations: " + container + "/" + alignment);
        }
    }

    public static void DocxMathLocalAlignmentOverridesDocumentDefaultsAcrossContainers()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer", "header-cell", "footnote", "endnote", "inlinebox", "floatingbox" })
        {
            var document = Read("left", "right", container: container);
            TestAssert.True(Paragraph(document).Equations.All(e => e.DefaultAlignmentValue == "right" &&
                e.AlignmentValue == "left" && e.EffectiveAlignmentValue == "left"),
                "Local alignment keeps precedence and both source values remain inspectable: " + container);
        }
    }
}
