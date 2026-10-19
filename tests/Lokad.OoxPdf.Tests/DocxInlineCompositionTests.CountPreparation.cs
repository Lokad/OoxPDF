using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxCountPreparationReusesUnchangedSourceAcrossFinalCount()
    {
        foreach (bool review in new[] { false, true })
        {
            var page = RenderPreparedCounter(DocxFieldKind.Page, review, warnings: false, capture: true);
            var count = RenderPreparedCounter(DocxFieldKind.NumPages, review, warnings: false, capture: true);
            TestAssert.Equal(0, page.Diagnostics.Count);
            TestAssert.Equal(0, count.Diagnostics.Count);
            TestAssert.True(page.Requests != 0 && page.Requests == count.Requests,
                $"A stable count reuses the same source preparation: PAGE={page.Requests}, NUMPAGES={count.Requests}.");
            TestAssert.Contains("1", string.Concat(count.Pages.Select(DocxPageFootnoteNumberingTests.DecodeText)));
        }
    }

    public static void DocxCountPreparationPreservesWarningsAndSilentPaintOutput()
    {
        foreach (bool review in new[] { false, true })
        {
            var page = RenderPreparedCounter(DocxFieldKind.Page, review, warnings: true, capture: true);
            var count = RenderPreparedCounter(DocxFieldKind.NumPages, review, warnings: true, capture: true);
            var silent = RenderPreparedCounter(DocxFieldKind.NumPages, review, warnings: true, capture: false);
            TestAssert.True(page.Diagnostics.Count != 0, "The Hebrew control requires the existing script approximation diagnostic.");
            TestAssert.True(page.Diagnostics.SequenceEqual(count.Diagnostics), "A count retry preserves each original warning exactly once.");
            TestAssert.True(count.Requests > page.Requests && silent.Requests == page.Requests,
                "Warnings retain full preparation/replay; a silent caller can reuse the same font ownership.");
            TestAssert.Equal(string.Concat(count.Pages.Select(item => item.Content)), string.Concat(silent.Pages.Select(item => item.Content)));
        }
    }

    public static void DocxCountPreparationRespectsConversionLifetimeAndCancellation()
    {
        var source = PreparationCounterDocument(DocxFieldKind.NumPages, warnings: false);
        var resolver = new PreparationCountingResolver(new PortableRenderingTests.PinnedResolver());
        var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var first = renderer.RenderBlankPages(source, null, default).ToArray();
        int firstRequests = resolver.Requests;
        var second = renderer.RenderBlankPages(source, null, default).ToArray();
        TestAssert.True(firstRequests != 0 && resolver.Requests == firstRequests * 2,
            "Font preparation belongs to one conversion, even when source and renderer are reused.");
        TestAssert.Equal(first.Single().Content, second.Single().Content);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(source, null, cancellation.Token).ToArray());
        TestAssert.Equal(firstRequests * 2, resolver.Requests);
        var other = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(source, null, default).Single();
        TestAssert.True(first.Single().Fonts.All(font => font.Font.Font.FamilyName.Contains("DejaVu Sans", StringComparison.Ordinal)) &&
            other.Fonts.All(font => font.Font.Font.FamilyName.Contains("DejaVu Math", StringComparison.Ordinal)),
            "A different resolver retains its own source faces.");
    }

    private static (PdfPage[] Pages, int Requests, List<OoxPdfDiagnostic> Diagnostics) RenderPreparedCounter(
        DocxFieldKind kind, bool review, bool warnings, bool capture)
    {
        var source = PreparationCounterDocument(kind, warnings);
        var resolver = new PreparationCountingResolver(new PortableRenderingTests.PinnedResolver());
        var diagnostics = new List<OoxPdfDiagnostic>();
        var renderer = new DocxRenderer(resolver, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final,
            review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var pages = renderer.RenderBlankPages(source, capture ? diagnostics.Add : null, default).ToArray();
        TestAssert.Equal(1, pages.Length);
        return (pages, resolver.Requests, diagnostics);
    }

    private static DocxDocument PreparationCounterDocument(DocxFieldKind kind, bool warnings)
    {
        var paragraph = DocxTests.CreateDocxLayoutParagraph("Field", 12d, 14d);
        var style = paragraph.Runs[0];
        var counter = style with { Text = DocxFieldText.Placeholder(kind)!, FieldKind = kind };
        paragraph = paragraph with { Runs = warnings ? [style with { Text = "\u05D0 " }, counter] : [counter] };
        return DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);
    }

    private sealed class PreparationCountingResolver(IFontResolver inner) : IFontResolver
    {
        public int Requests { get; private set; }
        public FontFaceResolution Resolve(FontRequest request)
        {
            Requests++;
            return inner.Resolve(request);
        }
    }
}
