using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingEmptyRevisedMathRunsKeepFractions()
    {
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            string fixture = "CompositionProbes/math-run-revision-body-" + kind;
            var beforeSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string[] beforeGlyphs = before.Select(ReadPortableEquationGlyphs).ToArray();
            TestAssert.Equal(RevisionBodyEquationText(kind, mode), string.Concat(beforeGlyphs));

            foreach (string contents in new[] { "", "<w:lastRenderedPageBreak/>", "<m:rPr><m:nor m:val='0'/></m:rPr>" })
            {
                var afterSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, w, m) =>
                {
                    var revision = xml.Descendants(w + kind).Single();
                    revision.Add(XElement.Parse($"<m:r xmlns:m='{m}' xmlns:w='{w}'>{contents}</m:r>"));
                });
                var afterDiagnostics = new List<OoxPdfDiagnostic>();
                var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(before.Length, after.Length);
                for (int page = 0; page < before.Length; page++)
                {
                    TestAssert.Equal(beforeGlyphs[page], ReadPortableEquationGlyphs(after[page]));
                    TestAssert.Equal(before[page].Content, after[page].Content);
                    TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                    TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                    TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
                    PdfContentValidator.ValidatePage(after[page], page, default);
                }
                TestAssert.True(beforeDiagnostics.SequenceEqual(afterDiagnostics),
                    "An empty readable run cannot change the revision's supported visible mathematics.");
            }
        }
    }

    private static string RevisionBodyEquationText(string kind, OoxPdfDocxMarkupMode mode)
    {
        // The first fraction has a revised numerator. The other equations are unmarked.
        bool numeratorVisible = kind == "ins" ? mode != OoxPdfDocxMarkupMode.Original :
            mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
        return (numeratorVisible ? "x+2" : "") + "ya+b=cx+2ya+b=c";
    }

    private static string ReadPortableEquationGlyphs(PdfPage page) => string.Concat(
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, honorActualText: false,
            fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).Select(show => show.Text))
        .Normalize(NormalizationForm.FormKC);

    public static void PortableRenderingMathRunPaginationHintsKeepComposedContent()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var source in new[]
        {
            (Fixture: "InlineBoxMath/inline-sibling-mixed", Part: "word/document.xml"),
            (Fixture: "InlineBoxMath/inline-box-math-cell", Part: "word/document.xml"),
            (Fixture: "InlineBoxMath/inline-box-math-header", Part: "word/inline-box-header.xml"),
            (Fixture: "CompositionProbes/math-note-content-footnote-inline", Part: "word/footnotes.xml")
        })
        {
            var beforeSource = DocxInlineCompositionTests.ReadMutation(source.Fixture, mode, false, false, partOverride: source.Part);
            var afterSource = DocxInlineCompositionTests.ReadMutation(source.Fixture, mode, false, false, (xml, w, m) =>
            {
                var runs = xml.Descendants(m + "r").ToArray();
                TestAssert.True(runs.Length != 0, "The source contains mathematical runs.");
                foreach (var run in runs) run.Add(new XElement(w + "lastRenderedPageBreak"));
            }, partOverride: source.Part);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var afterDiagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            TestAssert.True(beforeDiagnostics.SequenceEqual(afterDiagnostics), "Stored pagination markers do not change equation support diagnostics.");
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(before[page].Content, after[page].Content);
                TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
                PdfContentValidator.ValidatePage(after[page], page, default);
            }
        }
    }
}
