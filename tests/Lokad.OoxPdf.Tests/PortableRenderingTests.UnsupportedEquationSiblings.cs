using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;
using System.Text.RegularExpressions;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingUnsupportedRevisedEquationsKeepSupportedSiblings()
    {
        foreach (var fixture in new[]
        {
            (Path: "CompositionProbes/math-run-revision-body-ins", Body: true),
            (Path: "InlineBoxMath/inline-box-math-cell", Body: false)
        })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var beforeSource = DocxInlineCompositionTests.ReadUnsupportedEquationSibling(fixture.Path, mode, null);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string[] beforeGlyphs = before.Select(ReadPortableEquationGlyphs).ToArray();
            TestAssert.Equal(fixture.Body ? RevisionBodyEquationText("ins", mode) : "x+2y", string.Concat(beforeGlyphs));
            string[] beforePaint = before.Select(ReadTextPaint).ToArray();
            int unsupportedCount = beforeDiagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION");
            TestAssert.True(unsupportedCount > 0, "The control accent is unsupported.");

            foreach (string kind in new[] { "ins", "del" })
            {
                var afterSource = DocxInlineCompositionTests.ReadUnsupportedEquationSibling(fixture.Path, mode, kind);
                var afterDiagnostics = new List<OoxPdfDiagnostic>();
                var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(before.Length, after.Length);
                for (int page = 0; page < before.Length; page++)
                {
                    string expected = beforeGlyphs[page], actual = ReadPortableEquationGlyphs(after[page]);
                    string context = $"{fixture.Path}/{mode}/{kind}/page{page + 1}";
                    TestAssert.True(expected == actual, $"{context}: expected mathematical text '{expected}', got '{actual}'.");
                    TestAssert.True(beforePaint[page] == ReadTextPaint(after[page]),
                        context + ": supported text retains its exact paint, positions, fonts and spacing.");
                    if (mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.Original)
                        TestAssert.True(before[page].Content == after[page].Content,
                            context + ": an omitted equation cannot change unmarked page paint.");
                    TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                    TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                    TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
                    PdfContentValidator.ValidatePage(after[page], page, default);
                }
                TestAssert.Equal(unsupportedCount, afterDiagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"));
            }
        }

        static string ReadTextPaint(PdfPage page) => string.Join("\n", Regex.Matches(page.Content,
            @"(?ms)^BT\r?\n.*?^ET\r?$").Select(match => match.Value));
    }
}
