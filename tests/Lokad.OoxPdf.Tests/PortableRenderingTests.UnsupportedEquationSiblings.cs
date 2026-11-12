using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;
using System.Text.RegularExpressions;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingUnsupportedRevisedEquationsKeepSupportedSiblings()
    {
        foreach (string fixture in new[] { "CompositionProbes/math-run-revision-body-ins", "InlineBoxMath/inline-box-math-cell" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "ins", "del" })
        {
            var beforeSource = DocxInlineCompositionTests.ReadUnsupportedEquationSibling(fixture, mode, null);
            var afterSource = DocxInlineCompositionTests.ReadUnsupportedEquationSibling(fixture, mode, kind);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var afterDiagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                string Math(PdfPage value) => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(value, true)
                    .Where(show => value.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(font.Font, show.Font))).Select(show => show.Text));
                string expected = Math(before[page]), actual = Math(after[page]);
                string context = $"{fixture}/{mode}/{kind}/page{page + 1}";
                TestAssert.True(expected == actual, $"{context}: expected mathematical text '{expected}', got '{actual}'.");
                string TextPaint(PdfPage value) => string.Join("\n", Regex.Matches(value.Content,
                    @"(?ms)^BT\r?\n.*?^ET\r?$").Select(match => match.Value));
                TestAssert.True(TextPaint(before[page]) == TextPaint(after[page]),
                    context + ": supported text retains its exact paint, positions, fonts and spacing.");
                if (mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.Original)
                    TestAssert.True(before[page].Content == after[page].Content,
                        context + ": an omitted equation cannot change unmarked page paint.");
                TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
                PdfContentValidator.ValidatePage(after[page], page, default);
            }
            TestAssert.True(beforeDiagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION") &&
                beforeDiagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION") ==
                    afterDiagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"),
                "The added accent remains unsupported; supported neighboring formulas keep their own content.");
        }
    }
}
