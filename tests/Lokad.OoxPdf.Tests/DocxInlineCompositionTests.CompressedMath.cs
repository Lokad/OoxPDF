using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxCompressedDisplayMathRetainsSourceGlyphsAcrossPresets()
    {
        foreach (string kind in new[] { "literal", "fraction" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = PictureRunBoundaryRenderer(mode, mathFont: true);
            var normal = renderer.RenderPagesAsync(ReadCompressedDisplayMath(kind, mode, 240), null, default)
                .ToBlockingEnumerable().ToArray();
            string expected = ZeroAreaMathText(normal);
            TestAssert.Equal(kind == "literal" ? 5 : 4, expected.Normalize(NormalizationForm.FormKC).Length);
            foreach (int line in new[] { 120, 180 })
            {
                var source = ReadCompressedDisplayMath(kind, mode, line);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(expected, ZeroAreaMathText(pages));
                TestAssert.Equal(1, pages.Length);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION") &&
                    diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION"),
                    "Positive compressed spacing retains mathematical meaning with its approximation warning.");
                TestAssert.Contains("Following prose.", string.Concat(pages.SelectMany(page =>
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false)).Select(show => show.Text)));
                using var canceled = new CancellationTokenSource(); canceled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null,
                    canceled.Token).ToBlockingEnumerable().ToArray());
            }
        }
    }

    internal static DocxDocument ReadCompressedDisplayMath(string kind, OoxPdfDocxMarkupMode mode, int line) =>
        ReadMutation("docx-omml-display-spacing-" + kind + "-first-auto360", mode, false, false, (xml, w, m) =>
        {
            var paragraph = xml.Descendants(w + "p").Single(p => p.Element(m + "oMathPara") is not null);
            paragraph.Element(w + "pPr")!.Element(w + "spacing")!.SetAttributeValue(w + "line", line);
        });
}
