using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingInlineBoxFallbackKeepsOutsideSiblings()
    {
        foreach (string fixture in new[] { "inline-box-body-equation-inline-box-valid", "inline-sibling-mixed" })
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = DocxInlineCompositionTests.ReadMutation("InlineBoxMath/" + fixture, mode, false, false);
            var source = ReadNestedBoxFallback(fixture, mode);
            var expected = renderer.RenderPagesAsync(control, null, default).ToBlockingEnumerable().Single();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
            var expectedMath = MathShows(expected); var actualMath = MathShows(page);
            TestAssert.Equal(fixture == "inline-sibling-mixed" ? 8 : 4, actualMath.Length);
            TestAssert.Equal(expectedMath.Length, actualMath.Length);
            for (int i = 0; i < actualMath.Length; i++)
                TestAssert.True(actualMath[i].Text == expectedMath[i].Text && actualMath[i].Size == expectedMath[i].Size &&
                    Math.Abs(actualMath[i].X - expectedMath[i].X) < .001d && Math.Abs(actualMath[i].Y - expectedMath[i].Y) < .001d,
                    "An unsupported box child cannot change a supported outside formula's painted source frame.");
            string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
            foreach (string prose in new[] { "Before inline", "after equation.", "Inline box content.", "After public math table." })
                TestAssert.Contains(prose, text);
            TestAssert.True(!text.Contains("Public nested frame.", StringComparison.Ordinal),
                "Nested inline box content retains its existing ignored fallback.");
            TestAssert.Equal(fixture == "inline-sibling-mixed" ? 1 : 0, page.Images.Count);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Rendered outside equations must not retain an ignored-equation diagnostic.");
            var warning = diagnostics.Single(d => d.Id == "DOCX_UNSUPPORTED_NESTED_INLINE_TEXTBOX");
            TestAssert.Equal("Ignored", warning.Fallback);
            TestAssert.Equal(0, warning.PageIndex ?? -1);
            PdfContentValidator.ValidatePage(page, 0, default);
        }

        var failed = ReadNestedBoxFallback("inline-sibling-mixed", OoxPdfDocxMarkupMode.Final);
        var observerRenderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(),
            OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        bool observed = false;
        TestAssert.Throws<InvalidOperationException>(() => observerRenderer.RenderPagesAsync(failed, diagnostic =>
        {
            if (diagnostic.Id == "DOCX_UNSUPPORTED_NESTED_INLINE_TEXTBOX")
            {
                observed = true;
                throw new InvalidOperationException("Expected diagnostic observer failure.");
            }
        }, default).ToBlockingEnumerable().ToArray());
        TestAssert.True(observed, "The nested-box diagnostic reaches the supplied observer.");

        static (PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] MathShows(PdfPage page) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).ToArray();
    }

    private static DocxDocument ReadNestedBoxFallback(string fixture, OoxPdfDocxMarkupMode mode) =>
        DocxInlineCompositionTests.ReadMutation("InlineBoxMath/" + fixture, mode, false, false, (xml, w, _) =>
        {
            var outer = xml.Descendants(w + "txbxContent").Single();
            var nested = new XElement(outer.Ancestors(w + "drawing").First());
            var content = nested.Descendants(w + "txbxContent").Single();
            content.RemoveNodes();
            content.Add(new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "Public nested frame."))));
            int id = 3000;
            foreach (var node in nested.Descendants().Where(node => node.Name.LocalName is "docPr" or "cNvPr"))
            {
                if (node.Attribute("id") is not null) node.SetAttributeValue("id", id++);
                if (node.Attribute("name") is not null) node.SetAttributeValue("name", "PublicNested" + id);
            }
            outer.Elements(w + "p").First().Add(new XElement(w + "r", nested));
        });
}
