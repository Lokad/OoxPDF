using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxPairedScriptsKeepPaintWhenAlignmentIsDisabled()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = MathContentControlRenderer(mode);
            foreach (bool inHeaderCell in new[] { false, true })
            {
                var (control, baselineDiagnostics) = Render(inHeaderCell, null);
                var (actual, diagnostics) = Render(inHeaderCell, inHeaderCell ? "false" : "0");
                var glyphs = DocxCommentsTests.ReadEmbeddedGlyphTextShows(control.Single(), usePageCoordinates: true,
                    fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
                TestAssert.Equal(3, glyphs.Length);
                TestAssert.Equal("2", glyphs[2].Text);
                TestAssert.True(glyphs[1].Y < glyphs[0].Y && glyphs[0].Y < glyphs[2].Y,
                    "The base must paint between the subscript and superscript baselines.");
                TestAssert.True(diagnostics.SequenceEqual(baselineDiagnostics), "A disabled property keeps the control diagnostics.");
                AssertMathContentControlPages(control, actual);
            }

            var (enabled, enabledDiagnostics) = Render(false, "default");
            TestAssert.True(!enabled.Single().Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
                "A present flag without a value requests alignment and retains the unsupported fallback.");
            TestAssert.True(enabledDiagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"),
                "The active flag has its own diagnosed fallback.");

            (PdfPage[] Pages, List<OoxPdfDiagnostic> Diagnostics) Render(bool inHeaderCell, string? flag)
            {
                string fixture = inHeaderCell ? "TableGridOrigins/Containers/header-zero" : "docx-omml-scripts-inline";
                string part = inHeaderCell ? "word/static-inline-ordinary-header-box.xml" : "word/document.xml";
                var source = ReadMutation(fixture, mode, false, false, (xml, _, m) =>
                {
                    if (inHeaderCell)
                    {
                        xml.Descendants(m + "oMath").Single().ReplaceNodes(new XElement(m + "sSubSup",
                            Argument("e", "x"), Argument("sub", "i"), Argument("sup", "2")));
                    }
                    if (flag is null) return;
                    var alignment = new XElement(m + "alnScr");
                    if (flag != "default") alignment.SetAttributeValue(m + "val", flag);
                    xml.Descendants(m + "sSubSup").Single().AddFirst(new XElement(m + "sSubSupPr", alignment));

                    XElement Argument(string name, string value) =>
                        new(m + name, new XElement(m + "r", new XElement(m + "t", value)));
                }, part);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                return (pages, diagnostics);
            }
        }
    }
}
