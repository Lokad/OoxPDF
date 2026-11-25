using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMathSymbolsKeepFontFractionInkAndNavigation()
    {
        const string family = "DejaVu Sans";
        var resolver = new MathSymbolResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool header in new[] { false, true })
        {
            string denominator = header ? "y" : "pyq";
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var direct = Render("literal");
            var symbol = Render("symbol");
            CheckFraction(direct.Page);
            CheckFraction(symbol.Page);
            TestAssert.Equal(direct.Page.Content, symbol.Page.Content);
            TestAssert.True(direct.Diagnostics.SequenceEqual(symbol.Diagnostics),
                "An owned symbol and its explicit normal-text control must describe the same rendered content.");
            TestAssert.True(direct.Page.Annotations.SequenceEqual(symbol.Page.Annotations),
                "The symbol must retain its formula's link rectangles and destinations.");
            if (!header)
                TestAssert.Equal(8, symbol.Page.Annotations.Count);

            var malformed = Render("malformed");
            TestAssert.True(malformed.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"),
                "A malformed math symbol keeps the existing diagnosed equation fallback.");
            TestAssert.True(!malformed.Page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && font.Font.Font.FamilyName == family),
                "The malformed symbol must not cause font preparation for that equation.");

            (PdfPage Page, List<string> Diagnostics) Render(string variant)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var source = ReadMutation(header ? "CompositionProbes/table-math-revision-header-row-ins"
                    : "CompositionProbes/inline-math-navigation-cell", mode, false, false, (xml, w, m) =>
                    {
                        var fraction = xml.Descendants(w + "tbl").First().Descendants(m + "f").First();
                        fraction.Element(m + "num")!.Descendants(m + "t").Single().Value = "z+7";
                        var run = fraction.Element(m + "den")!.Element(m + "r")!;
                        if (variant == "literal")
                        {
                            if (!header) run.AddBeforeSelf(new XElement(m + "r", new XElement(m + "t", "p")));
                            if (!header) run.AddAfterSelf(new XElement(m + "r", new XElement(m + "t", "q")));
                            run.ReplaceNodes(new XElement(m + "rPr", new XElement(m + "nor")),
                                new XElement(w + "rPr", new XElement(w + "rFonts", new XAttribute(w + "ascii", family))),
                                new XElement(m + "t", "y"));
                        }
                        else
                        {
                            var symbolElement = new XElement(w + "sym", new XAttribute(w + "font", family),
                                new XAttribute(w + "char", variant == "malformed" ? "ZZZZ" : "0079"));
                            if (header) run.ReplaceNodes(symbolElement);
                            else run.ReplaceNodes(new XElement(m + "t", "p"), symbolElement, new XElement(m + "t", "q"));
                        }
                    }, partOverride: header ? "word/header1.xml" : "word/document.xml", diagnosticSink: Record);
                return (renderer.RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single(), diagnostics);
            }

            void CheckFraction(PdfPage page)
            {
                var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, "z+7" + denominator, mathOnly: true);
                TestAssert.Equal(3 + denominator.Length, glyphs.Count);
                var rectangles = ReadVisibleFilledRectangles(page).ToArray();
                ReadVisibleFractionRule(rectangles, glyphs[0].Origin, glyphs[2].Origin, glyphs[3].Origin, glyphs[^1].Origin);
                var shown = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, honorActualText: false,
                    fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal));
                TestAssert.Equal(family, shown.Single(show => show.Text == "y").Font.Font.FamilyName);
            }
        }
    }

    private sealed class MathSymbolResolver : IFontResolver
    {
        private readonly PortableRenderingTests.PinnedResolver math = new(mathFont: true);
        private readonly PortableRenderingTests.PinnedResolver prose = new();

        public FontFaceResolution Resolve(FontRequest request) =>
            request.FamilyName == "DejaVu Sans" ? prose.Resolve(request) : math.Resolve(request);
    }
}
