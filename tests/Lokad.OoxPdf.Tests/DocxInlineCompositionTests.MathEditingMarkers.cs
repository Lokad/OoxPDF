using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMathEditingMarkersKeepPortablePaintAcrossContainers()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var source in new[]
        {
            (Fixture: "docx-omml-fraction-inline", Part: "word/document.xml", Page: 0,
                Numerator: "a+b", Denominator: "c", Companion: "", Count: 1),
            (Fixture: "CompositionProbes/math-picture-first-table-single-cell", Part: "word/document.xml", Page: 0,
                Numerator: "rate", Denominator: "time", Companion: "", Count: 2),
            (Fixture: "CompositionProbes/equation-revision-header-ins", Part: "word/header1.xml", Page: 0,
                Numerator: "x+2", Denominator: "y", Companion: "a+b=c", Count: 2),
            (Fixture: "CompositionProbes/mixed-note-endnote-table", Part: "word/endnotes.xml", Page: 1,
                Numerator: "rate", Denominator: "time", Companion: "", Count: 2)
        })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var direct = DocxTests.ReadDocx(TestFixtures.PublicCasePath(source.Fixture + ".docx"), mode);
            var marked = ReadMutation(source.Fixture, mode, false, false, (xml, w, m) =>
            {
                var arguments = xml.Descendants(m + "num").ToArray();
                TestAssert.Equal(source.Count, arguments.Length);
                int markerId = 12;
                foreach (var argument in arguments)
                {
                    argument.AddFirst(new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")),
                        new XElement(w + "permStart", new XAttribute(w + "id", markerId),
                            new XAttribute(w + "edGrp", "everyone")));
                    argument.Add(new XElement(w + "permEnd", new XAttribute(w + "id", markerId)),
                        new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd")));
                    markerId++;
                }
            }, partOverride: source.Part);
            var before = renderer.RenderPagesAsync(direct, null, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(marked, null, default).ToBlockingEnumerable().ToArray();
            CheckEditingMarkerFractionPaint(before[source.Page], source.Numerator, source.Denominator, source.Companion, source.Count);
            CheckEditingMarkerFractionPaint(after[source.Page], source.Numerator, source.Denominator, source.Companion, source.Count);
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(before[page].Content, after[page].Content);
                TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
            }
        }
    }

    private static void CheckEditingMarkerFractionPaint(PdfPage page, string numerator, string denominator, string companion, int count)
    {
        string equation = numerator + denominator + companion;
        string expected = string.Concat(Enumerable.Repeat(equation, count));
        var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, expected, mathOnly: true);
        TestAssert.Equal(expected.Length, glyphs.Count);
        var rectangles = ReadVisibleFilledRectangles(page).ToArray();
        for (int i = 0; i < count; i++)
        {
            int start = i * equation.Length;
            int denominatorStart = start + numerator.Length;
            ReadVisibleFractionRule(rectangles, glyphs[start].Origin, glyphs[denominatorStart - 1].Origin,
                glyphs[denominatorStart].Origin, glyphs[denominatorStart + denominator.Length - 1].Origin);
        }
    }
}
