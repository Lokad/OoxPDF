using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingLeadingEmptyAndEquationPagesKeepStaticInk()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false), mode);
            foreach (bool equationPage in new[] { false, true })
            {
                var source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var equation = equationPage ? new XElement(body.Descendants(m + "oMath").First()) : null;
                    body.AddFirst(new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page")))));
                    if (equation is not null) body.AddFirst(new XElement(w + "p", new XElement(m + "oMathPara", equation)));
                });
                var pages = RenderAlignmentProbe(source, mode);
                TestAssert.Equal(baseline.Length + 1, pages.Length);
                foreach (var page in pages)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var expected = ReadPaintedSourceLabel(baseline[0], footer);
                    var actual = ReadPaintedSourceLabel(page, footer);
                    TestAssert.Equal(expected.Count, actual.Count);
                    AssertSamePaintOrigin(expected[0], actual[0]);
                }
                const string header = "Quarterly operations report";
                var expectedHeader = ReadPaintedSourceLabel(baseline[^1], header);
                var actualHeader = ReadPaintedSourceLabel(pages[0], header);
                AssertSamePaintOrigin(expectedHeader[0], actualHeader[0]);
                if (equationPage)
                {
                    var bodyMath = ReadPaintedSourceLabel(pages[0], "x+2y", mathOnly: true);
                    var headerMath = ReadPaintedSourceLabel(pages[0], "a+b=c", mathOnly: true);
                    TestAssert.True(bodyMath.Max(glyph => glyph.Origin.Y) < headerMath.Min(glyph => glyph.Origin.Y),
                        "A standalone body equation remains below the independent header equation.");
                }
            }
        }
    }

    public static void PortableRenderingColumnFallbackKeepsBodyAndFooterInk()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false), mode);
            foreach (bool tableFirst in new[] { false, true })
            {
                var source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var section = body.Element(w + "sectPr")!;
                    section.Element(w + "cols")?.Remove();
                    section.Add(new XElement(w + "cols", new XAttribute(w + "num", 2), new XAttribute(w + "space", 360)));
                    var table = body.Element(w + "tbl")!;
                    var properties = table.Element(w + "tblPr")!;
                    properties.Element(w + "tblW")?.Remove();
                    properties.Element(w + "tblLayout")?.Remove();
                    properties.Add(new XElement(w + "tblW", new XAttribute(w + "w", 4320), new XAttribute(w + "type", "dxa")),
                        new XElement(w + "tblLayout", new XAttribute(w + "type", "fixed")));
                    foreach (var column in table.Element(w + "tblGrid")!.Elements()) column.SetAttributeValue(w + "w", 2160);
                    foreach (var cell in table.Descendants(w + "tc"))
                    {
                        var cellProperties = cell.Element(w + "tcPr");
                        if (cellProperties is null) { cellProperties = new XElement(w + "tcPr"); cell.AddFirst(cellProperties); }
                        cellProperties.Element(w + "tcW")?.Remove();
                        cellProperties.Add(new XElement(w + "tcW", new XAttribute(w + "w", 2160), new XAttribute(w + "type", "dxa")));
                    }
                    if (tableFirst) { table.Remove(); body.AddFirst(table); }
                });
                // Balancing remains unsupported. Its fallback must still paint
                // the unchanged page-wide footer within the original frame.
                var pages = RenderAlignmentProbe(source, mode);
                foreach (var page in pages)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var expected = ReadPaintedSourceLabel(baseline[0], footer);
                    var actual = ReadPaintedSourceLabel(page, footer);
                    TestAssert.Equal(expected.Count, actual.Count);
                    AssertSamePaintOrigin(expected[0], actual[0]);
                }
                string bodyParagraph = source.Paragraphs.Select(p => string.Concat(p.Runs.Select(r => r.Text)))
                    .Single(text => text.StartsWith("Review section 6.", StringComparison.Ordinal));
                var owner = pages.Single(page => DocxPageFootnoteNumberingTests.DecodeText(page).Contains("Review section 6.", StringComparison.Ordinal));
                ReadPaintedSourceLabel(owner, bodyParagraph);
            }
        }
    }

    public static void PortableRenderingLeadingBodyContainersKeepStaticFooterInk()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false), mode);
            foreach (bool tableFirst in new[] { true, false })
            {
                var source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    XElement leading;
                    if (tableFirst)
                    {
                        leading = body.Elements(w + "tbl").Single();
                        leading.Remove();
                    }
                    else
                    {
                        var run = body.Descendants(w + "drawing").Single().Parent!;
                        TestAssert.Equal(w + "r", run.Name);
                        run.Remove();
                        leading = new XElement(w + "p", run);
                    }
                    body.AddFirst(leading);
                });
                var pages = RenderAlignmentProbe(source, mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var expected = ReadPaintedSourceLabel(baseline[index], footer);
                    var actual = ReadPaintedSourceLabel(pages[index], footer);
                    TestAssert.Equal(expected.Count, actual.Count);
                    AssertSamePaintOrigin(expected[0], actual[0]);
                }
            }
        }
    }
}
