using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
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
