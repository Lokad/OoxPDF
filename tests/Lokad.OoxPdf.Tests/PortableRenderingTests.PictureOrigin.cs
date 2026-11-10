using System.Xml.Linq;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingPictureOnlyNestedTablesKeepTheirBodyOrigin()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int markSize in new[] { 12, 48 })
        {
            (double X, double Top)? expected = null;
            foreach (int depth in new[] { 0, 1, 2, 3 })
            {
                var document = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var run = body.Descendants(w + "drawing").Single().Parent!;
                    run.Remove();
                    foreach (var extent in run.Descendants().Where(e => e.Name.LocalName is "extent" or "ext" && e.Attribute("cy") is not null))
                    {
                        extent.SetAttributeValue("cx", 360 * 12700);
                        extent.SetAttributeValue("cy", 180 * 12700);
                    }
                    var leading = new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "rPr",
                        new XElement(w + "sz", new XAttribute(w + "val", markSize * 2)))), run);
                    for (int level = 0; level < depth; level++) leading = Table(leading, level != 0);
                    body.AddFirst(new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page")))));
                    body.AddFirst(leading);

                    XElement Table(XElement child, bool nested) => new(w + "tbl",
                        new XElement(w + "tblPr",
                            new XElement(w + "tblW", new XAttribute(w + "w", 9360), new XAttribute(w + "type", "dxa")),
                            new XElement(w + "tblLayout", new XAttribute(w + "type", "fixed")),
                            new XElement(w + "tblCellMar", new[] { "top", "left", "bottom", "right" }.Select(side =>
                                new XElement(w + side, new XAttribute(w + "w", 0), new XAttribute(w + "type", "dxa"))))),
                        new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", 9360))),
                        new XElement(w + "tr", new XElement(w + "tc",
                            new XElement(w + "tcPr", new XElement(w + "tcW", new XAttribute(w + "w", 9360), new XAttribute(w + "type", "dxa"))),
                            child, nested ? new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
                                new XAttribute(w + "before", 0), new XAttribute(w + "after", 0),
                                new XAttribute(w + "line", 20), new XAttribute(w + "lineRule", "exact")))) : null)));
                });
                var page = RenderAlignmentProbe(document, mode)[0];
                var actual = DocxInlineCompositionTests.ReadPaintedPictureFrame(page);
                expected ??= actual;
                TestAssert.True(Math.Abs(expected.Value.X - actual.X) < .002d && Math.Abs(expected.Value.Top - actual.Top) < .002d,
                    $"Picture-only body origin: {mode}, mark={markSize}, depth={depth}, expected={expected.Value}, actual={actual}.");
                var header = ReadPaintedSourceLabel(page, "a+b=c", mathOnly: true);
                TestAssert.True(header.All(glyph => glyph.Ink.Y > actual.Top),
                    $"The opaque body picture cannot overpaint the source header equation: {mode}, mark={markSize}, depth={depth}.");
            }
        }
    }
}
