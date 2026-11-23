using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionPreferredBreakKeepsTailInsideExactCell()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var control = Paint(hardBreak: true, split: false);
            var wrapped = Paint(hardBreak: false, split: false);
            var split = Paint(hardBreak: false, split: true);
            var prefix = PortableRenderingTests.ReadPaintedSourceLabel(control, "scope-");
            var tail = PortableRenderingTests.ReadPaintedSourceLabel(control, "schedule followup");
            TestAssert.True(prefix[0].Origin.Y > tail[0].Origin.Y, "The hard break reserves the second line for the complete tail.");
            foreach (var actual in new[] { wrapped, split })
            {
                TestAssert.True(prefix.Select(glyph => glyph.Origin).SequenceEqual(
                    PortableRenderingTests.ReadPaintedSourceLabel(actual, "scope-").Select(glyph => glyph.Origin)),
                    "The hyphenated prefix paints at the end of the occupied line.");
                TestAssert.True(tail.Select(glyph => glyph.Origin).SequenceEqual(
                    PortableRenderingTests.ReadPaintedSourceLabel(actual, "schedule followup").Select(glyph => glyph.Origin)),
                    "The complete tail paints inside the same second-line frame as the control.");
                TestAssert.Equal(control.Images.Single().Image.ResourceKey, actual.Images.Single().Image.ResourceKey);
                TestAssert.Equal(PortableRenderingTests.ReadPaintedSourceLabel(control, "Following body.")[0].Origin,
                    PortableRenderingTests.ReadPaintedSourceLabel(actual, "Following body.")[0].Origin);
            }

            PdfPage Paint(bool hardBreak, bool split)
            {
                var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, _) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var section = body.Element(w + "sectPr")!;
                    var drawing = body.Descendants(w + "drawing").Single();
                    XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                    foreach (var ext in drawing.Descendants(wp + "extent").Concat(drawing.Descendants(a + "ext")
                        .Where(e => e.Parent?.Name == a + "xfrm")))
                    {
                        ext.SetAttributeValue("cx", 279400);
                        ext.SetAttributeValue("cy", 139700);
                    }
                    var table = XElement.Parse($$"""
                        <w:tbl xmlns:w="{{w}}">
                          <w:tblPr><w:tblW w:w="2360" w:type="dxa"/><w:tblLayout w:type="fixed"/>
                            <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>
                              <w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
                          </w:tblPr>
                          <w:tblGrid><w:gridCol w:w="2360"/></w:tblGrid>
                          <w:tr><w:trPr><w:trHeight w:val="600" w:hRule="exact"/></w:trPr>
                            <w:tc><w:tcPr><w:tcW w:w="2360" w:type="dxa"/></w:tcPr>
                              <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr></w:p>
                            </w:tc>
                          </w:tr>
                        </w:tbl>
                        """);
                    var paragraph = table.Descendants(w + "p").Single();
                    XElement Run(params object[] content) => new(w + "r",
                        new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "24"))), content);
                    XElement Text(string value) => new(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value);
                    paragraph.Add(Run(Text("Review "), drawing, Text(" ")));
                    if (hardBreak)
                        paragraph.Add(Run(Text("scope-"), new XElement(w + "br"), Text("schedule followup")));
                    else if (split)
                        paragraph.Add(Run(Text("scope")), Run(Text("-schedule followup")));
                    else
                        paragraph.Add(Run(Text("scope-schedule followup")));
                    body.ReplaceNodes(table, new XElement(w + "p", Run(Text("Following body."))), section);
                });
                return renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();
            }
        }
    }
}
