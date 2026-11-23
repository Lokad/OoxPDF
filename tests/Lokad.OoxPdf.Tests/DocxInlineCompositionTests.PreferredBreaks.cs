using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
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
                    var (paragraph, drawing) = CreateExactPictureCell(xml, w, 2360, 600, 280, "Following body.");
                    XElement Run(params object[] content) => PictureCellRun(w, content);
                    XElement Text(string value) => PictureCellText(w, value);
                    paragraph.Add(Run(Text("Review "), drawing, Text(" ")));
                    if (hardBreak)
                        paragraph.Add(Run(Text("scope-"), new XElement(w + "br"), Text("schedule followup")));
                    else if (split)
                        paragraph.Add(Run(Text("scope")), Run(Text("-schedule followup")));
                    else
                        paragraph.Add(Run(Text("scope-schedule followup")));
                });
                return renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();
            }
        }
    }

    public static void DocxHangingContinuationKeepsPrecedingFractionAndPicture()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            PdfPage manual = Paint(authoredSplit: true);
            PdfPage wrapped = Paint(authoredSplit: false);
            foreach (var (label, mathOnly) in new[] { ("437", true), ("293", true), ("RENDERINGORDER", false), ("After table.", false) })
            {
                var expected = PortableRenderingTests.ReadPaintedSourceLabel(manual, label, mathOnly);
                var actual = PortableRenderingTests.ReadPaintedSourceLabel(wrapped, label, mathOnly);
                TestAssert.True(expected.SequenceEqual(actual), "Complete glyph ink matches authored placement for " + label);
            }
            var continuation = PortableRenderingTests.ReadPaintedSourceLabel(wrapped, "RENDERINGORDER");
            TestAssert.True(continuation[0].Origin.Y > continuation[^1].Origin.Y,
                "The word wraps inside the narrower continuation rather than rejecting the preceding math.");
            TestAssert.Equal(ReadPaintedPictureFrame(manual), ReadPaintedPictureFrame(wrapped));
            TestAssert.Equal(manual.Images.Single().Image.ResourceKey, wrapped.Images.Single().Image.ResourceKey);

            PdfPage Paint(bool authoredSplit)
            {
                var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, m) =>
                {
                    var (paragraph, drawing) = CreateExactPictureCell(xml, w, 2800, 1320, 320, "After table.");
                    var properties = paragraph.Element(w + "pPr")!;
                    properties.AddFirst(new XElement(w + "wordWrap", new XAttribute(w + "val", "0")));
                    properties.Add(new XElement(w + "ind", new XAttribute(w + "left", "1340"), new XAttribute(w + "hanging", "1340")));
                    AddPictureFractionLabel(paragraph, drawing, w, m);
                    paragraph.Add(PictureCellRun(w, new XElement(w + "br")));
                    if (authoredSplit)
                        paragraph.Add(PictureCellRun(w, PictureCellText(w, "RENDERING"), new XElement(w + "br")),
                            PictureCellRun(w, PictureCellText(w, "ORDER")));
                    else
                        paragraph.Add(PictureCellRun(w, PictureCellText(w, "RENDERINGORDER")));
                });
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, CachedMathGeometry(mode))
                    .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
                TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                    diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "The original fraction is placed.");
                return page;
            }
        }
    }

    public static void DocxFittingDecimalTabsRetainMathPicturesAndNumericAnchors()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            PdfPage right = Paint("right", "123");
            PdfPage integer = Paint("decimal", "123");
            TestAssert.Equal(right.Content, integer.Content);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(integer, "123").Count);

            foreach (var (number, parts) in new[] { ("12.34", new[] { "12.34" }), ("12,34", new[] { "12", ",34" }) })
            {
                PdfPage page = Paint("decimal", parts);
                var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, number);
                var label = PortableRenderingTests.ReadPaintedSourceLabel(page, "Evidence");
                TestAssert.Equal(5, glyphs.Count);
                TestAssert.True(Math.Abs(glyphs[2].Origin.X - label[0].Origin.X - 132d) < .02d,
                    "The decimal separator paints at the authored 132pt tab stop.");
                foreach (string value in new[] { "437", "293" })
                    TestAssert.True(PortableRenderingTests.ReadPaintedSourceLabel(right, value, mathOnly: true).SequenceEqual(
                        PortableRenderingTests.ReadPaintedSourceLabel(page, value, mathOnly: true)),
                        "Decimal positioning must preserve the preceding fraction's complete glyph paint.");
                TestAssert.Equal(ReadPaintedPictureFrame(right), ReadPaintedPictureFrame(page));
                TestAssert.Equal(right.Images.Single().Image.ResourceKey, page.Images.Single().Image.ResourceKey);
            }

            PdfPage Paint(string alignment, params string[] number)
            {
                var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, m) =>
                {
                    var (paragraph, drawing) = CreateExactPictureCell(xml, w, 3200, 1320, 320, "After table.");
                    paragraph.Element(w + "pPr")!.AddFirst(new XElement(w + "tabs", new XElement(w + "tab",
                        new XAttribute(w + "val", alignment), new XAttribute(w + "pos", "2640"))));
                    AddPictureFractionLabel(paragraph, drawing, w, m);
                    paragraph.Add(PictureCellRun(w, new XElement(w + "tab")));
                    foreach (string part in number)
                        paragraph.Add(PictureCellRun(w, PictureCellText(w, part)));
                });
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, CachedMathGeometry(mode))
                    .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
                var equation = diagnostics.Single(d => d.Feature == "equation");
                TestAssert.True(equation.Id == "DOCX_APPROXIMATED_EQUATION",
                    $"The {alignment} stop keeps the fraction beside {string.Concat(number)} in {mode}. {equation.Message}");
                return page;
            }
        }
    }

    private static void AddPictureFractionLabel(XElement paragraph, XElement drawing, XNamespace w, XNamespace m)
    {
        XElement FractionPart(string kind, string value) =>
            new(m + kind, new XElement(m + "r", new XElement(m + "t", value)));
        paragraph.Add(PictureCellRun(w, PictureCellText(w, "Evidence ")), PictureCellRun(w, drawing),
            new XElement(m + "oMath", new XElement(m + "f", FractionPart("num", "437"), FractionPart("den", "293"))));
    }

    private static (XElement Paragraph, XElement Drawing) CreateExactPictureCell(XDocument xml, XNamespace w,
        int width, int height, int lineSpacing, string followingText)
    {
        var body = xml.Root!.Element(w + "body")!;
        var section = body.Element(w + "sectPr")!;
        var drawing = body.Descendants(w + "drawing").Single();
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (var extent in drawing.Descendants(wp + "extent").Concat(drawing.Descendants(a + "ext")
            .Where(e => e.Parent?.Name == a + "xfrm")))
        {
            extent.SetAttributeValue("cx", 279400);
            extent.SetAttributeValue("cy", 139700);
        }
        var table = XElement.Parse($$"""
            <w:tbl xmlns:w="{{w}}">
              <w:tblPr><w:tblW w:w="{{width}}" w:type="dxa"/><w:tblLayout w:type="fixed"/>
                <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>
                  <w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
              </w:tblPr>
              <w:tblGrid><w:gridCol w:w="{{width}}"/></w:tblGrid>
              <w:tr><w:trPr><w:trHeight w:val="{{height}}" w:hRule="exact"/></w:trPr>
                <w:tc><w:tcPr><w:tcW w:w="{{width}}" w:type="dxa"/></w:tcPr>
                  <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{{lineSpacing}}" w:lineRule="exact"/></w:pPr></w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
            """);
        body.ReplaceNodes(table, new XElement(w + "p", PictureCellRun(w, PictureCellText(w, followingText))), section);
        return (table.Descendants(w + "p").Single(), drawing);
    }

    private static XElement PictureCellRun(XNamespace w, params object[] content) => new(w + "r",
        new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "24"))), content);

    private static XElement PictureCellText(XNamespace w, string value) =>
        new(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value);
}
