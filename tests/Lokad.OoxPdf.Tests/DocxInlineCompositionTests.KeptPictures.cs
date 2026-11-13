using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxKeptFallbackPicturesShareOccupiedLinesWithPlacement()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool minimum in new[] { false, true })
        foreach (int pictures in new[] { 1, 2 })
        foreach (bool boundary in new[] { false, true })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var control = ReadKeptFallbackPicture(mode, minimum, pictures, boundary, keep: false, counter: false);
            var source = ReadKeptFallbackPicture(mode, minimum, pictures, boundary, keep: true, counter: false);
            var expected = renderer.RenderPagesAsync(control, null, default).ToBlockingEnumerable().Single();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, actual.Length);
            TestAssert.Equal(expected.Content, actual[0].Content);
            TestAssert.Equal(pictures, actual[0].Images.Count);
            TestAssert.Equal(1, diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"));
            AssertSourceLabelsPaint(actual[0], mode, ["KEPT PICTURE HEADING", "Before", "after picture."], "kept fallback pictures");
        }
    }

    public static void DocxKeptFallbackPicturesBindCountersAfterOccupiedAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadKeptFallbackPicture(mode, minimum: false, pictures: 2, boundary: true, keep: true, counter: true);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            TestAssert.Contains("PAGE 1", DocxPageFootnoteNumberingTests.DecodeText(pages[0]));
            TestAssert.Equal(2, pages[0].Images.Count);
            TestAssert.Equal(1, diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"));
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxDocument ReadKeptFallbackPicture(OoxPdfDocxMarkupMode mode, bool minimum, int pictures,
        bool boundary, bool keep, bool counter) => ReadMutation("docx-picture-rotated-raster", mode, false, false,
        (xml, w, m) =>
        {
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            var body = xml.Root!.Element(w + "body")!;
            var paragraph = body.Element(w + "p")!;
            var runs = paragraph.Elements(w + "r").ToArray();
            runs[0].Element(w + "t")!.Value = "Before ";
            runs[2].Element(w + "t")!.Value = " after picture.";
            paragraph.Element(w + "pPr")?.Remove();
            var spacing = new XElement(w + "spacing", new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                new XAttribute(w + "line", minimum ? "400" : "240"), new XAttribute(w + "lineRule", minimum ? "atLeast" : "auto"));
            paragraph.AddFirst(new XElement(w + "pPr", spacing));
            var drawing = runs[1].Element(w + "drawing")!;
            drawing.Descendants(a + "xfrm").Single().Attribute("rot")!.Remove();
            foreach (var extent in drawing.Descendants().Where(node => node.Name == a + "ext" || node.Name == wp + "extent"))
            { extent.SetAttributeValue("cx", 100 * 12700); extent.SetAttributeValue("cy", 50 * 12700); }
            for (int index = 1; index < pictures; index++)
            {
                var sibling = new XElement(runs[1]); sibling.Descendants(wp + "docPr").Single().SetAttributeValue("id", index + 1);
                runs[1].AddAfterSelf(sibling);
            }
            // Keep an unsupported neighbor on the established fallback while
            // testing the complete, otherwise supported text/picture frame.
            paragraph.Add(new XElement(m + "oMath", new XElement(m + "acc", new XElement(m + "e",
                new XElement(m + "r", new XElement(m + "t", "q"))))));
            var heading = new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "keepNext", new XAttribute(w + "val", keep ? "1" : "0")),
                new XElement(w + "spacing", new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                    new XAttribute(w + "line", "240"), new XAttribute(w + "lineRule", "exact"))),
                new XElement(w + "r", new XElement(w + "t", "KEPT PICTURE HEADING")));
            paragraph.AddBeforeSelf(heading);
            if (boundary) heading.AddBeforeSelf(new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
                new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                new XAttribute(w + "line", "11500"), new XAttribute(w + "lineRule", "exact")))));
            if (counter) paragraph.Add(new XElement(w + "r", new XElement(w + "t", " PAGE ")),
                new XElement(w + "fldSimple", new XAttribute(w + "instr", " PAGE "), new XElement(w + "r", new XElement(w + "t", "9"))));
            foreach (var run in body.Descendants(w + "r"))
            {
                run.Element(w + "rPr")?.Remove();
                run.AddFirst(new XElement(w + "rPr", new XElement(w + "rFonts", new XAttribute(w + "ascii", "Arial"),
                    new XAttribute(w + "hAnsi", "Arial")), new XElement(w + "sz", new XAttribute(w + "val", "24"))));
            }
            var section = body.Element(w + "sectPr")!;
            section.Element(w + "pgSz")!.SetAttributeValue(w + "w", "12240");
            section.Element(w + "pgSz")!.SetAttributeValue(w + "h", "15840");
            foreach (string edge in new[] { "top", "bottom", "left", "right" }) section.Element(w + "pgMar")!.SetAttributeValue(w + edge, "1440");
        });
}
