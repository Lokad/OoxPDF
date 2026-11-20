using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDrawingSelectsLastPlacementWithinEachDrawing()
    {
        XNamespace wp = WordprocessingDrawingNamespace;
        XNamespace mc = MarkupCompatibilityNamespace;
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            PdfPage[] inline = Render();
            PdfPage[] anchored = Render((drawing, picture) => picture.ReplaceWith(Anchor(picture)));
            TestAssert.Equal(1, CountImagePaints(inline));
            TestAssert.Equal(1, CountImagePaints(anchored));
            TestAssert.True(!inline.Select(page => page.Content).SequenceEqual(anchored.Select(page => page.Content)),
                "The independently framed inline and anchored placements paint differently.");

            AssertSamePaint(inline, Render((drawing, picture) => drawing.AddFirst(NarrowPicture(picture))));
            AssertSamePaint(inline, Render((drawing, picture) => drawing.AddFirst(Anchor(picture))));
            AssertSamePaint(anchored, Render((drawing, picture) => drawing.Add(Anchor(picture))));
            AssertSamePaint(inline, Render((drawing, picture) =>
            {
                drawing.AddFirst(Anchor(picture));
                picture.ReplaceWith(new XElement(mc + "AlternateContent",
                    new XElement(mc + "Choice", new XAttribute(XNamespace.Xmlns + "wp", wp.NamespaceName),
                        new XAttribute("Requires", "wp"), new XElement(picture)),
                    new XElement(mc + "Fallback", Anchor(picture))));
            }));

            PdfPage[] separate = Render((drawing, picture) =>
                drawing.AddBeforeSelf(new XElement(drawing.Name, NarrowPicture(picture))));
            TestAssert.Equal(2, CountImagePaints(separate));

            PdfPage[] Render(Action<XElement, XElement>? mutate = null)
            {
                DocxDocument document = ReadMutation("Composition/report", mode, bookmark: false, split: false,
                    change: (source, w, _) =>
                    {
                        XElement drawing = source.Descendants(w + "drawing").Single();
                        mutate?.Invoke(drawing, drawing.Elements(wp + "inline").Single());
                    });
                var geometry = mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                return new DocxRenderer(resolver, mode, geometry)
                    .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            }
        }

        XElement NarrowPicture(XElement source)
        {
            var picture = new XElement(source);
            var extent = picture.Element(wp + "extent")!;
            extent.SetAttributeValue("cx", "228600");
            extent.SetAttributeValue("cy", "114300");
            picture.Element(wp + "docPr")!.SetAttributeValue("id", "710");
            return picture;
        }

        XElement Anchor(XElement source)
        {
            XElement picture = NarrowPicture(source);
            return new XElement(wp + "anchor",
                new XAttribute("simplePos", "0"), new XAttribute("relativeHeight", "251658240"),
                new XAttribute("behindDoc", "0"), new XAttribute("locked", "0"),
                new XAttribute("layoutInCell", "1"), new XAttribute("allowOverlap", "1"),
                new XElement(wp + "simplePos", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(wp + "positionH", new XAttribute("relativeFrom", "page"),
                    new XElement(wp + "posOffset", "6350000")),
                new XElement(wp + "positionV", new XAttribute("relativeFrom", "page"),
                    new XElement(wp + "posOffset", "1397000")),
                new XElement(picture.Element(wp + "extent")!), new XElement(wp + "wrapNone"), picture.Elements().Skip(1));
        }

        static int CountImagePaints(IEnumerable<PdfPage> pages) => pages.Sum(page =>
            DocxNoteClippingTests.ReadPhysicalPaintFrames(page).Count(frame =>
                page.Images.Any(image => frame.Block == $"/{image.ResourceName} Do")));

        static void AssertSamePaint(PdfPage[] expected, PdfPage[] actual)
        {
            TestAssert.True(expected.Select(page => (page.Width, page.Height, page.Content))
                    .SequenceEqual(actual.Select(page => (page.Width, page.Height, page.Content))),
                "One drawing paints exactly its last resolved placement and preserves surrounding content.");
            TestAssert.Equal(1, CountImagePaints(actual));
        }
    }
}
