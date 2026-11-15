using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFittingNarrowPicturesKeepAuthoredHeightInFixedRows()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (double width in new[] { .25d, .5d, 1d, 36d })
        foreach (bool fixedRow in new[] { false, true })
        {
            var source = ReadPictureFit(container, mode, width, fixedRow);
            var renderer = PictureRunBoundaryRenderer(mode, mathFont: false);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var page = pages.Single();
            var frame = ReadPictureRangeFrame(page);
            TestAssert.True(Math.Abs(frame.Width - width) < .001d && Math.Abs(frame.Height - 18d) < .001d,
                "Width fitting preserves the authored aspect ratio at every positive width: " + container + "/" + mode + "/" + width + "/" + fixedRow);
            TestAssert.Equal(1, page.Images.Count);
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                honorActualText: false).Select(show => show.Text));
            TestAssert.Contains("Before", text);
            TestAssert.Contains("After", text);
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null,
                canceled.Token).ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxFixedPictureFitPreservesZeroWidthHeightAndBoundedRatios()
    {
        foreach (double width in new[] { 0d, .25d, .5d, 1d, 36d, 600d })
        foreach (double scale in new[] { 1d, .75d })
        {
            var source = ReadPictureFit("body", OoxPdfDocxMarkupMode.Final, width, fixedRow: true);
            var paragraph = source.Paragraphs.Single();
            var image = paragraph.Images.Single();
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, null, new TrackingMeasurer(),
                12d, 9d * scale, 12d * scale, scale, default, out var measured), "The captured ordinary fixed-row picture is measurable.");
            var atoms = TestAssert.NotNull(measured);
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 100d, 100d, default, out var wrapped),
                "The existing bounded width fit must consume source once.");
            var lines = TestAssert.NotNull(wrapped);
            var fragment = lines.SelectMany(line => line.Fragments).Single(fragment => fragment.Source is DocxInlinePicture);
            TestAssert.True(ReferenceEquals(image, ((DocxInlinePicture)fragment.Source).Image), "Fitting retains the original picture owner.");
            double ratio = width == 0d ? scale : Math.Min(scale, 100d / width);
            TestAssert.True(Math.Abs(fragment.Advance - width * ratio) < 1e-9d &&
                Math.Abs(fragment.BaselineInset - 18d * ratio) < 1e-9d,
                "Width admission scales both extents together; zero width retains the other authored extent.");
            TestAssert.True(lines[^1].EndCursor.IsComplete(atoms), "A nonpainting or oversized frame cannot stall the source cursor.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryWrapOrderedInlineAtoms(
                atoms, 100d, 100d, canceled.Token, out _));
        }
    }

    private static DocxDocument ReadPictureFit(string container, OoxPdfDocxMarkupMode mode, double width,
        bool fixedRow) => ReadMutation("docx-picture-rotated-raster", mode, false, false, (xml, w, _) =>
        {
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            var paragraph = xml.Descendants(w + "p").Single();
            paragraph.Descendants(a + "xfrm").Single().Attribute("rot")!.Remove();
            foreach (var extent in paragraph.Descendants().Where(node => node.Name == a + "ext" || node.Name == wp + "extent"))
            {
                extent.SetAttributeValue("cx", (long)(width * 12700d));
                extent.SetAttributeValue("cy", 18 * 12700);
            }
            if (fixedRow) paragraph.AddFirst(new XElement(w + "pPr", new XElement(w + "spacing",
                new XAttribute(w + "line", "240"), new XAttribute(w + "lineRule", "exact"))));
            if (container == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "9360"))),
                new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
        });
}
