using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrderedPictureRotationKeepsMeasuredFramesAcrossContainers()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Read(bool rotate) => ReadMutation("CompositionProbes/" + container, mode, false, false,
                (xml, w, _) =>
                {
                    if (!rotate) return;
                    var drawing = xml.Descendants(w + "drawing").Single();
                    var extent = drawing.Descendants(wp + "extent").Single();
                    string width = extent.Attribute("cx")!.Value, height = extent.Attribute("cy")!.Value;
                    extent.SetAttributeValue("cx", height);
                    extent.SetAttributeValue("cy", width);
                    drawing.Descendants(a + "xfrm").Single().SetAttributeValue("rot", "5400000");
                }, partOverride: container is "header" or "footer" ? "word/" + container + "1.xml" : null);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(Read(false), null, default).ToBlockingEnumerable().Single();
            var source = Read(true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(1, page.Images.Count);
            TestAssert.Contains("0 -1 1 0 ", page.Content);
            TestAssert.True(diagnostics.All(d => d.Id is not ("DOCX_UNSUPPORTED_IMAGE_TRANSFORM" or "DOCX_UNSUPPORTED_EQUATION")),
                "Supported rotation and mathematics share the same placed line frame in " + container + "/" + mode);
            static string Characters(Lokad.OoxPdf.Pdf.PdfPage value) => string.Concat(
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(value, usePageCoordinates: true)
                    .SelectMany(show => show.Text).Where(character => !char.IsWhiteSpace(character)).Order());
            TestAssert.Equal(Characters(before), Characters(page));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }
}
