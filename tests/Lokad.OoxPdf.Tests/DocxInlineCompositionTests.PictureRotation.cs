using System.Xml.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
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
            // Check the actual emitted quarter-turn and raster rectangle
            // against its enclosing line clip, independently of text presence.
            const string number = @"([-\d.]+)";
            var draw = Regex.Match(page.Content, number + " " + number + " " + number + " " + number +
                @" re\s+W\s+n\s+0 -1 1 0 " + number + " " + number + @" cm\s+q\s+" + number +
                " 0 0 " + number + " " + number + " " + number + @" cm\s+/" + Regex.Escape(page.Images.Single().ResourceName) + @" Do");
            TestAssert.True(draw.Success, "The rotated raster keeps its explicit measured line clip.");
            double Value(int index) => double.Parse(draw.Groups[index].Value, CultureInfo.InvariantCulture);
            double left = Value(10) + Value(5), right = left + Value(8);
            double top = Value(6) - Value(9), bottom = top - Value(7);
            TestAssert.True(left >= Value(1) - .001d && right <= Value(1) + Value(3) + .001d &&
                bottom >= Value(2) - .001d && top <= Value(2) + Value(4) + .001d,
                "Rotated picture ink must remain inside its declared measured frame in " + container + "/" + mode);
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
