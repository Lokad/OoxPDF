using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void RenderImage(DocxInlineImage image, double x, double y, double width, double height,
        PdfGraphicsBuilder graphics, List<PdfImageResource> pageImages, Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageIndex, CancellationToken cancellationToken, Dictionary<string, PdfImageXObject?> imageCache, ref int imageIndex)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
        {
            Action<OoxPdfDiagnostic>? svgSink = diagnosticSink is null ? null : diagnostic =>
                diagnosticSink(image.RasterFallback is not null && diagnostic.Severity == OoxPdfSeverity.Error && diagnostic.Fallback == "Ignored"
                    ? diagnostic with { Severity = OoxPdfSeverity.Warning, Fallback = "RasterImage", Message = diagnostic.Message + " The authored raster preview is used." }
                    : diagnostic);
            if (PptxRenderer.TryRenderSvgImage(graphics, x, y, width, height, image.Bytes, svgSink, pageIndex, image.PartName, cancellationToken)) return;
            if (image.RasterFallback is not { } preview) return;
            image = preview;
        }
        PdfImageXObject? resource = CreateImage(image, imageCache, diagnosticSink, pageIndex, cancellationToken);
        if (resource is null) return;
        string name = "Im" + imageIndex++;
        graphics.DrawImage(name, x, y, width, height);
        pageImages.Add(new PdfImageResource(name, resource));
    }
}
