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
        DocxPictureCrop crop = image.Crop;
        bool hasFlip = image.FlipHorizontal || image.FlipVertical;
        if (crop.IsEmpty && !hasFlip)
        {
            RenderImageCore(image, x, y, width, height, graphics, pageImages, diagnosticSink,
                pageIndex, cancellationToken, imageCache, ref imageIndex);
            return;
        }
        double sourceWidth = width / (1d - crop.Left - crop.Right);
        double sourceHeight = height / (1d - crop.Top - crop.Bottom);
        double sourceX = x - crop.Left * sourceWidth;
        double sourceY = y - crop.Bottom * sourceHeight;
        if (!crop.IsUsable || !double.IsFinite(sourceWidth) || !double.IsFinite(sourceHeight) ||
            !double.IsFinite(sourceX) || !double.IsFinite(sourceY))
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic("DOCX_UNSUPPORTED_IMAGE_CROP", OoxPdfSeverity.Warning,
                "The picture source rectangle is unusable; the uncropped image is retained.", image.PartName,
                PageIndex: pageIndex, Feature: "image crop", Fallback: "UncroppedImage"));
            sourceX = x; sourceY = y; sourceWidth = width; sourceHeight = height;
            if (!hasFlip)
            {
                RenderImageCore(image, x, y, width, height, graphics, pageImages, diagnosticSink,
                    pageIndex, cancellationToken, imageCache, ref imageIndex);
                return;
            }
        }
        // Crop changes paint inside the authored viewport, never its layout extent.
        // The same mapping applies to native vectors and their raster fallback.
        graphics.SaveState();
        try
        {
            graphics.ClipRectangle(x, y, width, height);
            if (hasFlip)
            {
                // Mirror the cropped paint around the authored viewport. Existing
                // static-line clips remain in page coordinates outside this state.
                graphics.Transform(image.FlipHorizontal ? -1d : 1d, 0d, 0d, image.FlipVertical ? -1d : 1d,
                    image.FlipHorizontal ? 2d * x + width : 0d, image.FlipVertical ? 2d * y + height : 0d);
            }
            RenderImageCore(image, sourceX, sourceY, sourceWidth, sourceHeight, graphics, pageImages,
                diagnosticSink, pageIndex, cancellationToken, imageCache, ref imageIndex);
        }
        finally { graphics.RestoreState(); }
    }

    private static void RenderImageCore(DocxInlineImage image, double x, double y, double width, double height,
        PdfGraphicsBuilder graphics, List<PdfImageResource> pageImages, Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageIndex, CancellationToken cancellationToken, Dictionary<string, PdfImageXObject?> imageCache, ref int imageIndex)
    {
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
