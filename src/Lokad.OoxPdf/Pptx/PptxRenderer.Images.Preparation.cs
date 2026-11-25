using System.Globalization;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Keep decode results across picture owners. A speculative raster preview
    // can select SVG without consuming the decode diagnostic.
    internal sealed class RasterPicture(PdfImageXObject? image, Func<int, OoxPdfDiagnostic>? pendingDiagnostic = null)
    {
        public PdfImageXObject? Image { get; } = image;

        public void ReportDiagnostic(Action<OoxPdfDiagnostic>? sink, int slideIndex)
        {
            if (sink is null || pendingDiagnostic is not { } create) return;
            pendingDiagnostic = null;
            sink(create(slideIndex));
        }
    }

    private static PptxSceneSlide PreparePictureSources(PptxSceneSlide slide,
        Dictionary<string, RasterPicture> cache, Action<OoxPdfDiagnostic>? diagnostics, int slideNumber, CancellationToken token)
    {
        var master = slide.ShowMasterShapes ? Nodes(slide.MasterNodes, false, slide.MasterPartName) : slide.MasterNodes;
        var layout = Nodes(slide.LayoutNodes, false, slide.LayoutPartName);
        var body = Nodes(slide.SlideNodes, true, slide.PartName);
        return ReferenceEquals(master, slide.MasterNodes) && ReferenceEquals(layout, slide.LayoutNodes) && ReferenceEquals(body, slide.SlideNodes)
            ? slide : slide with { MasterNodes = master, LayoutNodes = layout, SlideNodes = body };

        IReadOnlyList<PptxSceneNode> Nodes(IReadOnlyList<PptxSceneNode> nodes, bool placeholders, string? sourcePart)
        {
            List<PptxSceneNode>? changed = null;
            for (int i = 0; i < nodes.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var node = nodes[i];
                var prepared = node;
                if (ShouldRenderSceneNode(node, placeholders))
                {
                    var children = node.Kind == PptxSceneNodeKind.Group ? Nodes(node.Children, placeholders, sourcePart) : node.Children;
                    var picture = node.Picture is { } image && node.Bounds is not null ? Picture(image, sourcePart) : node.Picture;
                    if (!ReferenceEquals(children, node.Children) || !ReferenceEquals(picture, node.Picture))
                        prepared = PptxSceneNode.Create(node.Kind, node.Id, node.Name, node.IsPlaceholder, node.IsSmartArtGraphicFrame,
                            node.HyperlinkClick, node.Bounds, node.Shape, node.TextBody, picture, node.Table, node.Chart, node.GroupTransform, children, node.Source);
                }
                if (!ReferenceEquals(prepared, node)) changed ??= nodes.Take(i).ToList();
                changed?.Add(prepared);
            }
            return changed is null ? nodes : changed;
        }

        PptxScenePicture Picture(PptxScenePicture picture, string? sourcePart)
        {
            if (picture.SvgFallback is not { } svg || !svg.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase) ||
                picture.Resource?.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase) == true) return picture;
            if (picture.Resource is { } preferred)
            {
                var crop = ToCropRect(picture.Crop);
                if (PrepareRasterPicture(preferred, picture.Recolor, cache, diagnostics, slideNumber, token,
                    ref crop, selectingPreview: true) is not null)
                    return picture;
            }
            token.ThrowIfCancellationRequested();
            string id = picture.Resource is not null ? "IMAGE_UNSUPPORTED_FORMAT"
                : picture.TargetPartName is null ? "IMAGE_MISSING_RELATIONSHIP" : "IMAGE_MISSING_PART";
            diagnostics?.Invoke(new OoxPdfDiagnostic(id, OoxPdfSeverity.Warning,
                "The preferred picture source was unavailable or undecodable; using its authored embedded SVG.",
                picture.TargetPartName ?? sourcePart ?? slide.PartName, SlideIndex: slideNumber,
                Feature: "image", Fallback: "Embedded SVG"));
            return picture with { RelationshipId = picture.SvgFallbackRelationshipId, TargetPartName = svg.PartName, Resource = svg };
        }
    }

    private static PdfImageXObject? PrepareRasterPicture(PptxSceneImageResource resource,
        PptxSceneImageRecolor recolor, Dictionary<string, RasterPicture>? cache,
        Action<OoxPdfDiagnostic>? diagnostics, int slideIndex, CancellationToken token, ref CropRect crop,
        bool selectingPreview = false)
    {
        RasterPicture? cropped = null;
        if (!crop.IsEmpty)
        {
            cropped = GetOrCreateCroppedImage(resource, recolor, cache, token, crop);
            if (cropped.Image is { } image)
            {
                crop = default;
                return image;
            }
            if (!selectingPreview)
                cropped.ReportDiagnostic(diagnostics, slideIndex);
        }

        RasterPicture original = GetOrCreateImage(resource, recolor, cache, token);
        if (selectingPreview && original.Image is not null)
            cropped?.ReportDiagnostic(diagnostics, slideIndex);
        if (!selectingPreview || original.Image is not null)
            original.ReportDiagnostic(diagnostics, slideIndex);
        return original.Image;
    }

    private static RasterPicture GetOrCreateCroppedImage(
        PptxSceneImageResource imageResource, PptxSceneImageRecolor recolor,
        Dictionary<string, RasterPicture>? imageCache, CancellationToken cancellationToken, CropRect crop)
    {
        string cacheKey = imageResource.PartName + "\u001f" + ImageRecolorCacheKey(recolor) + "\u001fcrop:" +
            crop.Left.ToString("R", CultureInfo.InvariantCulture) + "," +
            crop.Top.ToString("R", CultureInfo.InvariantCulture) + "," +
            crop.Right.ToString("R", CultureInfo.InvariantCulture) + "," +
            crop.Bottom.ToString("R", CultureInfo.InvariantCulture);
        if (imageCache is not null && imageCache.TryGetValue(cacheKey, out RasterPicture? cached))
        {
            return cached;
        }

        RasterPicture croppedImage = CreateCroppedImage();
        imageCache?.TryAdd(cacheKey, croppedImage);
        return croppedImage;

        RasterPicture CreateCroppedImage()
        {
            try
            {
                // D02: shared pixel decoding; unsupported types keep the silent
                // PDF-clipping fallback (no diagnostic), exactly as before.
                if (!OoxImageDecoder.IsSupportedContentType(imageResource.ContentType))
                {
                    return new(null);
                }

                using DecodedPixels decoded = OoxImageDecoder.DecodePixelsOwned(imageResource.ContentType, imageResource.Bytes, cancellationToken);
                return new(CreateCroppedRgbImage(decoded.Width, decoded.Height, decoded.Rgb, decoded.Alpha, recolor, crop, cancellationToken));
            }
            catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or IndexOutOfRangeException)
            {
                string partName = imageResource.PartName;
                return new(null, slideIndex => new OoxPdfDiagnostic(
                    "IMAGE_CROP_UNSUPPORTED_FORMAT",
                    OoxPdfSeverity.Warning,
                    $"Image '{partName}' could not be decoded for Office-style cropped image embedding on slide {slideIndex}; falling back to PDF clipping.",
                    partName,
                    PageIndex: null,
                    SlideIndex: slideIndex,
                    Feature: "image crop",
                    Fallback: "PDF clipping"));
            }
        }
    }
}
