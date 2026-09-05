using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderPicture(
        PptxSceneNode picture,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        GroupTransform transform,
        List<PdfImageResource> images,
        ref int index)
    {
        if (picture.Picture is null || picture.Bounds is null)
        {
            return;
        }

        RenderPicture(
            context.Document,
            graphics,
            context.DiagnosticSink,
            context.SlideNumber,
            transform,
            images,
            context.ImageCache,
            ref index,
            picture.Picture.TargetPartName,
            picture.Picture.Resource,
            ToShapeBounds(picture.Bounds),
            ToCropRect(picture.Picture.Crop),
            ToFillRect(picture.Picture.Fill),
            picture.Picture.Alpha,
            picture.Picture.Recolor,
            ToLineStyle(picture.Picture.Line),
            ToOuterShadow(picture.Picture.OuterShadow),
            context.CancellationToken);
    }

    private static void RenderPicture(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform transform,
        List<PdfImageResource> images,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int index,
        string? targetPartName,
        PptxSceneImageResource? imageResource,
        ShapeBounds rawBounds,
        CropRect crop,
        FillRect fillRect,
        double alpha,
        PptxSceneImageRecolor recolor,
        LineStyle line,
        OuterShadow? outerShadow,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (targetPartName is null)
        {
            return;
        }

        if (imageResource is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                targetPartName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return;
        }

        ShapeBounds transformedBounds = transform.Apply(rawBounds);
        double x = OoxUnits.EmuToPoints(transformedBounds.X);
        double yTop = OoxUnits.EmuToPoints(transformedBounds.Y);
        double width = OoxUnits.EmuToPoints(transformedBounds.Width);
        double height = OoxUnits.EmuToPoints(transformedBounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool hasTransform = Math.Abs(transformedBounds.RotationDegrees) > 0.001d || transformedBounds.FlipHorizontal || transformedBounds.FlipVertical;
        RenderPictureOuterShadow(document, graphics, transformedBounds, x, y, width, height, hasTransform, outerShadow, images, ref index);
        if (imageResource.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
        {
            RenderSvgPicture(graphics, document, transformedBounds, imageResource.Bytes, crop, fillRect, cancellationToken);
            StrokePictureFrame(document, graphics, transformedBounds, x, y, width, height, line, hasTransform);
            return;
        }

        PdfImageXObject? image = null;
        if (!crop.IsEmpty)
        {
            image = GetOrCreateCroppedImage(crop);
            if (image is not null)
            {
                crop = default;
            }
        }

        PdfImageXObject? GetOrCreateCroppedImage(CropRect crop)
        {
            string cacheKey = imageResource.PartName + "\u001f" + ImageRecolorCacheKey(recolor) + "\u001fcrop:" +
                crop.Left.ToString("R", CultureInfo.InvariantCulture) + "," +
                crop.Top.ToString("R", CultureInfo.InvariantCulture) + "," +
                crop.Right.ToString("R", CultureInfo.InvariantCulture) + "," +
                crop.Bottom.ToString("R", CultureInfo.InvariantCulture);
            if (imageCache is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
            {
                return cached;
            }

            PdfImageXObject? croppedImage = CreateCroppedImage(
                imageResource.PartName,
                imageResource.ContentType,
                imageResource.Bytes);
            imageCache?.TryAdd(cacheKey, croppedImage);
            return croppedImage;

            PdfImageXObject? CreateCroppedImage(string partName, string contentType, byte[] bytes)
            {
                try
                {
                    if (contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
                    {
                        PngImage png = PngImage.Read(bytes);
                        return CreateCroppedRgbImage(png.Width, png.Height, png.Rgb, png.Alpha, recolor, crop);
                    }

                    if (contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                        contentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
                    {
                        BmpImage bmp = BmpImage.Read(bytes);
                        return CreateCroppedRgbImage(bmp.Width, bmp.Height, bmp.Rgb, bmp.Alpha, recolor, crop);
                    }

                    if (contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                        contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
                    {
                        JpegImage jpeg = JpegImage.Read(bytes);
                        return CreateCroppedRgbImage(jpeg.Width, jpeg.Height, jpeg.Rgb, alpha: null, recolor, crop);
                    }
                }
                catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or IndexOutOfRangeException)
                {
                    diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                        "IMAGE_CROP_UNSUPPORTED_FORMAT",
                        OoxPdfSeverity.Warning,
                        $"Image '{partName}' could not be decoded for Office-style cropped image embedding on slide {slideIndex}; falling back to PDF clipping.",
                        partName,
                        PageIndex: null,
                        SlideIndex: slideIndex,
                        Feature: "image crop",
                        Fallback: "PDF clipping"));
                }

                return null;
            }
        }

        image ??= GetOrCreateImage(imageResource, recolor, imageCache, diagnosticSink, slideIndex);
        if (image is null)
        {
            return;
        }

        string name = "Im" + index++;
        double imageX = x + fillRect.Left * width;
        double imageY = y + fillRect.Bottom * height;
        double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
        double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
        bool transparent = alpha < 0.999d;
        if (hasTransform)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(document, graphics);
            ApplyShapeTransform(graphics, x, y, width, height, transformedBounds);
        }

        if (transparent)
        {
            graphics.SaveState();
            graphics.SetAlpha(alpha, 1d);
        }

        graphics.SaveState();
        double clipX = imageX;
        double clipY = imageY;
        double clipWidth = imageWidth;
        double clipHeight = imageHeight;
        if (!hasTransform &&
            !TryIntersectWithSlideBounds(imageX, imageY, imageWidth, imageHeight, document, out clipX, out clipY, out clipWidth, out clipHeight))
        {
            graphics.RestoreState();
            if (transparent)
            {
                graphics.RestoreState();
            }

            return;
        }

        graphics.ClipRectangleEvenOdd(clipX, clipY, clipWidth, clipHeight);
        if (crop.IsEmpty)
        {
            graphics.DrawImage(name, imageX, imageY, imageWidth, imageHeight);
        }
        else
        {
            graphics.DrawImageCropped(name, imageX, imageY, imageWidth, imageHeight, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }

        graphics.RestoreState();

        if (transparent)
        {
            graphics.RestoreState();
        }

        if (hasTransform)
        {
            graphics.RestoreState();
        }

        images.Add(new PdfImageResource(name, image));

        StrokePictureFrame(document, graphics, transformedBounds, x, y, width, height, line, hasTransform);
    }

    private static void RenderPictureOuterShadow(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        double x,
        double y,
        double width,
        double height,
        bool transformed,
        OuterShadow? outerShadow,
        List<PdfImageResource> images,
        ref int imageIndex)
    {
        if (outerShadow is not { } shadow)
        {
            return;
        }

        graphics.SaveState();
        ClipSlideBoundsEvenOdd(document, graphics);
        if (transformed)
        {
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        if (shadow.BlurRadius > 0d)
        {
            DrawRasterOuterShadow(graphics, x, y, width, height, shadow, images, ref imageIndex);
        }
        else
        {
            DrawOuterShadow(graphics, "rect", x, y, width, height, shadow);
        }

        graphics.RestoreState();
    }

    private static void StrokePictureFrame(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        double x,
        double y,
        double width,
        double height,
        LineStyle line,
        bool transformed)
    {
        if (!line.HasLine)
        {
            return;
        }

        if (transformed)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(document, graphics);
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }
        else
        {
            ClipSlideBoundsEvenOdd(document, graphics);
        }

        bool transparentStroke = line.Alpha < 0.999d;
        if (transparentStroke)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, line.Alpha);
        }

        graphics.SetStrokeRgb(line.Color.Red, line.Color.Green, line.Color.Blue);
        graphics.SetLineWidth(line.Width);
        if (line.DashPattern is { Count: > 0 })
        {
            graphics.SetLineDash(line.DashPattern);
        }

        if (line.Cap is { } cap)
        {
            graphics.SetLineCap(cap);
        }

        if (line.Join is { } join)
        {
            graphics.SetLineJoin(join);
        }

        double outlineOutset = line.Width / 2d;
        graphics.StrokeRectangle(
            x - outlineOutset,
            y - outlineOutset,
            width + line.Width,
            height + line.Width);

        if (line.DashPattern is { Count: > 0 })
        {
            graphics.ClearLineDash();
        }

        if (line.Cap is not null)
        {
            graphics.SetLineCap(0);
        }

        if (line.Join is not null)
        {
            graphics.SetLineJoin(0);
        }

        if (transparentStroke)
        {
            graphics.RestoreState();
        }

        if (transformed)
        {
            graphics.RestoreState();
        }
    }

    private static ShapeBounds ToShapeBounds(PptxSceneBounds bounds)
    {
        return new ShapeBounds(
            bounds.XEmu,
            bounds.YEmu,
            bounds.WidthEmu,
            bounds.HeightEmu,
            bounds.RotationDegrees,
            bounds.FlipHorizontal,
            bounds.FlipVertical);
    }

    private static CropRect ToCropRect(PptxSceneRect rect)
    {
        return new CropRect(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    private static FillRect ToFillRect(PptxSceneRect rect)
    {
        return new FillRect(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    private static PdfImageXObject? GetOrCreateImage(
        PptxSceneImageResource imageResource,
        PptxSceneImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        return GetOrCreateImage(imageResource.PartName, imageResource.ContentType, imageResource.Bytes, recolor, imageCache, diagnosticSink, slideIndex);
    }

    private static PdfImageXObject? GetOrCreateImage(
        string partName,
        string contentType,
        byte[] bytes,
        PptxSceneImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        string cacheKey = partName + "\u001f" + ImageRecolorCacheKey(recolor);
        if (imageCache is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
        {
            return cached;
        }

        PdfImageXObject? image = CreateImage(partName, contentType, bytes, recolor, diagnosticSink, slideIndex);
        imageCache?.TryAdd(cacheKey, image);
        return image;
    }


    private static PdfImageXObject? CreateImage(
        string partName,
        string contentType,
        byte[] bytes,
        PptxSceneImageRecolor recolor,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        try
        {
            if (contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegInfo info = JpegInfo.Read(bytes);
                if (!IsNoImageRecolor(recolor))
                {
                    if (info.IsBaselineDct && info.BitsPerComponent == 8 && info.ComponentCount is 1 or 3)
                    {
                        try
                        {
                            JpegImage jpeg = JpegImage.Read(bytes);
                            byte[] rgb = ApplyImageRecolor(jpeg.Rgb, recolor);
                            return PdfImageXObject.RgbPng(jpeg.Width, jpeg.Height, rgb, alpha: null);
                        }
                        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or IndexOutOfRangeException)
                        {
                        }
                    }

                    diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                        "PPTX_UNSUPPORTED_IMAGE_RECOLOR",
                        OoxPdfSeverity.Warning,
                        $"PPTX {ImageRecolorKindName()} image recolor could not be applied to {contentType} {info.FrameProfileName} image data and was ignored.",
                        partName,
                        PageIndex: null,
                        SlideIndex: slideIndex,
                        Feature: "image recolor",
                        Fallback: "Original image"));
                }

                return PdfImageXObject.Jpeg(info.Width, info.Height, bytes, info.ComponentCount, info.BitsPerComponent);
            }

            if (contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
            {
                PngImage png = PngImage.Read(bytes);
                byte[] rgb = ApplyImageRecolor(png.Rgb, recolor);
                return PdfImageXObject.RgbPng(png.Width, png.Height, rgb, png.Alpha);
            }

            if (contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                contentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
            {
                BmpImage bmp = BmpImage.Read(bytes);
                byte[] rgb = ApplyImageRecolor(bmp.Rgb, recolor);
                return PdfImageXObject.RgbPng(bmp.Width, bmp.Height, rgb, bmp.Alpha);
            }

            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{contentType}' could not be rendered and was ignored: Unsupported image content type.",
                partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: contentType,
                Fallback: "Ignored"));
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{contentType}' could not be rendered and was ignored: {ex.Message}",
                partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: contentType,
                Fallback: "Ignored"));
        }

        return null;

        string ImageRecolorKindName()
        {
            return recolor.Kind switch
            {
                PptxSceneImageRecolorKind.Luminance => "luminance",
                PptxSceneImageRecolorKind.Duotone => "duotone",
                PptxSceneImageRecolorKind.Grayscale => "grayscale",
                PptxSceneImageRecolorKind.BiLevel => "bi-level",
                _ => "none"
            };
        }
    }

    private static PdfImageXObject? CreateCroppedRgbImage(
        int width,
        int height,
        byte[] rgb,
        byte[]? alpha,
        PptxSceneImageRecolor recolor,
        CropRect crop)
    {
        int left = (int)Math.Floor(width * Math.Clamp(crop.Left, 0d, 1d));
        int top = (int)Math.Floor(height * Math.Clamp(crop.Top, 0d, 1d));
        int right = (int)Math.Ceiling(width * (1d - Math.Clamp(crop.Right, 0d, 1d)));
        int bottom = (int)Math.Ceiling(height * (1d - Math.Clamp(crop.Bottom, 0d, 1d)));

        left = Math.Clamp(left, 0, width);
        top = Math.Clamp(top, 0, height);
        right = Math.Clamp(right, left + 1, width);
        bottom = Math.Clamp(bottom, top + 1, height);

        int croppedWidth = right - left;
        int croppedHeight = bottom - top;
        if (croppedWidth <= 0 || croppedHeight <= 0)
        {
            return null;
        }

        byte[] croppedRgb = new byte[croppedWidth * croppedHeight * 3];
        byte[]? croppedAlpha = alpha is null ? null : new byte[croppedWidth * croppedHeight];
        for (int y = 0; y < croppedHeight; y++)
        {
            int sourceY = top + y;
            Buffer.BlockCopy(rgb, (sourceY * width + left) * 3, croppedRgb, y * croppedWidth * 3, croppedWidth * 3);
            if (alpha is not null && croppedAlpha is not null)
            {
                Buffer.BlockCopy(alpha, sourceY * width + left, croppedAlpha, y * croppedWidth, croppedWidth);
            }
        }

        byte[] recoloredRgb = ApplyImageRecolor(croppedRgb, recolor);
        return PdfImageXObject.RgbPng(croppedWidth, croppedHeight, recoloredRgb, croppedAlpha);
    }

    private static byte[] ApplyImageRecolor(byte[] rgb, PptxSceneImageRecolor recolor)
    {
        if (IsNoImageRecolor(recolor))
        {
            return rgb;
        }

        byte[] transformed = new byte[rgb.Length];
        for (int i = 0; i < rgb.Length; i += 3)
        {
            double red = rgb[i];
            double green = rgb[i + 1];
            double blue = rgb[i + 2];
            if (recolor.Kind == PptxSceneImageRecolorKind.Luminance)
            {
                transformed[i] = ApplyBrightnessContrast(red);
                transformed[i + 1] = ApplyBrightnessContrast(green);
                transformed[i + 2] = ApplyBrightnessContrast(blue);
                continue;
            }

            double luma = (0.2126d * red + 0.7152d * green + 0.0722d * blue) / 255d;
            if (recolor.Kind == PptxSceneImageRecolorKind.Grayscale)
            {
                byte gray = ToByte(luma * 255d);
                transformed[i] = gray;
                transformed[i + 1] = gray;
                transformed[i + 2] = gray;
                continue;
            }

            if (recolor.Kind == PptxSceneImageRecolorKind.BiLevel)
            {
                byte value = luma >= recolor.Threshold ? (byte)255 : (byte)0;
                transformed[i] = value;
                transformed[i + 1] = value;
                transformed[i + 2] = value;
                continue;
            }

            transformed[i] = Interpolate(recolor.Dark.Red, recolor.Light.Red, luma);
            transformed[i + 1] = Interpolate(recolor.Dark.Green, recolor.Light.Green, luma);
            transformed[i + 2] = Interpolate(recolor.Dark.Blue, recolor.Light.Blue, luma);
        }

        return transformed;

        byte ApplyBrightnessContrast(double channel)
        {
            double value = channel / 255d;
            value = ApplyContrast(value, recolor.Contrast);
            value = ApplyBrightness(value, recolor.Brightness);
            return ToByte(value * 255d);

            double ApplyContrast(double value, double contrast)
            {
                value = Math.Clamp(value, 0d, 1d);
                if (contrast < 0d)
                {
                    return value * (1d + contrast);
                }

                if (contrast > 0d)
                {
                    double scale = 1d - contrast;
                    return value < 0.5d
                        ? value * scale
                        : 1d - (1d - value) * scale;
                }

                return value;
            }

            double ApplyBrightness(double value, double brightness)
            {
                value = Math.Clamp(value, 0d, 1d);
                return brightness < 0d
                    ? value * (1d + brightness)
                    : value + brightness;
            }
        }
    }

    private static bool IsNoImageRecolor(PptxSceneImageRecolor recolor)
    {
        return recolor.Kind == PptxSceneImageRecolorKind.None;
    }

    private static string ImageRecolorCacheKey(PptxSceneImageRecolor recolor)
    {
        return recolor.Kind switch
        {
            PptxSceneImageRecolorKind.Luminance => FormattableString.Invariant($"lum:{recolor.Brightness:0.#####}:{recolor.Contrast:0.#####}"),
            PptxSceneImageRecolorKind.Duotone => FormattableString.Invariant($"duo:{recolor.Dark.Red:X2}{recolor.Dark.Green:X2}{recolor.Dark.Blue:X2}:{recolor.Light.Red:X2}{recolor.Light.Green:X2}{recolor.Light.Blue:X2}"),
            PptxSceneImageRecolorKind.Grayscale => "gray",
            PptxSceneImageRecolorKind.BiLevel => FormattableString.Invariant($"bi:{recolor.Threshold:0.#####}"),
            _ => "none"
        };
    }

    private static byte Interpolate(byte from, byte to, double ratio)
    {
        return ToByte(from + (to - from) * Math.Clamp(ratio, 0d, 1d));
    }

    private static bool TryReadShapePictureFill(
        XElement shapeProperties,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex,
        ShapePictureFill? pictureFillOverride,
        out string? name,
        out PdfImageXObject? image,
        out CropRect crop,
        out FillRect fillRect,
        out double alpha)
    {
        name = null;
        image = null;
        crop = default;
        fillRect = default;
        alpha = 1d;
        if (images is null || pictureFillOverride is null || !CanRenderPictureFillPreset(ReadPreset(shapeProperties)))
        {
            return false;
        }

        ShapePictureFill resolvedPictureFill = pictureFillOverride.Value;
        if (resolvedPictureFill.Resource is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                resolvedPictureFill.TargetPartName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return false;
        }

        image = GetOrCreateImage(resolvedPictureFill.Resource, PptxSceneImageRecolor.None, imageCache, diagnosticSink, slideIndex);
        if (image is null)
        {
            return false;
        }

        crop = resolvedPictureFill.Crop;
        fillRect = resolvedPictureFill.Fill;
        alpha = resolvedPictureFill.Alpha;
        name = "Im" + imageIndex++;
        return true;
    }

    private static bool CanRenderPictureFillPreset(string preset)
    {
        return preset is "rect" or "ellipse" or "roundRect" ||
            TryCreatePresetPolygonPoints(preset, 0d, 0d, 1d, 1d, out _);
    }

    private static void ClipToPresetShape(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        if (preset == "ellipse")
        {
            graphics.ClipEllipse(x, y, width, height);
        }
        else if (preset == "roundRect")
        {
            graphics.ClipRoundedRectangle(x, y, width, height, ReadRoundRectangleRadius(shapeProperties, presetAdjustmentsOverride, width, height));
        }
        else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
        {
            graphics.ClipPolygon(polygonPoints);
        }
        else
        {
            graphics.ClipRectangle(x, y, width, height);
        }
    }

    private static void DrawImageFill(PdfGraphicsBuilder graphics, string imageName, double x, double y, double width, double height, CropRect crop)
    {
        if (crop.IsEmpty)
        {
            graphics.DrawImage(imageName, x, y, width, height);
        }
        else
        {
            graphics.DrawImageCropped(imageName, x, y, width, height, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }
    }
}
