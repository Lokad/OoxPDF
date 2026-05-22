using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Globalization;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<PdfImageResource> RenderPictures(PptxRenderContext context, PdfGraphicsBuilder graphics)
    {
        var images = new List<PdfImageResource>();
        int index = 1;
        foreach (XElement shapeTree in context.SlideXml.Descendants(PresentationNamespace + "spTree"))
        {
            RenderPictureContainer(shapeTree, context, graphics, GroupTransform.Identity, images, ref index);
        }

        return images;
    }

    private static void RenderPictureContainer(
        XElement container,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        GroupTransform transform,
        List<PdfImageResource> images,
        ref int index)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == PresentationNamespace + "pic")
            {
                RenderPicture(
                    child,
                    context.SlideRelationships,
                    context.Package,
                    context.Document,
                    context.Theme,
                    graphics,
                    context.DiagnosticSink,
                    context.SlideNumber,
                    transform,
                    images,
                    context.ImageCache,
                    ref index);
                continue;
            }

            if (child.Name == PresentationNamespace + "grpSp")
            {
                GroupTransform childTransform = transform.Combine(ReadGroupTransform(child));
                RenderPictureContainer(child, context, graphics, childTransform, images, ref index);
                continue;
            }
        }
    }

    private static void RenderPicture(
        XElement picture,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPackage package,
        PptxDocument document,
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform transform,
        List<PdfImageResource> images,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int index)
    {
        string? relationshipId = (string?)picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip")
            ?.Attribute(RelationshipsNamespace + "embed");
        XElement? shapeProperties = picture.Element(PresentationNamespace + "spPr");
        ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
        if (relationshipId is null || bounds is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
        {
            return;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        if (imagePart is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                relationship.ResolvedTarget,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return;
        }

        ImageRecolor recolor = ReadImageRecolor(picture, theme);
        PdfImageXObject? image = GetOrCreateImage(imagePart, recolor, imageCache, diagnosticSink, slideIndex);
        if (image is null)
        {
            return;
        }

        ShapeBounds transformedBounds = transform.Apply(bounds.Value);
        string name = "Im" + index++;
        double x = OoxUnits.EmuToPoints(transformedBounds.X);
        double yTop = OoxUnits.EmuToPoints(transformedBounds.Y);
        double width = OoxUnits.EmuToPoints(transformedBounds.Width);
        double height = OoxUnits.EmuToPoints(transformedBounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        CropRect crop = ReadCrop(picture);
        FillRect fillRect = ReadFillRect(picture);
        double imageX = x + fillRect.Left * width;
        double imageY = y + fillRect.Bottom * height;
        double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
        double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
        double alpha = ReadPictureAlpha(picture);
        bool transparent = alpha < 0.999d;
        bool hasTransform = Math.Abs(transformedBounds.RotationDegrees) > 0.001d || transformedBounds.FlipHorizontal || transformedBounds.FlipVertical;
        if (hasTransform)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, transformedBounds);
        }

        if (transparent)
        {
            graphics.SaveState();
            graphics.SetAlpha(alpha, 1d);
        }

        if (crop.IsEmpty)
        {
            graphics.DrawImage(name, imageX, imageY, imageWidth, imageHeight);
        }
        else
        {
            graphics.DrawImageCropped(name, imageX, imageY, imageWidth, imageHeight, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }

        if (transparent)
        {
            graphics.RestoreState();
        }

        if (hasTransform)
        {
            graphics.RestoreState();
        }

        images.Add(new PdfImageResource(name, image));
    }

    private static PdfImageXObject? GetOrCreateImage(
        OoxPart imagePart,
        ImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        string cacheKey = imagePart.Name + "\u001f" + recolor.CacheKey;
        if (imageCache is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
        {
            return cached;
        }

        PdfImageXObject? image = CreateImage(imagePart, recolor, diagnosticSink, slideIndex);
        imageCache?.TryAdd(cacheKey, image);
        return image;
    }

    private static PdfImageXObject? CreateImage(OoxPart imagePart, ImageRecolor recolor, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex)
    {
        byte[] bytes = imagePart.Bytes;
        try
        {
            if (imagePart.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                imagePart.ContentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegInfo info = JpegInfo.Read(bytes);
                if (!recolor.IsNone)
                {
                    diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                        "PPTX_UNSUPPORTED_IMAGE_RECOLOR",
                        OoxPdfSeverity.Warning,
                        "PPTX image recolor could not be applied to JPEG image data and was ignored.",
                        imagePart.Name,
                        SlideIndex: slideIndex,
                        Feature: "image recolor",
                        Fallback: "Original image"));
                }

                return PdfImageXObject.Jpeg(info.Width, info.Height, bytes);
            }

            if (imagePart.ContentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
            {
                PngImage png = PngImage.Read(bytes);
                byte[] rgb = ApplyImageRecolor(png.Rgb, recolor);
                return PdfImageXObject.RgbPng(png.Width, png.Height, rgb, png.Alpha);
            }

            if (imagePart.ContentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                imagePart.ContentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
            {
                BmpImage bmp = BmpImage.Read(bytes);
                byte[] rgb = ApplyImageRecolor(bmp.Rgb, recolor);
                return PdfImageXObject.RgbPng(bmp.Width, bmp.Height, rgb, bmp.Alpha);
            }

            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{imagePart.ContentType}' could not be rendered and was ignored: Unsupported image content type.",
                imagePart.Name,
                SlideIndex: slideIndex,
                Feature: imagePart.ContentType,
                Fallback: "Ignored"));
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{imagePart.ContentType}' could not be rendered and was ignored: {ex.Message}",
                imagePart.Name,
                SlideIndex: slideIndex,
                Feature: imagePart.ContentType,
                Fallback: "Ignored"));
        }

        return null;
    }

    private static ImageRecolor ReadImageRecolor(XElement picture, PptxTheme theme)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        if (blip is null)
        {
            return ImageRecolor.None;
        }

        XElement? luminance = blip.Element(DrawingNamespace + "lum");
        if (luminance is not null)
        {
            double brightness = ParseOptionalLongAttribute(luminance, "bright", 0) / 100000d;
            double contrast = ParseOptionalLongAttribute(luminance, "contrast", 0) / 100000d;
            return ImageRecolor.Luminance(brightness, contrast);
        }

        XElement? duotone = blip.Element(DrawingNamespace + "duotone");
        if (duotone is not null)
        {
            XElement[] colors = duotone.Elements().Take(2).ToArray();
            if (colors.Length == 2 &&
                TryReadImageRecolorColor(colors[0], theme, out RgbColor dark) &&
                TryReadImageRecolorColor(colors[1], theme, out RgbColor light))
            {
                return ImageRecolor.Duotone(dark, light);
            }
        }

        return ImageRecolor.None;
    }

    private static bool TryReadImageRecolorColor(XElement colorElement, PptxTheme theme, out RgbColor color)
    {
        if (colorElement.Name == DrawingNamespace + "prstClr")
        {
            string? preset = (string?)colorElement.Attribute("val");
            color = preset switch
            {
                "black" => new RgbColor(0, 0, 0),
                "white" => new RgbColor(255, 255, 255),
                _ => default
            };
            return preset is "black" or "white";
        }

        XElement wrapper = new(DrawingNamespace + "solidFill", new XElement(colorElement));
        return TryReadSolidColorWithAlpha(wrapper, theme, out color, out _);
    }

    private static byte[] ApplyImageRecolor(byte[] rgb, ImageRecolor recolor)
    {
        if (recolor.IsNone)
        {
            return rgb;
        }

        byte[] transformed = new byte[rgb.Length];
        for (int i = 0; i < rgb.Length; i += 3)
        {
            double red = rgb[i];
            double green = rgb[i + 1];
            double blue = rgb[i + 2];
            if (recolor.Kind == ImageRecolorKind.Luminance)
            {
                transformed[i] = ApplyBrightnessContrast(red, recolor.Brightness, recolor.Contrast);
                transformed[i + 1] = ApplyBrightnessContrast(green, recolor.Brightness, recolor.Contrast);
                transformed[i + 2] = ApplyBrightnessContrast(blue, recolor.Brightness, recolor.Contrast);
                continue;
            }

            double luma = (0.2126d * red + 0.7152d * green + 0.0722d * blue) / 255d;
            transformed[i] = Interpolate(recolor.Dark.Red, recolor.Light.Red, luma);
            transformed[i + 1] = Interpolate(recolor.Dark.Green, recolor.Light.Green, luma);
            transformed[i + 2] = Interpolate(recolor.Dark.Blue, recolor.Light.Blue, luma);
        }

        return transformed;
    }

    private static byte ApplyBrightnessContrast(double channel, double brightness, double contrast)
    {
        double value = channel / 255d;
        value = (value - 0.5d) * Math.Max(0d, 1d + contrast) + 0.5d;
        value = brightness >= 0d
            ? value + (1d - value) * brightness
            : value * (1d + brightness);
        return ToByte(value * 255d);
    }

    private static byte Interpolate(byte from, byte to, double ratio)
    {
        return ToByte(from + (to - from) * Math.Clamp(ratio, 0d, 1d));
    }

    private static CropRect ReadCrop(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? sourceRectangle = blipFill?.Element(DrawingNamespace + "srcRect");
        if (sourceRectangle is null)
        {
            return default;
        }

        return new CropRect(
            ParsePercentage(sourceRectangle, "l"),
            ParsePercentage(sourceRectangle, "t"),
            ParsePercentage(sourceRectangle, "r"),
            ParsePercentage(sourceRectangle, "b"));
    }

    private static FillRect ReadFillRect(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? fillRectangle = blipFill
            ?.Element(DrawingNamespace + "stretch")
            ?.Element(DrawingNamespace + "fillRect");
        if (fillRectangle is null)
        {
            return default;
        }

        return new FillRect(
            ParsePercentage(fillRectangle, "l"),
            ParsePercentage(fillRectangle, "t"),
            ParsePercentage(fillRectangle, "r"),
            ParsePercentage(fillRectangle, "b"));
    }

    private static double ReadPictureAlpha(XElement picture)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        XElement? alphaModFix = blip?.Element(DrawingNamespace + "alphaModFix");
        if (alphaModFix?.Attribute("amt") is { } amount &&
            int.TryParse(amount.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedAmount))
        {
            return Math.Clamp(parsedAmount / 100000d, 0d, 1d);
        }

        return 1d;
    }

    private static double ParsePercentage(XElement element, string attribute)
    {
        return element.Attribute(attribute) is { } value
            ? Math.Clamp(int.Parse(value.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 0.999d)
            : 0d;
    }

    private enum ImageRecolorKind
    {
        None,
        Luminance,
        Duotone
    }

    private readonly record struct ImageRecolor(ImageRecolorKind Kind, double Brightness, double Contrast, RgbColor Dark, RgbColor Light)
    {
        public static ImageRecolor None { get; } = new(ImageRecolorKind.None, 0d, 0d, default, default);

        public bool IsNone => Kind == ImageRecolorKind.None;

        public string CacheKey => Kind switch
        {
            ImageRecolorKind.Luminance => FormattableString.Invariant($"lum:{Brightness:0.#####}:{Contrast:0.#####}"),
            ImageRecolorKind.Duotone => FormattableString.Invariant($"duo:{Dark.Red:X2}{Dark.Green:X2}{Dark.Blue:X2}:{Light.Red:X2}{Light.Green:X2}{Light.Blue:X2}"),
            _ => "none"
        };

        public static ImageRecolor Luminance(double brightness, double contrast)
        {
            return new ImageRecolor(
                ImageRecolorKind.Luminance,
                Math.Clamp(brightness, -1d, 1d),
                Math.Clamp(contrast, -1d, 1d),
                default,
                default);
        }

        public static ImageRecolor Duotone(RgbColor dark, RgbColor light)
        {
            return new ImageRecolor(ImageRecolorKind.Duotone, 0d, 0d, dark, light);
        }
    }
}
