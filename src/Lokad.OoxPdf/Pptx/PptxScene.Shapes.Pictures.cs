using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneShapePictureFill ReadShapePictureFill(
        XElement? shapeProperties,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? blipFill = shapeProperties?.Element(DrawingNamespace + "blipFill");
        XElement? blip = blipFill?.Element(DrawingNamespace + "blip");
        string? relationshipId = (string?)blip?.Attribute(RelationshipsNamespace + "embed");
        string? targetPartName = ResolveRelationshipTarget(relationshipId, relationships);
        return blipFill is null || shapeProperties is null
            ? default
            : new PptxSceneShapePictureFill(
                true,
                relationshipId ?? string.Empty,
                targetPartName,
                ReadImageResource(package, targetPartName),
                ReadPictureCrop(shapeProperties),
                ReadPictureFill(shapeProperties),
                ReadPictureAlpha(shapeProperties),
                ReadPictureAlphaValue(shapeProperties),
                HasPictureTile(shapeProperties)) { Recolor = ReadImageRecolor(shapeProperties, theme, colorMap) };
    }

    private static PptxSceneImageResource? ReadImageResource(OoxPackage package, string? targetPartName)
    {
        if (targetPartName is null)
        {
            return null;
        }

        OoxPart? imagePart = package.GetPart(targetPartName);
        return imagePart is null
            ? null
            : new PptxSceneImageResource(imagePart.Name, imagePart.ContentType, imagePart.Bytes);
    }

    private static PptxScenePackageResource? ReadPackageResource(OoxPackage package, string? targetPartName)
    {
        if (targetPartName is null)
        {
            return null;
        }

        OoxPart? packagePart = package.GetPart(targetPartName);
        return packagePart is null
            ? null
            : new PptxScenePackageResource(packagePart.Name, packagePart.ContentType, packagePart.Bytes);
    }

    private static string? ResolveRelationshipTarget(
        string? relationshipId,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        return relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
            !relationship.IsExternal
            ? relationship.ResolvedTarget
            : null;
    }

    internal static PptxSceneRect ReadPictureCrop(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? sourceRectangle = blipFill?.Element(DrawingNamespace + "srcRect");
        return sourceRectangle is null
            ? default
            : ReadPercentageRectangle(sourceRectangle);
    }

    internal static PptxSceneRect ReadPictureFill(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? fillRectangle = blipFill
            ?.Element(DrawingNamespace + "stretch")
            ?.Element(DrawingNamespace + "fillRect");
        return fillRectangle is null
            ? default
            : ReadPercentageRectangle(fillRectangle);
    }

    internal static bool HasPictureTile(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        return blipFill?.Element(DrawingNamespace + "tile") is not null;
    }

    internal static double ReadPictureAlpha(XElement picture)
    {
        XElement? blip = ReadPictureBlip(picture);
        XElement? alphaModFix = blip?.Element(DrawingNamespace + "alphaModFix");
        if (alphaModFix?.Attribute("amt") is { } amount &&
            OoxXml.TryParsePercentage(amount.Value, out double parsedAmount))
        {
            return Math.Clamp(parsedAmount, 0d, 1d);
        }

        return 1d;
    }

    internal static string? ReadPictureAlphaValue(XElement picture)
    {
        XElement? blip = ReadPictureBlip(picture);
        return (string?)blip
            ?.Element(DrawingNamespace + "alphaModFix")
            ?.Attribute("amt");
    }

    private static XElement? ReadPictureBlip(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        return blipFill?.Element(DrawingNamespace + "blip");
    }

    internal static PptxSceneImageRecolor ReadImageRecolor(XElement picture, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? blip = ReadPictureBlip(picture);
        if (blip is null)
        {
            return PptxSceneImageRecolor.None;
        }

        XElement? grayscale = blip.Element(DrawingNamespace + "grayscl");
        if (grayscale is not null)
        {
            return PptxSceneImageRecolor.Grayscale(grayscale.Name.LocalName);
        }

        XElement? biLevel = blip.Element(DrawingNamespace + "biLevel");
        if (biLevel is not null)
        {
            string? thresholdValue = (string?)biLevel.Attribute("thresh");
            double threshold = thresholdValue is not null
                ? Math.Clamp(OoxXml.ParsePercentage(thresholdValue), 0d, 1d)
                : 0.5d;
            return PptxSceneImageRecolor.BiLevel(threshold, biLevel.Name.LocalName, thresholdValue);
        }

        XElement? luminance = blip.Element(DrawingNamespace + "lum");
        if (luminance is not null)
        {
            string? brightnessValue = (string?)luminance.Attribute("bright");
            string? contrastValue = (string?)luminance.Attribute("contrast");
            double brightness = brightnessValue is not null
                ? Math.Clamp(OoxXml.ParsePercentage(brightnessValue), -1d, 1d)
                : 0d;
            double contrast = contrastValue is not null
                ? Math.Clamp(OoxXml.ParsePercentage(contrastValue), -1d, 1d)
                : 0d;
            return PptxSceneImageRecolor.Luminance(brightness, contrast, luminance.Name.LocalName, brightnessValue, contrastValue);
        }

        XElement? duotone = blip.Element(DrawingNamespace + "duotone");
        if (duotone is not null)
        {
            XElement[] colors = duotone.Elements().Take(2).ToArray();
            if (colors.Length == 2 &&
                TryReadImageRecolorColor(colors[0], theme, colorMap, out RgbColor dark) &&
                TryReadImageRecolorColor(colors[1], theme, colorMap, out RgbColor light))
            {
                return PptxSceneImageRecolor.Duotone(dark, light, duotone.Name.LocalName);
            }
        }

        return PptxSceneImageRecolor.None;
    }

    private static bool TryReadImageRecolorColor(XElement colorElement, PptxTheme theme, PptxColorMap colorMap, out RgbColor color)
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

        return PptxColorResolver.TryReadColorDeclarationWithAlpha(colorElement, theme, colorMap, out color, out _);
    }

    private static PptxSceneRect ReadPercentageRectangle(XElement element)
    {
        string? left = (string?)element.Attribute("l");
        string? top = (string?)element.Attribute("t");
        string? right = (string?)element.Attribute("r");
        string? bottom = (string?)element.Attribute("b");
        return new PptxSceneRect(
            ParsePercentage(left),
            ParsePercentage(top),
            ParsePercentage(right),
            ParsePercentage(bottom),
            left,
            top,
            right,
            bottom);
    }

    private static double ParsePercentage(string? value)
    {
        return value is not null
            ? Math.Clamp(OoxXml.ParsePercentage(value), 0d, 0.999d)
            : 0d;
    }
}
