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
        var (style, relationshipId, samplingBlip) = ReadDrawingPictureFill(shapeProperties);
        if (!style.HasPicture) return style;
        string? targetPartName = ResolveRelationshipTarget(relationshipId, relationships);
        return style with
        {
            TargetPartName = targetPartName,
            Resource = ReadImageResource(package, targetPartName),
            Recolor = ReadImageRecolor(samplingBlip, theme, colorMap)
        };
    }

    // Charts use the same source projection and resolve the relationship later.
    private static (PptxSceneShapePictureFill Style, string? RelationshipId, XElement? SamplingBlip)
        ReadDrawingPictureFill(XElement? shapeProperties)
    {
        if (shapeProperties?.Element(DrawingNamespace + "blipFill") is not { } blipFill) return default;

        XElement samplingFill = shapeProperties.Element(PresentationNamespace + "blipFill") ?? blipFill;
        XElement? relationshipBlip = blipFill.Element(DrawingNamespace + "blip");
        XElement? samplingBlip = ReferenceEquals(samplingFill, blipFill)
            ? relationshipBlip : samplingFill.Element(DrawingNamespace + "blip");
        string? relationshipId = (string?)relationshipBlip?.Attribute(RelationshipsNamespace + "embed");
        var sampling = ReadPictureSampling(samplingFill, samplingBlip);
        var style = new PptxSceneShapePictureFill(true, relationshipId ?? string.Empty,
            null, null, sampling.Crop, sampling.Fill, sampling.Alpha, sampling.AlphaValue, sampling.HasTile);
        return (style, relationshipId, samplingBlip);
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

    internal static (PptxSceneRect Crop, PptxSceneRect Fill, double Alpha, string? AlphaValue, bool HasTile)
        ReadPictureSampling(XElement? blipFill, XElement? blip)
    {
        XElement? crop = blipFill?.Element(DrawingNamespace + "srcRect");
        XElement? fill = blipFill?.Element(DrawingNamespace + "stretch")?.Element(DrawingNamespace + "fillRect");
        string? alphaValue = (string?)blip?.Element(DrawingNamespace + "alphaModFix")?.Attribute("amt");
        double alpha = OoxXml.TryParsePercentage(alphaValue, out double parsed) ? Math.Clamp(parsed, 0d, 1d) : 1d;
        return (crop is null ? default : ReadPercentageRectangle(crop),
            fill is null ? default : ReadPercentageRectangle(fill),
            alpha, alphaValue, blipFill?.Element(DrawingNamespace + "tile") is not null);
    }

    internal static PptxSceneImageRecolor ReadImageRecolor(XElement? blip, PptxTheme theme, PptxColorMap colorMap)
    {
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
                PptxColorResolver.TryReadColorDeclarationWithAlpha(colors[0], theme, colorMap, out RgbColor dark, out _) &&
                PptxColorResolver.TryReadColorDeclarationWithAlpha(colors[1], theme, colorMap, out RgbColor light, out _))
            {
                return PptxSceneImageRecolor.Duotone(dark, light, duotone.Name.LocalName);
            }
        }

        return PptxSceneImageRecolor.None;
    }

    private static PptxSceneRect ReadPercentageRectangle(XElement element) => new(
        ParsePercentage((string?)element.Attribute("l")),
        ParsePercentage((string?)element.Attribute("t")),
        ParsePercentage((string?)element.Attribute("r")),
        ParsePercentage((string?)element.Attribute("b")));

    private static double ParsePercentage(string? value)
    {
        return value is not null
            ? Math.Clamp(OoxXml.ParsePercentage(value), 0d, 0.999d)
            : 0d;
    }
}
