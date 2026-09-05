using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneShapeEffectFamily ReadShapeEffects(XElement? shapeProperties)
    {
        XElement? effectList = shapeProperties?.Element(DrawingNamespace + "effectLst");
        XElement? effectDag = shapeProperties?.Element(DrawingNamespace + "effectDag");
        IReadOnlyList<string> unsupportedEffects = effectList is null
            ? []
            : effectList
                .Elements()
                .Where(IsUnsupportedDirectEffect)
                .Select(effect => effect.Name.LocalName)
                .ToArray();
        return new PptxSceneShapeEffectFamily(effectList is not null, effectDag is not null, unsupportedEffects);
    }

    private static bool IsUnsupportedDirectEffect(XElement effect)
    {
        return effect.Name != DrawingNamespace + "outerShdw" &&
            effect.Name != DrawingNamespace + "glow";
    }

    internal static bool TryReadGlow(XElement? shapeProperties, PptxTheme theme, out PptxSceneGlow glow)
    {
        return TryReadGlow(shapeProperties, theme, PptxColorMap.Default, out glow);
    }

    internal static bool TryReadGlow(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, out PptxSceneGlow glow)
    {
        XElement? glowElement = shapeProperties
            ?.Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "glow");
        if (glowElement is null)
        {
            glow = default;
            return false;
        }

        XElement? colorElement = glowElement.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, colorMap, out RgbColor color))
        {
            double radius = OoxUnits.EmuToPoints(OoxXml.ParseOptionalLong(glowElement, "rad", 0));
            glow = new PptxSceneGlow(
                true,
                color,
                ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement))),
                radius);
            return radius > SceneEffectTolerance;
        }

        glow = default;
        return false;
    }

    internal static bool TryReadOuterShadow(XElement? shapeProperties, PptxTheme theme, out PptxSceneOuterShadow shadow)
    {
        return TryReadOuterShadow(shapeProperties, theme, PptxColorMap.Default, out shadow);
    }

    internal static bool TryReadOuterShadow(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, out PptxSceneOuterShadow shadow)
    {
        XElement? outerShadow = shapeProperties
            ?.Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "outerShdw");
        if (outerShadow is null)
        {
            shadow = default;
            return false;
        }

        XElement? colorElement = outerShadow.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, colorMap, out RgbColor color))
        {
            double alpha = ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement)));
            double blurRadius = OoxUnits.EmuToPoints(OoxXml.ParseOptionalLong(outerShadow, "blurRad", 0));
            double distance = OoxUnits.EmuToPoints(OoxXml.ParseOptionalLong(outerShadow, "dist", 0));
            double direction = OoxXml.ParseOptionalLong(outerShadow, "dir", 0) / 60000d * Math.PI / 180d;
            shadow = new PptxSceneOuterShadow(
                true,
                color,
                alpha,
                distance * Math.Cos(direction),
                -distance * Math.Sin(direction),
                blurRadius);
            return true;
        }

        shadow = default;
        return false;
    }
}
