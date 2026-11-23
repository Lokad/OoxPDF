using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneEffectFamily ReadEffects(XElement? owner)
    {
        XElement? effectList = owner?.Element(DrawingNamespace + "effectLst");
        bool hasEffectDag = owner?.Element(DrawingNamespace + "effectDag") is not null;
        IReadOnlyList<string> unsupportedEffects = effectList is null
            ? []
            : effectList
                .Elements()
                .Where(effect => effect.Name != DrawingNamespace + "outerShdw" && effect.Name != DrawingNamespace + "glow")
                .Select(effect => effect.Name.LocalName)
                .ToArray();
        return new PptxSceneEffectFamily(effectList is not null, hasEffectDag, unsupportedEffects);
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

        if (PptxColorResolver.TryReadSolidColorWithAlpha(glowElement, theme, colorMap, out RgbColor color, out double alpha))
        {
            double radius = OoxUnits.EmuToPoints(OoxXml.ParseOptionalLong(glowElement, "rad", 0));
            glow = new PptxSceneGlow(
                true,
                color,
                alpha,
                radius);
            return radius > SceneEffectTolerance;
        }

        glow = default;
        return false;
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

        if (PptxColorResolver.TryReadSolidColorWithAlpha(outerShadow, theme, colorMap, out RgbColor color, out double alpha))
        {
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
