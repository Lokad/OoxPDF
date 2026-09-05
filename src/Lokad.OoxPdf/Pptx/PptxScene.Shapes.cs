using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    internal static PptxSceneGroupTransform ReadGroupTransform(XElement group)
    {
        XElement? transform = group
            .Element(PresentationNamespace + "grpSpPr")
            ?.Element(DrawingNamespace + "xfrm");
        XElement? offset = transform?.Element(DrawingNamespace + "off");
        XElement? extents = transform?.Element(DrawingNamespace + "ext");
        XElement? childOffset = transform?.Element(DrawingNamespace + "chOff");
        XElement? childExtents = transform?.Element(DrawingNamespace + "chExt");
        if (transform is null || offset is null || extents is null || childOffset is null || childExtents is null)
        {
            return PptxSceneGroupTransform.Identity;
        }

        long width = OoxXml.ParseOptionalLong(extents, "cx", 0L);
        long height = OoxXml.ParseOptionalLong(extents, "cy", 0L);
        long childWidth = Math.Max(1, OoxXml.ParseOptionalLong(childExtents, "cx", 0L));
        long childHeight = Math.Max(1, OoxXml.ParseOptionalLong(childExtents, "cy", 0L));
        return new PptxSceneGroupTransform(
            OoxXml.ParseOptionalLong(offset, "x", 0L),
            OoxXml.ParseOptionalLong(offset, "y", 0L),
            width,
            height,
            OoxXml.ParseOptionalLong(childOffset, "x", 0L),
            OoxXml.ParseOptionalLong(childOffset, "y", 0L),
            width / (double)childWidth,
            height / (double)childHeight,
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            OoxXml.ReadBool(transform, "flipH"),
            OoxXml.ReadBool(transform, "flipV"));
    }

    private static PptxSceneShape ReadShape(
        XElement shape,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        PptxFormatSchemeReference fillReference = PptxFormatSchemeResolver.ResolveFillReference(shape, theme);
        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        PptxSceneFillStyle fill = TryReadShapeFill(shapeProperties, theme, colorMap, fillReference, out RgbColor fillColor, out double fillAlpha) ||
            TryReadInheritedGroupFill(shape, theme, colorMap, out fillColor, out fillAlpha)
                ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
                : default;
        PptxSceneLineStyle line = TryReadShapeLine(shapeProperties, theme, colorMap, lineReference, out RgbColor lineColor, out double lineWidth, out double lineAlpha)
            ? new PptxSceneLineStyle(
                true,
                lineColor,
                lineWidth,
                lineAlpha,
                TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
                ReadPresetDashValue(shapeProperties),
                ReadLineCompound(shapeProperties),
                ReadLineCompoundValue(shapeProperties),
                ReadLineCap(shapeProperties) switch
                {
                    "rnd" => 1,
                    "sq" => 2,
                    _ => null
                },
                ReadLineCap(shapeProperties),
                ReadLineJoin(shapeProperties),
                ReadLineJoinValue(shapeProperties), true)
            : default;
        string ReadShapePreset()
        {
            return (string?)shapeProperties
                ?.Element(DrawingNamespace + "prstGeom")
                ?.Attribute("prst") ?? "rect";
        }

        return new PptxSceneShape(
            ReadShapePreset(),
            ReadPresetAdjustments(shapeProperties),
            shapeProperties?.Element(DrawingNamespace + "custGeom") is not null,
            ReadCustomGeometry(shapeProperties),
            fillReference,
            lineReference,
            shapeProperties?.Element(DrawingNamespace + "noFill") is not null,
            shapeProperties?.Element(DrawingNamespace + "ln")?.Element(DrawingNamespace + "noFill") is not null,
            fill,
            ReadShapeGradientFill(shapeProperties, theme, colorMap),
            ReadShapePatternFill(shapeProperties, theme, colorMap),
            ReadShapePictureFill(shapeProperties, package, relationships),
            HasUnsupportedAlpha(shapeProperties),
            TryReadGlow(shapeProperties, theme, colorMap, out PptxSceneGlow glow) ? glow : default,
            TryReadOuterShadow(shapeProperties, theme, colorMap, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            ReadShapeEffects(shapeProperties),
            line,
            ReadLineEnd(shapeProperties, "headEnd"),
            ReadLineEnd(shapeProperties, "tailEnd"));
    }

    internal static bool HasUnsupportedAlpha(XElement? container)
    {
        return container?.Descendants(DrawingNamespace + "alpha").Any(IsUnsupportedAlpha) == true;
    }

    internal static bool IsUnsupportedAlpha(XElement alpha)
    {
        if (alpha.Attribute("val") is not { } value ||
            !int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ||
            parsed >= 100000)
        {
            return false;
        }

        XElement? color = alpha.Parent;
        XElement? fill = color?.Parent;
        XElement? owner = fill?.Parent;
        XElement? lineOwner = owner?.Parent;
        XElement? gradientStopList = fill?.Parent;
        XElement? gradientFill = gradientStopList?.Parent;
        XElement? gradientOwner = gradientFill?.Parent;
        bool supportedUniformGradientFill = fill?.Name == DrawingNamespace + "gs" &&
            gradientStopList?.Name == DrawingNamespace + "gsLst" &&
            gradientFill?.Name == DrawingNamespace + "gradFill" &&
            gradientOwner?.Name is { } ownerName &&
            (ownerName == PresentationNamespace + "spPr" || ownerName == PresentationNamespace + "bgPr") &&
            IsSupportedAlphaGradientFill();
        bool supportedShapeFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "spPr";
        bool supportedBackgroundFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "bgPr";
        bool supportedShapeLine = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "ln" &&
            lineOwner?.Name == PresentationNamespace + "spPr";
        bool supportedTextFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "rPr";
        bool supportedTableCellFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "tcPr";
        bool supportedTableBorder = fill?.Name == DrawingNamespace + "solidFill" &&
            owner is not null &&
            owner.Name.Namespace == DrawingNamespace &&
            owner.Name.LocalName is "lnL" or "lnR" or "lnT" or "lnB" &&
            lineOwner?.Name == DrawingNamespace + "tcPr";
        bool supportedOuterShadow = fill?.Name == DrawingNamespace + "outerShdw" &&
            owner?.Name == DrawingNamespace + "effectLst";
        bool supportedGlow = fill?.Name == DrawingNamespace + "glow" &&
            owner?.Name == DrawingNamespace + "effectLst";
        return !supportedUniformGradientFill && !supportedShapeFill && !supportedBackgroundFill && !supportedShapeLine && !supportedTextFill && !supportedTableCellFill && !supportedTableBorder && !supportedOuterShadow && !supportedGlow;

        bool IsSupportedAlphaGradientFill()
        {
            if (gradientFill.Element(DrawingNamespace + "gsLst") is not { } alphaStopList ||
                gradientFill.Element(DrawingNamespace + "lin") is not { })
            {
                return false;
            }

            XElement[] stops = alphaStopList
                .Elements(DrawingNamespace + "gs")
                .ToArray();
            return stops.Length >= 2 &&
                stops.All(stop => stop.Elements().FirstOrDefault(PptxColorResolver.IsDrawingColorElement) is not null) &&
                PptxColorResolver.HasSupportedGradientStopAlpha(stops);
        }
    }

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

    internal static bool TryReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, out PptxSceneGradientFill fill)
    {
        return TryReadShapeGradientFill(shapeProperties, theme, PptxColorMap.Default, out fill);
    }

    private static PptxSceneGradientFill ReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        TryReadShapeGradientFill(shapeProperties, theme, colorMap, out PptxSceneGradientFill fill);
        return fill;
    }

    internal static bool TryReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, out PptxSceneGradientFill fill)
    {
        XElement? gradientFill = shapeProperties?.Element(DrawingNamespace + "gradFill");
        XElement? gradientStopList = gradientFill?.Element(DrawingNamespace + "gsLst");
        if (gradientFill is null)
        {
            fill = new PptxSceneGradientFill(false, false, false, 0d, []);
            return false;
        }

        if (gradientStopList is null)
        {
            fill = new PptxSceneGradientFill(false, true, true, 0d, []);
            return false;
        }

        XElement[] rawStops = gradientStopList
            .Elements(DrawingNamespace + "gs")
            .ToArray();
        PptxSceneGradientStop[] stops = rawStops
            .Select(stop => TryReadGradientStop(stop, theme, colorMap, out PptxSceneGradientStop parsed) ? parsed : (PptxSceneGradientStop?)null)
            .OfType<PptxSceneGradientStop>()
            .OrderBy(stop => stop.Offset)
            .ToArray();
        if (stops.Length < 2 ||
            gradientFill.Element(DrawingNamespace + "lin") is not { } linear)
        {
            fill = new PptxSceneGradientFill(false, true, true, 0d, []);
            return false;
        }

        double angleDegrees = OoxXml.ParseOptionalLong(linear, "ang", 0) / 60000d;
        bool hasUnsupportedGradient =
            stops.Length != rawStops.Length ||
            !PptxColorResolver.HasUniformGradientAlpha(stops, static stop => stop.Alpha);
        fill = new PptxSceneGradientFill(true, true, hasUnsupportedGradient, angleDegrees, stops);
        return true;
    }

    private static bool TryReadGradientStop(XElement gradientStop, PptxTheme theme, PptxColorMap colorMap, out PptxSceneGradientStop stop)
    {
        if (!TryReadSolidColorWithAlpha(gradientStop, theme, colorMap, out RgbColor color, out double alpha))
        {
            stop = default;
            return false;
        }

        double ParseGradientPercentage(XElement element, string attribute)
        {
            return element.Attribute(attribute) is { } value
                ? Math.Clamp(int.Parse(value.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 1d)
                : 0d;
        }

        stop = new PptxSceneGradientStop(ParseGradientPercentage(gradientStop, "pos"), color, alpha);
        return true;
    }

    private static IReadOnlyDictionary<string, double> ReadPresetAdjustments(XElement? shapeProperties)
    {
        Dictionary<string, double> adjustments = new(StringComparer.Ordinal);
        foreach (XElement guide in shapeProperties
                     ?.Element(DrawingNamespace + "prstGeom")
                     ?.Element(DrawingNamespace + "avLst")
                     ?.Elements(DrawingNamespace + "gd") ?? [])
        {
            string? name = (string?)guide.Attribute("name");
            string? formula = (string?)guide.Attribute("fmla");
            if (!string.IsNullOrWhiteSpace(name) &&
                formula is not null &&
                formula.StartsWith("val ", StringComparison.Ordinal) &&
                double.TryParse(formula[4..], NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                adjustments[name] = value;
            }
        }

        return adjustments;
    }

    private static PptxSceneCustomGeometry ReadCustomGeometry(XElement? shapeProperties)
    {
        XElement? customGeometry = shapeProperties?.Element(DrawingNamespace + "custGeom");
        if (customGeometry is null)
        {
            return new PptxSceneCustomGeometry(false, false, [], []);
        }

        XElement? pathList = customGeometry.Element(DrawingNamespace + "pathLst");
        XElement[] rawPaths = pathList?.Elements(DrawingNamespace + "path").ToArray() ?? [];
        bool hasUnsupportedGeometry = pathList is null ||
            rawPaths.Length == 0 ||
            rawPaths.Any(path =>
                !path.Elements().Any() ||
                path.Elements().Any(command => ReadCustomCommand(command) is null));

        IReadOnlyList<PptxSceneCustomGuide> guides = customGeometry
            .Element(DrawingNamespace + "gdLst")
            ?.Elements(DrawingNamespace + "gd")
            .Select(guide => new PptxSceneCustomGuide(
                (string?)guide.Attribute("name") ?? string.Empty,
                (string?)guide.Attribute("fmla") ?? string.Empty))
            .Where(guide => !string.IsNullOrWhiteSpace(guide.Name) && !string.IsNullOrWhiteSpace(guide.Formula))
            .ToArray() ?? [];

        IReadOnlyList<PptxSceneCustomPath> paths = customGeometry
            .Element(DrawingNamespace + "pathLst")
            ?.Elements(DrawingNamespace + "path")
            .Select(ReadCustomPath)
            .Where(path => path.Commands.Count > 0)
            .ToArray() ?? [];

        return new PptxSceneCustomGeometry(paths.Count > 0, hasUnsupportedGeometry, guides, paths);
    }

    private static PptxSceneCustomPath ReadCustomPath(XElement path)
    {
        PptxSceneCustomCommand?[] commands = path.Elements().Select(ReadCustomCommand).ToArray();
        if (commands.Any(command => command is null))
        {
            return new PptxSceneCustomPath(0d, 0d, false, false, []);
        }

        return new PptxSceneCustomPath(
            OoxXml.ReadOptionalDouble(path, "w", 21600d),
            OoxXml.ReadOptionalDouble(path, "h", 21600d),
            !string.Equals((string?)path.Attribute("fill"), "none", StringComparison.Ordinal),
            OoxXml.ParseBoolOrDefault(path, "stroke", defaultValue: true),
            commands
                .Cast<PptxSceneCustomCommand>()
                .ToArray());
    }

    private static PptxSceneCustomCommand? ReadCustomCommand(XElement command)
    {
        return command.Name.LocalName switch
        {
            "moveTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.MoveTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "lnTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.LineTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "cubicBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.CubicBezierTo,
                ReadCustomPoints(command).Take(3).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "quadBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.QuadraticBezierTo,
                ReadCustomPoints(command).Take(2).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "arcTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.ArcTo,
                [],
                (string?)command.Attribute("wR") ?? string.Empty,
                (string?)command.Attribute("hR") ?? string.Empty,
                (string?)command.Attribute("stAng") ?? string.Empty,
                (string?)command.Attribute("swAng") ?? string.Empty),
            "close" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.Close,
                [],
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            _ => null
        };
    }

    private static IReadOnlyList<PptxSceneCustomPoint> ReadCustomPoints(XElement command)
    {
        XElement[] points = command.Elements(DrawingNamespace + "pt").ToArray();
        if (points.Length == 0 && command.Name.LocalName is "moveTo" or "lnTo")
        {
            points = [command];
        }

        return points
            .Select(point => new PptxSceneCustomPoint(
                (string?)point.Attribute("x") ?? string.Empty,
                (string?)point.Attribute("y") ?? string.Empty))
            .ToArray();
    }

    private static PptxSceneShapePictureFill ReadShapePictureFill(
        XElement? shapeProperties,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
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
                ReadPictureTile(shapeProperties));
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

    private static PptxScenePatternFill ReadShapePatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        TryReadShapePatternFill(shapeProperties, theme, colorMap, out PptxScenePatternFill fill);
        return fill;
    }

    private static bool TryReadShapePatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, out PptxScenePatternFill fill)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        string? preset = (string?)patternFill?.Attribute("prst");
        if (patternFill is null)
        {
            fill = default;
            return false;
        }

        if (preset is null || !IsSupportedDiagonalPatternFill(preset))
        {
            fill = new PptxScenePatternFill(false, true, true, preset ?? string.Empty, default, default, 1d);
            return false;
        }

        RgbColor foreground = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, colorMap, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, colorMap, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        fill = new PptxScenePatternFill(true, true, false, preset, foreground, background, 1d);
        return true;
    }

    private static bool IsSupportedDiagonalPatternFill(string? preset)
    {
        return preset is not null &&
            (preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase) ||
             preset.Contains("DnDiag", StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryReadShapeFill(
        XElement? shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxFormatSchemeReference fillReference,
        out RgbColor color,
        out double alpha)
    {
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        if (TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out color, out alpha))
        {
            return true;
        }

        if (fillReference.Style is not null &&
            TryReadSolidColorWithAlpha(fillReference.Style, theme, colorMap, fillReference.Reference, out color, out alpha))
        {
            return true;
        }

        return fillReference.Index > 0 && TryReadSolidColorWithAlpha(fillReference.Reference, theme, colorMap, out color, out alpha);
    }

    private static bool TryReadInheritedGroupFill(
        XElement shape,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double alpha)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null ||
            shapeProperties?.Element(DrawingNamespace + "grpFill") is null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            XElement? groupProperties = group.Element(PresentationNamespace + "grpSpPr");
            if (groupProperties?.Element(DrawingNamespace + "noFill") is not null)
            {
                color = default;
                alpha = 1d;
                return false;
            }

            if (TryReadSolidColorWithAlpha(groupProperties, theme, colorMap, out color, out alpha))
            {
                return true;
            }
        }

        color = default;
        alpha = 1d;
        return false;
    }

    private static bool TryReadShapeLine(
        XElement? shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxFormatSchemeReference lineReference,
        out RgbColor color,
        out double lineWidth,
        out double alpha)
    {
        XElement? explicitLine = shapeProperties?.Element(DrawingNamespace + "ln");
        if (explicitLine?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        double? styleLineWidth = TryReadStyleLineWidth(lineReference, out double inheritedLineWidth)
            ? inheritedLineWidth
            : null;
        if (shapeProperties is not null && explicitLine is not null && TryReadLineWithAlpha(shapeProperties, theme, colorMap, out color, out lineWidth, out alpha, styleLineWidth))
        {
            return true;
        }

        if (explicitLine?.Attribute("w") is { } explicitWidthAttribute &&
            lineReference.Style is not null &&
            TryReadSolidColorWithAlpha(lineReference.Style, theme, colorMap, lineReference.Reference, out color, out alpha))
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(explicitWidthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        if (lineReference.Style is null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        if (TryReadThemeLineReference(lineReference, theme, colorMap, out PptxSceneLineStyle line))
        {
            color = line.Color;
            lineWidth = line.Width;
            alpha = line.Alpha;
            return true;
        }

        color = default;
        lineWidth = 0d;
        alpha = 1d;
        return false;
    }

    private static bool TryReadStyleLineWidth(PptxFormatSchemeReference lineReference, out double lineWidth)
    {
        if (lineReference.Style?.Attribute("w") is { } widthAttribute)
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        lineWidth = 0d;
        return false;
    }

    private static bool TryReadThemeLineReference(XElement? lineReference, PptxTheme theme, out PptxSceneLineStyle line)
    {
        return TryReadThemeLineReference(
            new PptxFormatSchemeReference(
                lineReference,
                PptxFormatSchemeResolver.ReadIndex(lineReference),
                null),
            theme,
            PptxColorMap.Default,
            out line);
    }

    private static bool TryReadThemeLineReference(PptxFormatSchemeReference lineReference, PptxTheme theme, out PptxSceneLineStyle line)
    {
        return TryReadThemeLineReference(lineReference, theme, PptxColorMap.Default, out line);
    }

    private static bool TryReadThemeLineReference(PptxFormatSchemeReference lineReference, PptxTheme theme, PptxColorMap colorMap, out PptxSceneLineStyle line)
    {
        XElement? lineStyle = lineReference.Style;
        if (lineStyle is null &&
            lineReference.Index > 0 &&
            theme.TryGetLineStyle(lineReference.Index, out XElement? resolvedLineStyle))
        {
            lineStyle = resolvedLineStyle;
        }

        if (lineStyle is null ||
            lineStyle.Element(DrawingNamespace + "noFill") is not null ||
            !TryReadSolidColorWithAlpha(lineStyle, theme, colorMap, lineReference.Reference, out RgbColor color, out double alpha))
        {
            line = default;
            return false;
        }

        double lineWidth = lineStyle.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        XElement shapeProperties = new(DrawingNamespace + "spPr", new XElement(lineStyle));
        string? lineCap = ReadLineCap(shapeProperties);
        line = new PptxSceneLineStyle(
            true,
            color,
            lineWidth,
            alpha,
            TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
            ReadPresetDashValue(shapeProperties),
            ReadLineCompound(shapeProperties),
            ReadLineCompoundValue(shapeProperties),
            lineCap switch
            {
                "rnd" => 1,
                "sq" => 2,
                _ => null
            },
            lineCap,
            ReadLineJoin(shapeProperties),
            ReadLineJoinValue(shapeProperties), true);
        return true;
    }

    private static bool TryReadLineWithAlpha(
        XElement shapeProperties,
        PptxTheme theme,
        out RgbColor color,
        out double lineWidth,
        out double alpha,
        double? fallbackLineWidth)
    {
        return TryReadLineWithAlpha(shapeProperties, theme, PptxColorMap.Default, out color, out lineWidth, out alpha, fallbackLineWidth);
    }

    private static bool TryReadLineWithAlpha(
        XElement shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double lineWidth,
        out double alpha,
        double? fallbackLineWidth)
    {
        return PptxLineStyleReader.TryReadLineWithAlpha(shapeProperties, theme, colorMap, out color, out lineWidth, out alpha, fallbackLineWidth);
    }

    private static bool TryReadPresetDash(XElement? shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = ReadPresetDashValue(shapeProperties);
        double w = Math.Max(lineWidth, MinimumStrokeWidth);
        dashPattern = presetDash switch
        {
            "dot" or "sysDot" => [w, w * 2d],
            "dash" or "sysDash" => [w * 4d, w * 3d],
            "lgDash" => [w * 8d, w * 3d],
            "dashDot" or "sysDashDot" => [w * 4d, w * 3d, w, w * 3d],
            "lgDashDot" => [w * 8d, w * 3d, w, w * 3d],
            "lgDashDotDot" or "sysDashDotDot" => [w * 8d, w * 3d, w, w * 3d, w, w * 3d],
            _ => []
        };
        return dashPattern.Count > 0;
    }

    private static string? ReadPresetDashValue(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
    }

    private static PptxSceneLineCompound? ReadLineCompound(XElement? shapeProperties)
    {
        string? compound = ReadLineCompoundValue(shapeProperties);
        return compound switch
        {
            "sng" => PptxSceneLineCompound.Single,
            "dbl" => PptxSceneLineCompound.Double,
            "thickThin" => PptxSceneLineCompound.ThickThin,
            "thinThick" => PptxSceneLineCompound.ThinThick,
            "tri" => PptxSceneLineCompound.Triple,
            _ => null
        };
    }

    private static string? ReadLineCompoundValue(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Attribute("cmpd");
    }

    private static string? ReadLineCap(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Attribute("cap");
    }

    private static int? ReadLineJoin(XElement? shapeProperties)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return 1;
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return 2;
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return 0;
        }

        return null;
    }

    private static string? ReadLineJoinValue(XElement? shapeProperties)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return "round";
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return "bevel";
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return "miter";
        }

        return null;
    }

    private static PptxSceneLineEnd ReadLineEnd(XElement? shapeProperties, string elementName)
    {
        XElement? end = shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName);
        string? type = (string?)end?.Attribute("type");
        string? width = (string?)end?.Attribute("w");
        string? length = (string?)end?.Attribute("len");
        return new PptxSceneLineEnd(
            ReadLineEndKind(type),
            type,
            ReadLineEndScale(width),
            width,
            ReadLineEndScale(length),
            length);
    }

    private static PptxSceneLineEndKind ReadLineEndKind(string? type)
    {
        return type switch
        {
            "triangle" => PptxSceneLineEndKind.Triangle,
            "arrow" => PptxSceneLineEndKind.Arrow,
            "stealth" => PptxSceneLineEndKind.Stealth,
            "diamond" => PptxSceneLineEndKind.Diamond,
            "oval" => PptxSceneLineEndKind.Oval,
            _ => PptxSceneLineEndKind.None
        };
    }

    private static double ReadLineEndScale(string? value)
    {
        return value switch
        {
            "sm" => 0.5d,
            "lg" => 1.5d,
            _ => 1d
        };
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

    internal static PptxScenePictureTile ReadPictureTile(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? tile = blipFill?.Element(DrawingNamespace + "tile");
        return tile is null
            ? default
            : new PptxScenePictureTile(
                true,
                tile.Name.LocalName,
                (string?)tile.Attribute("algn"),
                (string?)tile.Attribute("flip"),
                (string?)tile.Attribute("sx"),
                (string?)tile.Attribute("sy"),
                (string?)tile.Attribute("tx"),
                (string?)tile.Attribute("ty"));
    }

    internal static double ReadPictureAlpha(XElement picture)
    {
        XElement? blip = ReadPictureBlip(picture);
        XElement? alphaModFix = blip?.Element(DrawingNamespace + "alphaModFix");
        if (alphaModFix?.Attribute("amt") is { } amount &&
            int.TryParse(amount.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedAmount))
        {
            return Math.Clamp(parsedAmount / 100000d, 0d, 1d);
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

    internal static PptxSceneImageRecolor ReadImageRecolor(XElement picture, PptxTheme theme)
    {
        return ReadImageRecolor(picture, theme, PptxColorMap.Default);
    }

    internal static PptxSceneImageRecolor ReadImageRecolor(XElement picture, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
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
                ? Math.Clamp(int.Parse(thresholdValue, CultureInfo.InvariantCulture) / 100000d, 0d, 1d)
                : 0.5d;
            return PptxSceneImageRecolor.BiLevel(threshold, biLevel.Name.LocalName, thresholdValue);
        }

        XElement? luminance = blip.Element(DrawingNamespace + "lum");
        if (luminance is not null)
        {
            string? brightnessValue = (string?)luminance.Attribute("bright");
            string? contrastValue = (string?)luminance.Attribute("contrast");
            double brightness = brightnessValue is not null
                ? Math.Clamp(int.Parse(brightnessValue, CultureInfo.InvariantCulture) / 100000d, -1d, 1d)
                : 0d;
            double contrast = contrastValue is not null
                ? Math.Clamp(int.Parse(contrastValue, CultureInfo.InvariantCulture) / 100000d, -1d, 1d)
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

    private static bool TryReadImageRecolorColor(XElement colorElement, PptxTheme theme, out RgbColor color)
    {
        return TryReadImageRecolorColor(colorElement, theme, PptxColorMap.Default, out color);
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

        XElement wrapper = new(DrawingNamespace + "solidFill", new XElement(colorElement));
        return TryReadSolidColorWithAlpha(wrapper, theme, colorMap, out color, out _);
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
            ? Math.Clamp(int.Parse(value, CultureInfo.InvariantCulture) / 100000d, 0d, 0.999d)
            : 0d;
    }
}
