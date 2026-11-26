using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneChartExternalData ReadChartExternalData(
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        XDocument? chartXml,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? externalData = chartXml?.Root?.Element(ChartNamespace + "externalData");
        if (externalData is null)
        {
            return default;
        }

        string? relationshipId = (string?)externalData.Attribute(RelationshipsNamespace + "id");
        string? targetPartName = null;
        if (!string.IsNullOrWhiteSpace(relationshipId) &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
            !relationship.IsExternal && relationship.Type == ChartExternalDataPackageRelationshipType)
        {
            targetPartName = relationship.ResolvedTarget;
        }

        return PptxSceneChartExternalData.Defined(
            relationshipId,
            targetPartName,
            targetPartName is null ? null : package.GetPart(targetPartName));
    }

    private static IReadOnlyList<RgbColor> ReadChartPaletteColors(
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        PptxTheme theme,
        PptxColorMap colorMap,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? colorRelationship = relationships.Values
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartColorStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (colorRelationship?.ResolvedTarget is not { } target || package.GetPart(target) is not { } part)
            return [];

        XDocument document = package.LoadXml(part, cancellationToken);
        if (document.Root is not { } root)
            return [];

        // The palette uses resolved root colors and skips chart color-style variation transforms.
        var colors = new List<RgbColor>();
        foreach (XElement element in root.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (PptxColorResolver.IsDrawingColorElement(element) &&
                PptxColorResolver.TryReadColorDeclarationWithAlpha(element, theme, colorMap, out RgbColor color, out _))
                colors.Add(color);
        }

        return colors;
    }

    private static PptxSceneChartStyle ReadChartStylePart(
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        PptxTheme theme,
        PptxColorMap colorMap,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? styleRelationship = relationships.Values
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (styleRelationship?.ResolvedTarget is null)
        {
            return PptxSceneChartStyle.Undefined(null);
        }

        OoxPart? stylePart = package.GetPart(styleRelationship.ResolvedTarget);
        if (stylePart is null)
        {
            return PptxSceneChartStyle.Undefined(styleRelationship.ResolvedTarget);
        }

        XDocument document = package.LoadXml(stylePart, cancellationToken);
        return PptxSceneChartStyle.Defined(
            stylePart.Name,
            (string?)document.Root?.Attribute("id") ?? string.Empty,
            ReadChartStyleEntries(document, theme, colorMap));
    }

    private static IReadOnlyList<PptxSceneChartStyleEntry> ReadChartStyleEntries(XDocument document, PptxTheme theme, PptxColorMap colorMap)
    {
        if (document.Root is null)
        {
            return [];
        }

        var entries = new List<PptxSceneChartStyleEntry>();
        int sourceIndex = 0;
        foreach (XElement roleElement in document.Root.Elements())
        {
            XElement? lineReference = roleElement
                .Descendants()
                .FirstOrDefault(element => element.Name.LocalName == "lnRef");
            int lineReferenceIndex = PptxFormatSchemeResolver.ReadIndex(lineReference);
            XElement? fillReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fillRef");
            int fillReferenceIndex = PptxFormatSchemeResolver.ReadIndex(fillReference);
            PptxSceneFillStyle fillReferenceFill = ReadChartStyleFillReference(fillReference, fillReferenceIndex, theme, colorMap);
            XElement? effectReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "effectRef");
            int effectReferenceIndex = PptxFormatSchemeResolver.ReadIndex(effectReference);
            PptxSceneEffectFamily effectReferenceEffects = ReadChartStyleEffectReference(effectReferenceIndex, theme);
            string fontReferenceIndex = (string?)roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fontRef")
                ?.Attribute("idx") ?? string.Empty;
            PptxSceneLineStyle line = lineReferenceIndex > 0 &&
                lineReference is not null &&
                TryReadThemeLineReference(
                    new PptxFormatSchemeReference(
                        lineReference,
                        lineReferenceIndex,
                        null),
                    theme,
                    colorMap,
                    out PptxSceneLineStyle resolvedLine)
                    ? resolvedLine
                    : default;
            PptxSceneLineStyle shapeLine = ReadChartLine(
                roleElement
                    .Elements()
                    .FirstOrDefault(element => element.Name.LocalName == "spPr"),
                theme,
                colorMap);
            PptxSceneChartShapeStyle shapeStyle = ReadChartShapeStyle(
                roleElement
                    .Elements()
                    .FirstOrDefault(element => element.Name.LocalName == "spPr"),
                theme,
                colorMap);
            PptxSceneChartTextStyleOverride textStyle = ReadChartStyleRoleTextStyle(roleElement, theme, colorMap);
            if (lineReference is null &&
                fillReference is null &&
                effectReference is null &&
                string.IsNullOrWhiteSpace(fontReferenceIndex) &&
                !HasChartShapeStyle(shapeStyle) &&
                !shapeLine.HasLine &&
                !HasChartTextStyleOverride(textStyle))
            {
                sourceIndex++;
                continue;
            }

            entries.Add(new PptxSceneChartStyleEntry(
                roleElement.Name.LocalName,
                sourceIndex,
                roleElement.Name.NamespaceName,
                lineReferenceIndex > 0 ? lineReferenceIndex : null,
                fillReferenceIndex > 0 ? fillReferenceIndex : null,
                fillReferenceFill,
                effectReferenceIndex > 0 ? effectReferenceIndex : null,
                effectReferenceEffects,
                fontReferenceIndex,
                line,
                shapeStyle,
                shapeLine,
                textStyle));
            sourceIndex++;
        }

        return entries;

        bool HasChartTextStyleOverride(PptxSceneChartTextStyleOverride textStyle)
        {
            return textStyle.FontFamily is not null ||
                textStyle.RequestedTypeface is not null ||
                textStyle.TypefaceSource is not null ||
                textStyle.FontSize is not null ||
                textStyle.CharacterSpacing is not null ||
                textStyle.Color is not null ||
                textStyle.Alpha is not null ||
                textStyle.Bold is not null ||
                textStyle.Italic is not null ||
                textStyle.Underline is not null ||
                textStyle.Strike is not null;
        }
    }

    private static PptxSceneFillStyle ReadChartStyleFillReference(XElement? fillReference, int fillReferenceIndex,
        PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? fillStyle = null;
        if (fillReferenceIndex > 0)
        {
            theme.TryGetFillStyle(fillReferenceIndex, out fillStyle);
        }

        if (fillStyle is not null &&
            PptxColorResolver.TryReadSolidColorWithAlpha(fillStyle, theme, colorMap, fillReference, out RgbColor color, out double alpha))
        {
            return new PptxSceneFillStyle(true, color, alpha);
        }

        return fillReferenceIndex > 0 &&
            PptxColorResolver.TryReadSolidColorWithAlpha(fillReference, theme, colorMap, out color, out alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxSceneEffectFamily ReadChartStyleEffectReference(int effectReferenceIndex, PptxTheme theme)
    {
        return effectReferenceIndex > 0 &&
            theme.TryGetEffectStyle(effectReferenceIndex, out XElement? effectStyle)
            ? ReadEffects(effectStyle)
            : default;
    }

    private static bool HasChartShapeStyle(PptxSceneChartShapeStyle style)
    {
        return style.NoFill ||
            style.Fill.HasFill ||
            style.GradientFill?.HasGradient == true ||
            style.PatternFill.HasPattern ||
            style.PictureFill.HasPicture ||
            style.Line.HasLine ||
            style.Glow.HasGlow ||
            style.OuterShadow.HasShadow ||
            style.Effects.HasUnsupportedEffects;
    }

    private static PptxSceneChartTextStyleOverride ReadChartStyleRoleTextStyle(XElement roleElement, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? defaultRunProperties = roleElement
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName == "defRPr");
        XElement? fontReference = roleElement
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName == "fontRef");
        string? typeface = (string?)defaultRunProperties?.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)defaultRunProperties?.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)defaultRunProperties?.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        if (string.IsNullOrWhiteSpace(typeface))
        {
            typeface = (string?)fontReference?.Attribute("idx") switch
            {
                "major" => "+mj-lt",
                "minor" => "+mn-lt",
                _ => null
            };
        }

        PptxThemeTypefaceResolution typefaceResolution = string.IsNullOrWhiteSpace(typeface)
            ? default
            : theme.ResolveTypefaceWithSource(typeface);
        string? fontFamily = typefaceResolution.Typeface;
        double? fontSize = defaultRunProperties?.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        double? characterSpacing = PptxRunTextAttributeReaders.ReadOptionalCharacterSpacing(defaultRunProperties);
        RgbColor? color = PptxColorResolver.TryReadSolidColorWithAlpha(defaultRunProperties?.Element(DrawingNamespace + "solidFill"), theme, colorMap, out RgbColor parsedColor, out double alpha) ||
            PptxColorResolver.TryReadSolidColorWithAlpha(fontReference, theme, colorMap, out parsedColor, out alpha)
                ? parsedColor
                : null;
        bool? bold = defaultRunProperties is null ? null : ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "b");
        bool? italic = defaultRunProperties is null ? null : ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "i");
        bool? underline = defaultRunProperties is null ? null : ReadChartUnderline(defaultRunProperties);
        bool? strike = defaultRunProperties is null ? null : ReadChartStrike(defaultRunProperties);
        return new PptxSceneChartTextStyleOverride(
            fontFamily,
            typefaceResolution.RequestedTypeface,
            typefaceResolution.RequestedTypeface is null ? null : typefaceResolution.Source,
            fontSize,
            characterSpacing,
            color,
            color is null ? null : alpha,
            bold,
            italic,
            underline,
            strike);
    }
}
