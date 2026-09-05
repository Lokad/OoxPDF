using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneChartExternalData ReadChartExternalData(OoxPackage package, string chartPartName, XDocument? chartXml, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? externalData = chartXml?.Root?.Element(ChartNamespace + "externalData");
        if (externalData is null)
        {
            return default;
        }

        string? relationshipId = (string?)externalData.Attribute(RelationshipsNamespace + "id");
        string? targetPartName = null;
        if (!string.IsNullOrWhiteSpace(relationshipId))
        {
            targetPartName = package.GetRelationships(chartPartName, cancellationToken)
                .FirstOrDefault(relationship => !relationship.IsExternal &&
                    relationship.Id == relationshipId &&
                    relationship.Type == ChartExternalDataPackageRelationshipType)
                ?.ResolvedTarget;
        }

        (bool? autoUpdate, string autoUpdateValue) = ReadOptionalOoxmlBooleanElementWithValue(externalData, "autoUpdate");
        return new PptxSceneChartExternalData(
            true,
            relationshipId,
            targetPartName,
            ReadPackageResource(package, targetPartName),
            autoUpdate,
            autoUpdateValue);
    }

    private static PptxSceneChartColorStyle ReadChartColorStyle(OoxPackage package, string chartPartName, PptxTheme theme, PptxColorMap colorMap, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? colorRelationship = package.GetRelationships(chartPartName, cancellationToken)
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartColorStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (colorRelationship?.ResolvedTarget is null)
        {
            return new PptxSceneChartColorStyle(false, null, string.Empty, string.Empty, [], 0, [], [], [], null);
        }

        OoxPart? colorPart = package.GetPart(colorRelationship.ResolvedTarget);
        if (colorPart is null)
        {
            return new PptxSceneChartColorStyle(false, colorRelationship.ResolvedTarget, string.Empty, string.Empty, [], 0, [], [], [], null);
        }

        XDocument document = LoadXml(colorPart, cancellationToken);
        IReadOnlyList<PptxSceneChartColorDeclaration> rootDeclarations = ReadChartColorStyleRootDeclarations(document, theme, colorMap);
        IReadOnlyList<PptxSceneChartColorDeclaration> declarations = ReadChartColorStyleDeclarations(document, theme, colorMap, rootDeclarations);
        IReadOnlyList<PptxSceneChartColorVariation> variations = ReadChartColorStyleVariations(document, theme, colorMap);
        var colors = new List<RgbColor>();
        foreach (XElement colorElement in document.Root?.Elements().Where(element => element.Name.Namespace == DrawingNamespace) ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            var wrapper = new XElement(DrawingNamespace + "solidFill", new XElement(colorElement));
            if (PptxColorResolver.TryReadSolidColorWithAlpha(wrapper, theme, colorMap, out RgbColor color, out _))
            {
                colors.Add(color);
            }
        }

        return new PptxSceneChartColorStyle(
            true,
            colorPart.Name,
            (string?)document.Root?.Attribute("meth") ?? string.Empty,
            (string?)document.Root?.Attribute("id") ?? string.Empty,
            colors,
            variations.Count,
            declarations,
            rootDeclarations,
            variations,
            document);
    }

    private static IReadOnlyList<PptxSceneChartColorDeclaration> ReadChartColorStyleRootDeclarations(XDocument document, PptxTheme theme, PptxColorMap colorMap)
    {
        if (document.Root is null)
        {
            return [];
        }

        return document.Root.Elements()
            .Where(PptxColorResolver.IsDrawingColorElement)
            .Select(colorElement => ReadChartColorStyleDeclaration(colorElement, theme, colorMap, variationIndex: null))
            .ToArray();
    }

    private static IReadOnlyList<PptxSceneChartColorDeclaration> ReadChartColorStyleDeclarations(XDocument document, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<PptxSceneChartColorDeclaration> rootDeclarations)
    {
        if (document.Root is null)
        {
            return [];
        }

        var declarations = new List<PptxSceneChartColorDeclaration>(rootDeclarations);

        int variationIndex = 0;
        foreach (XElement variation in document.Root.Elements().Where(IsChartColorStyleVariationElement))
        {
            foreach (XElement colorElement in variation.Descendants().Where(PptxColorResolver.IsDrawingColorElement))
            {
                declarations.Add(ReadChartColorStyleDeclaration(colorElement, theme, colorMap, variationIndex));
            }

            variationIndex++;
        }

        return declarations;
    }

    private static IReadOnlyList<PptxSceneChartColorVariation> ReadChartColorStyleVariations(XDocument document, PptxTheme theme, PptxColorMap colorMap)
    {
        if (document.Root is null)
        {
            return [];
        }

        var variations = new List<PptxSceneChartColorVariation>();
        int variationIndex = 0;
        foreach (XElement variation in document.Root.Elements().Where(IsChartColorStyleVariationElement))
        {
            IReadOnlyList<PptxSceneChartColorDeclaration> declarations = variation
                .Descendants()
                .Where(PptxColorResolver.IsDrawingColorElement)
                .Select(colorElement => ReadChartColorStyleDeclaration(colorElement, theme, colorMap, variationIndex))
                .ToArray();
            variations.Add(new PptxSceneChartColorVariation(
                variationIndex,
                declarations,
                declarations.Where(declaration => declaration.IsResolved).Select(declaration => declaration.Color).OfType<RgbColor>()
                    .ToArray()));
            variationIndex++;
        }

        return variations;
    }

    private static PptxSceneChartColorDeclaration ReadChartColorStyleDeclaration(XElement colorElement, PptxTheme theme, PptxColorMap colorMap, int? variationIndex)
    {
        var wrapper = new XElement(DrawingNamespace + "solidFill", new XElement(colorElement));
        bool isResolved = PptxColorResolver.TryReadSolidColorWithAlpha(wrapper, theme, colorMap, out RgbColor color, out double alpha);
        return new PptxSceneChartColorDeclaration(
            colorElement.Name.LocalName,
            (string?)colorElement.Attribute("val") ?? string.Empty,
            variationIndex,
            isResolved,
            isResolved ? color : null,
            isResolved ? alpha : 1d);
    }

    private static bool IsChartColorStyleVariationElement(XElement element)
    {
        return element.Name.LocalName == "variation";
    }

    private static PptxSceneChartStyle ReadChartStylePart(OoxPackage package, string chartPartName, PptxTheme theme, PptxColorMap colorMap, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? styleRelationship = package.GetRelationships(chartPartName, cancellationToken)
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (styleRelationship?.ResolvedTarget is null)
        {
            return new PptxSceneChartStyle(false, null, string.Empty, null, []);
        }

        OoxPart? stylePart = package.GetPart(styleRelationship.ResolvedTarget);
        if (stylePart is null)
        {
            return new PptxSceneChartStyle(false, styleRelationship.ResolvedTarget, string.Empty, null, []);
        }

        XDocument document = LoadXml(stylePart, cancellationToken);
        return new PptxSceneChartStyle(
            true,
            stylePart.Name,
            (string?)document.Root?.Attribute("id") ?? string.Empty,
            document,
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
            string lineReferenceIndexRaw = (string?)lineReference?.Attribute("idx") ?? string.Empty;
            int lineReferenceIndexValue = lineReference is null ? 0 : OoxXml.ReadOptionalInt(lineReference, "idx", 0);
            int? lineReferenceIndex = lineReferenceIndexValue > 0 ? lineReferenceIndexValue : null;
            XElement? fillReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fillRef");
            string fillReferenceIndexRaw = (string?)fillReference?.Attribute("idx") ?? string.Empty;
            int fillReferenceIndexValue = fillReference is null ? 0 : OoxXml.ReadOptionalInt(fillReference, "idx", 0);
            int? fillReferenceIndex = fillReferenceIndexValue > 0 ? fillReferenceIndexValue : null;
            PptxSceneFillStyle fillReferenceFill = ReadChartStyleFillReference(fillReference, theme, colorMap);
            XElement? effectReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "effectRef");
            string effectReferenceIndexRaw = (string?)effectReference?.Attribute("idx") ?? string.Empty;
            int effectReferenceIndexValue = effectReference is null ? 0 : OoxXml.ReadOptionalInt(effectReference, "idx", 0);
            int? effectReferenceIndex = effectReferenceIndexValue > 0 ? effectReferenceIndexValue : null;
            PptxSceneChartEffectFamily effectReferenceEffects = ReadChartStyleEffectReference(effectReference, theme);
            string fontReferenceIndex = (string?)roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fontRef")
                ?.Attribute("idx") ?? string.Empty;
            PptxSceneLineStyle line = lineReferenceIndex is not null &&
                lineReference is not null &&
                TryReadThemeLineReference(
                    new PptxFormatSchemeReference(
                        lineReference,
                        PptxFormatSchemeResolver.ReadIndex(lineReference),
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
                lineReferenceIndex,
                lineReferenceIndexRaw,
                fillReferenceIndex,
                fillReferenceIndexRaw,
                fillReferenceFill,
                effectReferenceIndex,
                effectReferenceIndexRaw,
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

    private static PptxSceneFillStyle ReadChartStyleFillReference(XElement? fillReference, PptxTheme theme, PptxColorMap colorMap)
    {
        int fillReferenceIndex = PptxFormatSchemeResolver.ReadIndex(fillReference);
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

    private static PptxSceneChartEffectFamily ReadChartStyleEffectReference(XElement? effectReference, PptxTheme theme)
    {
        int effectReferenceIndex = PptxFormatSchemeResolver.ReadIndex(effectReference);
        return effectReferenceIndex > 0 &&
            theme.TryGetEffectStyle(effectReferenceIndex, out XElement? effectStyle)
            ? ReadChartEffects(effectStyle)
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
            style.Effects.HasEffectDag ||
            style.Effects.UnsupportedEffectNames?.Count > 0;
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
        double? characterSpacing = ReadOptionalChartCharacterSpacing(defaultRunProperties);
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
