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
            return PptxSceneChartStyle.Undefined;
        }

        OoxPart? stylePart = package.GetPart(styleRelationship.ResolvedTarget);
        if (stylePart is null)
        {
            return PptxSceneChartStyle.Undefined;
        }

        XDocument document = package.LoadXml(stylePart, cancellationToken);
        return PptxSceneChartStyle.Defined(ReadChartStyleEntries(document, theme, colorMap));
    }

    private static IReadOnlyList<PptxSceneChartStyleEntry> ReadChartStyleEntries(XDocument document, PptxTheme theme, PptxColorMap colorMap)
    {
        if (document.Root is null)
        {
            return [];
        }

        var entries = new List<PptxSceneChartStyleEntry>();
        foreach (XElement roleElement in document.Root.Elements())
        {
            XElement? lineReference = roleElement
                .Descendants()
                .FirstOrDefault(element => element.Name.LocalName == "lnRef");
            int lineReferenceIndex = PptxFormatSchemeResolver.ReadIndex(lineReference);
            XElement? fillReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fillRef");
            XElement? effectReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "effectRef");
            int effectReferenceIndex = PptxFormatSchemeResolver.ReadIndex(effectReference);
            PptxSceneEffectFamily effectReferenceEffects = ReadChartStyleEffectReference(effectReferenceIndex, theme);
            XElement? fontReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fontRef");
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
            PptxSceneChartShapeStyle shapeStyle = ReadChartShapeStyle(
                roleElement
                    .Elements()
                    .FirstOrDefault(element => element.Name.LocalName == "spPr"),
                theme,
                colorMap);
            XElement? defaultRunProperties = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "defRPr");
            PptxSceneChartTextStyleOverride textStyle = ReadChartTextRunStyle(defaultRunProperties, theme, colorMap, fontReference);
            if (lineReference is null &&
                fillReference is null &&
                effectReference is null &&
                string.IsNullOrWhiteSpace((string?)fontReference?.Attribute("idx")) &&
                !HasChartShapeStyle(shapeStyle) &&
                !HasChartTextStyleOverride(textStyle))
            {
                continue;
            }

            entries.Add(new PptxSceneChartStyleEntry(
                roleElement.Name.LocalName,
                effectReferenceEffects,
                line,
                shapeStyle,
                textStyle));
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
}
