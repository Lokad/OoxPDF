using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneChart ReadChart(
        XElement frame,
        OoxPackage package,
        PptxTheme theme,
        PptxColorMap colorMap,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? graphicData = frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData");
        string? relationshipId = (string?)graphicData
            ?.Element(ChartNamespace + "chart")
            ?.Attribute(RelationshipsNamespace + "id");
        string? targetPartName = relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship)
            ? relationship.ResolvedTarget
            : null;
        OoxPart? chartPart = targetPartName is null ? null : package.GetPart(targetPartName);
        XDocument? chartXml = chartPart is null ? null : LoadXml(chartPart, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PptxSceneChartExternalData externalData = chartPart is null
            ? default
            : ReadChartExternalData(package, chartPart.Name, chartXml, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PptxSceneChartColorStyle colorStyle = chartPart is null
            ? new PptxSceneChartColorStyle(false, null, string.Empty, string.Empty, [], 0, [], [], [], null)
            : ReadChartColorStyle(package, chartPart.Name, theme, colorMap, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PptxSceneChartStyle stylePart = chartPart is null
            ? new PptxSceneChartStyle(false, null, string.Empty, null, [])
            : ReadChartStylePart(package, chartPart.Name, theme, colorMap, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<RgbColor>? paletteColors = colorStyle.Colors.Count == 0 ? null : colorStyle.Colors;
        IReadOnlyList<PptxSceneChartPlot> plots = ReadChartPlots(chartXml, theme, colorMap);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyDictionary<string, OoxRelationship> chartRelationships = chartPart is null
            ? new Dictionary<string, OoxRelationship>(StringComparer.Ordinal)
            : package.GetRelationships(chartPart.Name, cancellationToken).ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<PptxSceneChartAxis> axes = ReadChartAxes(chartXml, theme, colorMap, stylePart);
        return new PptxSceneChart(
            relationshipId,
            targetPartName,
            chartXml,
            colorMap,
            externalData,
            ReadChartOptions(chartXml),
            paletteColors,
            colorStyle,
            stylePart,
            ReadChartElementValue(chartXml?.Root, "style"),
            plots,
            axes,
            ReadChartTitle(chartXml, theme, colorMap, plots),
            ReadChartLegend(chartXml, theme, colorMap),
            ReadChartTextStyleOverride(chartXml?.Root, theme, colorMap),
            ReadChartPlotAreaManualLayout(chartXml),
            ReadChartShapeStyle(chartXml?.Root?.Element(ChartNamespace + "spPr"), theme, colorMap),
            ReadChartShapeStyle(chartXml?
                .Descendants(ChartNamespace + "plotArea")
                .FirstOrDefault()
                ?.Element(ChartNamespace + "spPr"), theme, colorMap),
            chartRelationships);
    }

    internal static PptxSceneChartOptions ReadChartOptions(XDocument? chartXml)
    {
        XElement? chartSpace = chartXml?.Root;
        XElement? chart = chartSpace?.Element(ChartNamespace + "chart");
        string displayBlanksAs = chart is null ? string.Empty : ReadChartElementValue(chart, "dispBlanksAs");
        (bool? date1904, string date1904Value) = ReadOptionalOoxmlBooleanElementWithValue(chartSpace, "date1904");
        (bool? roundedCorners, string roundedCornersValue) = ReadOptionalOoxmlBooleanElementWithValue(chartSpace, "roundedCorners");
        (bool? plotVisibleOnly, string plotVisibleOnlyValue) = ReadOptionalOoxmlBooleanElementWithValue(chart, "plotVisOnly");
        (bool? showDataLabelsOverMaximum, string showDataLabelsOverMaximumValue) = ReadOptionalOoxmlBooleanElementWithValue(chart, "showDLblsOverMax");
        return new PptxSceneChartOptions(
            date1904,
            date1904Value,
            roundedCorners,
            roundedCornersValue,
            plotVisibleOnly,
            plotVisibleOnlyValue,
            showDataLabelsOverMaximum,
            showDataLabelsOverMaximumValue,
            ParseChartDisplayBlanksAs(displayBlanksAs),
            displayBlanksAs);
    }

    private static IReadOnlyList<PptxSceneChartPlot> ReadChartPlots(XDocument? chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        if (plotArea is null)
        {
            return [];
        }

        var plots = new List<PptxSceneChartPlot>();
        var kindIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (XElement plot in plotArea.Elements().Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)))
        {
            string kind = plot.Name.LocalName;
            int kindIndex = kindIndexes.TryGetValue(kind, out int nextKindIndex) ? nextKindIndex : 0;
            kindIndexes[kind] = kindIndex + 1;
            string grouping = ReadChartElementValue(plot, "grouping");
            string barDirection = ReadChartElementValue(plot, "barDir");
            string scatterStyle = ReadChartElementValue(plot, "scatterStyle");
            string radarStyle = ReadChartElementValue(plot, "radarStyle");
            (bool? markersEnabled, string markersEnabledValue) = ReadChartPlotMarkersEnabled(plot);
            (bool? varyColors, string varyColorsValue) = ReadChartPlotVaryColors(plot);
            PptxSceneChartPlotKind plotKind = ParseChartPlotKind(kind);
            (double? gapWidth, string gapWidthValue) = ReadChartElementDoubleWithValue(plot, "gapWidth");
            (double? overlap, string overlapValue) = ReadChartElementDoubleWithValue(plot, "overlap");
            (double? holeSize, string holeSizeValue) = ReadChartElementDoubleWithValue(plot, "holeSize");
            (double? firstSliceAngle, string firstSliceAngleValue) = ReadChartElementDoubleWithValue(plot, "firstSliceAng");
            string[] axisIds = plot
                .Elements(ChartNamespace + "axId")
                .Select(ReadChartValueAttribute)
                .Where(value => value.Length != 0)
                .ToArray();
            plots.Add(new PptxSceneChartPlot(
                plotKind,
                kind,
                plots.Count,
                kindIndex,
                plot.Elements(ChartNamespace + "ser").Count(),
                axisIds,
                ReadChartSeries(plot, theme, colorMap, plotKind, markersEnabled == true),
                ParseChartGrouping(grouping),
                grouping,
                ParseChartBarDirection(barDirection),
                barDirection,
                ParseChartScatterStyle(scatterStyle),
                scatterStyle,
                ParseChartRadarStyle(radarStyle),
                radarStyle,
                markersEnabled,
                markersEnabledValue,
                varyColors,
                varyColorsValue,
                gapWidth,
                gapWidthValue,
                overlap,
                overlapValue,
                holeSize,
                holeSizeValue,
                firstSliceAngle,
                firstSliceAngleValue,
                ReadChartDataLabels(plot, theme, colorMap),
                plot));
        }

        return plots;
    }

    internal static PptxSceneChartDataLabels ReadChartDataLabels(XElement plot, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? labels = plot.Element(ChartNamespace + "dLbls") ??
            plot.Elements(ChartNamespace + "ser")
                .Select(series => series.Element(ChartNamespace + "dLbls"))
                .FirstOrDefault(element => element is not null);
        if (labels is null)
        {
            return new PptxSceneChartDataLabels(
                ShowValue: null,
                ShowValueValue: string.Empty,
                ShowPercent: null,
                ShowPercentValue: string.Empty,
                ShowCategoryName: null,
                ShowCategoryNameValue: string.Empty,
                ShowSeriesName: null,
                ShowSeriesNameValue: string.Empty,
                ShowLeaderLines: null,
                ShowLeaderLinesValue: string.Empty,
                ShowLegendKey: null,
                ShowLegendKeyValue: string.Empty,
                ShowBubbleSize: null,
                ShowBubbleSizeValue: string.Empty,
                LeaderLines: default,
                PositionKind: PptxSceneChartDataLabelPosition.Unknown,
                Position: string.Empty,
                Separator: string.Empty,
                NumberFormat: string.Empty,
                NumberFormatInfo: default,
                Layout: default,
                TextStyle: new PptxSceneChartTextStyleOverride(null, null, null, null, null, null, null, null),
                TextBodyProperties: default,
                ShapeStyle: new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
                RejectedOverrideIndexValues: [],
                Overrides: [],
                IsDefined: false);
        }

        PptxSceneChartNumberFormat numberFormat = ReadChartNumberFormat(labels);
        (bool? showValue, string showValueValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showVal");
        (bool? showPercent, string showPercentValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showPercent");
        (bool? showCategoryName, string showCategoryNameValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showCatName");
        (bool? showSeriesName, string showSeriesNameValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showSerName");
        (bool? showLeaderLines, string showLeaderLinesValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showLeaderLines");
        (bool? showLegendKey, string showLegendKeyValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showLegendKey");
        (bool? showBubbleSize, string showBubbleSizeValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showBubbleSize");
        return new PptxSceneChartDataLabels(
                showValue,
                showValueValue,
                showPercent,
                showPercentValue,
                showCategoryName,
                showCategoryNameValue,
                showSeriesName,
                showSeriesNameValue,
                showLeaderLines,
                showLeaderLinesValue,
                showLegendKey,
                showLegendKeyValue,
                showBubbleSize,
                showBubbleSizeValue,
                ReadChartLeaderLines(labels, theme, colorMap),
                ParseChartDataLabelPosition(ReadChartElementValue(labels, "dLblPos")),
                ReadChartElementValue(labels, "dLblPos"),
                labels.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                numberFormat.FormatCode,
                numberFormat,
                ReadChartManualLayout(labels),
                ReadChartTextStyleOverride(labels, theme, colorMap),
                ReadChartTextBodyProperties(labels),
                ReadChartShapeStyle(labels.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadRejectedChartNonNegativeIndexValues(labels, "dLbl"),
                ReadChartDataLabelOverrides(labels, theme, colorMap),
                IsDefined: true);
    }

    private static PptxSceneChartLeaderLines ReadChartLeaderLines(XElement labels, PptxTheme theme)
    {
        return ReadChartLeaderLines(labels, theme, PptxColorMap.Default);
    }

    private static PptxSceneChartLeaderLines ReadChartLeaderLines(XElement labels, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? leaderLines = labels.Element(ChartNamespace + "leaderLines");
        return leaderLines is null
            ? default
            : new PptxSceneChartLeaderLines(true, ReadChartLine(leaderLines.Element(ChartNamespace + "spPr"), theme, colorMap));
    }

    private static IReadOnlyList<PptxSceneChartDataLabelOverride> ReadChartDataLabelOverrides(XElement labels, PptxTheme theme)
    {
        return ReadChartDataLabelOverrides(labels, theme, PptxColorMap.Default);
    }

    private static IReadOnlyList<PptxSceneChartDataLabelOverride> ReadChartDataLabelOverrides(XElement labels, PptxTheme theme, PptxColorMap colorMap)
    {
        var overrides = new List<PptxSceneChartDataLabelOverride>();
        foreach (XElement label in labels.Elements(ChartNamespace + "dLbl"))
        {
            if (!TryReadChartNonNegativeIndex(label, out int index, out string indexValue))
            {
                continue;
            }

            PptxSceneChartNumberFormat numberFormat = ReadChartNumberFormat(label);
            (bool? isDeleted, string isDeletedValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "delete");
            (bool? showValue, string showValueValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showVal");
            (bool? showPercent, string showPercentValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showPercent");
            (bool? showCategoryName, string showCategoryNameValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showCatName");
            (bool? showSeriesName, string showSeriesNameValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showSerName");
            (bool? showLeaderLines, string showLeaderLinesValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showLeaderLines");
            (bool? showLegendKey, string showLegendKeyValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showLegendKey");
            (bool? showBubbleSize, string showBubbleSizeValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showBubbleSize");
            overrides.Add(new PptxSceneChartDataLabelOverride(
                index,
                indexValue,
                isDeleted,
                isDeletedValue,
                showValue,
                showValueValue,
                showPercent,
                showPercentValue,
                showCategoryName,
                showCategoryNameValue,
                showSeriesName,
                showSeriesNameValue,
                showLeaderLines,
                showLeaderLinesValue,
                showLegendKey,
                showLegendKeyValue,
                showBubbleSize,
                showBubbleSizeValue,
                ReadChartLeaderLines(label, theme, colorMap),
                ReadChartText(label.Element(ChartNamespace + "tx"), false) ?? string.Empty,
                ReadChartTextRuns(label.Element(ChartNamespace + "tx"), theme, colorMap),
                ParseChartDataLabelPosition(ReadChartElementValue(label, "dLblPos")),
                ReadChartElementValue(label, "dLblPos"),
                label.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                numberFormat.FormatCode,
                numberFormat,
                ReadChartManualLayout(label),
                ReadChartTextStyleOverride(label, theme, colorMap),
                ReadChartTextBodyProperties(label),
                ReadChartShapeStyle(label.Element(ChartNamespace + "spPr"), theme, colorMap)));
        }

        return overrides;
    }

    internal static PptxSceneChartNumberFormat ReadChartNumberFormat(XElement parent)
    {
        XElement? numberFormat = parent.Element(ChartNamespace + "numFmt");
        return numberFormat is null
            ? default
            : new PptxSceneChartNumberFormat(
                IsDefined: true,
                FormatCode: (string?)numberFormat.Attribute("formatCode") ?? string.Empty,
                SourceLinked: ReadOptionalOoxmlBooleanAttribute(numberFormat, "sourceLinked"),
                SourceLinkedValue: (string?)numberFormat.Attribute("sourceLinked") ?? string.Empty);
    }

    internal static PptxSceneChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme)
    {
        return ReadChartShapeStyle(shapeProperties, theme, PptxColorMap.Default);
    }

    internal static PptxSceneChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        bool noFill = shapeProperties?.Element(DrawingNamespace + "noFill") is not null;
        PptxSceneFillStyle fill = !noFill && PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartShapeStyle(
            noFill,
            fill,
            noFill ? new PptxSceneGradientFill(false, false, false, 0d, []) : ReadShapeGradientFill(shapeProperties, theme, colorMap),
            noFill ? default : ReadChartPatternFill(shapeProperties, theme, colorMap),
            noFill ? default : ReadChartPictureFill(shapeProperties),
            ReadChartLine(shapeProperties, theme, colorMap),
            TryReadGlow(shapeProperties, theme, colorMap, out PptxSceneGlow glow) ? glow : default,
            TryReadOuterShadow(shapeProperties, theme, colorMap, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            ReadChartEffects(shapeProperties));
    }

    private static PptxSceneShapePictureFill ReadChartPictureFill(XElement? shapeProperties)
    {
        XElement? blipFill = shapeProperties?.Element(DrawingNamespace + "blipFill");
        if (blipFill is null || shapeProperties is null)
        {
            return default;
        }

        string relationshipId = (string?)blipFill
            .Element(DrawingNamespace + "blip")
            ?.Attribute(RelationshipsNamespace + "embed")
            ?? string.Empty;
        return new PptxSceneShapePictureFill(
            true,
            relationshipId,
            null,
            null,
            ReadPictureCrop(shapeProperties),
            ReadPictureFill(shapeProperties),
            ReadPictureAlpha(shapeProperties),
            ReadPictureAlphaValue(shapeProperties),
            ReadPictureTile(shapeProperties));
    }

    private static PptxSceneChartEffectFamily ReadChartEffects(XElement? shapeProperties)
    {
        XElement? effectList = shapeProperties?.Element(DrawingNamespace + "effectLst");
        XElement? effectDag = shapeProperties?.Element(DrawingNamespace + "effectDag");
        IReadOnlyList<string> unsupportedEffects = effectList is null
            ? []
            : effectList
                .Elements()
                .Where(IsUnsupportedChartEffect)
                .Select(effect => effect.Name.LocalName)
                .ToArray();
        return new PptxSceneChartEffectFamily(effectList is not null, effectDag is not null, unsupportedEffects);
    }

    private static bool IsUnsupportedChartEffect(XElement effect)
    {
        return IsUnsupportedDirectEffect(effect);
    }

    internal static string ReadChartElementValue(XElement? element, string childName)
    {
        return ReadChartValueAttribute(element?.Element(ChartNamespace + childName));
    }

    internal static string ReadChartValueAttribute(XElement? element)
    {
        return ReadOptionalChartValueAttribute(element) ?? string.Empty;
    }

    internal static string? ReadOptionalChartValueAttribute(XElement? element)
    {
        return (string?)element?.Attribute("val");
    }

    internal static (int? Value, string RawValue) ReadChartPointIndexAttribute(XElement? point, bool requireNonNegative)
    {
        string value = (string?)point?.Attribute("idx") ?? string.Empty;
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) &&
            (!requireNonNegative || parsed >= 0)
            ? (parsed, value)
            : (null, value);
    }

    private static double? ReadChartElementDouble(XElement element, string childName)
    {
        (double? parsed, _) = ReadChartElementDoubleWithValue(element, childName);
        return parsed;
    }

    internal static (double? Value, string RawValue) ReadChartElementDoubleWithValue(XElement element, string childName)
    {
        string value = ReadChartElementValue(element, childName);
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static int? ReadChartElementInt(XElement element, string childName)
    {
        (int? parsed, _) = ReadChartElementIntWithValue(element, childName);
        return parsed;
    }

    internal static (int? Value, string RawValue) ReadChartElementIntWithValue(XElement? element, string childName)
    {
        string value = ReadChartElementValue(element, childName);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? (parsed, value)
            : (null, value);
    }

    internal static (int? Value, string RawValue) ReadChartCachePointCount(XElement? cache)
    {
        string value = ReadChartValueAttribute(cache?.Element(ChartNamespace + "ptCount"));
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? (parsed, value)
            : (null, value);
    }

    internal static (bool? Value, string RawValue) ReadChartPlotVaryColors(XElement plot)
    {
        XElement? varyColors = plot.Element(ChartNamespace + "varyColors");
        return varyColors is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(varyColors, defaultValue: true), (string?)varyColors.Attribute("val") ?? string.Empty);
    }

    private static (bool? Value, string RawValue) ReadChartPlotMarkersEnabled(XElement plot)
    {
        XElement? marker = plot.Element(ChartNamespace + "marker");
        return marker is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(marker), (string?)marker.Attribute("val") ?? string.Empty);
    }

    private static IReadOnlyList<PptxSceneChartAxis> ReadChartAxes(XDocument? chartXml, PptxTheme theme, PptxColorMap colorMap, PptxSceneChartStyle stylePart)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        if (plotArea is null)
        {
            return [];
        }

        var axes = new List<PptxSceneChartAxis>();
        foreach (XElement axis in plotArea.Elements().Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Ax", StringComparison.Ordinal)))
        {
            string id = ReadChartValueAttribute(axis.Element(ChartNamespace + "axId"));
            if (id.Length == 0)
            {
                continue;
            }

            string axisPosition = ReadChartElementValue(axis, "axPos");
            string crosses = ReadChartElementValue(axis, "crosses");
            string crossBetween = ReadChartElementValue(axis, "crossBetween");
            string orientation = ReadChartElementValue(axis.Element(ChartNamespace + "scaling"), "orientation");
            PptxSceneChartAxisOrientation orientationKind = ParseChartAxisOrientation(orientation);
            string tickLabelPosition = ReadChartElementValue(axis, "tickLblPos");
            string majorTickMark = ReadChartElementValue(axis, "majorTickMark");
            string minorTickMark = ReadChartElementValue(axis, "minorTickMark");
            (double? crossesAt, string crossesAtValue) = ReadChartElementDoubleWithValue(axis, "crossesAt");
            (double? minimum, string minimumValue) = ReadChartAxisScalingValueWithValue(axis, "min");
            (double? maximum, string maximumValue) = ReadChartAxisScalingValueWithValue(axis, "max");
            (double? majorUnit, string majorUnitValue) = ReadChartAxisUnitValueWithValue(axis, "majorUnit");
            (double? minorUnit, string minorUnitValue) = ReadChartAxisUnitValueWithValue(axis, "minorUnit");
            (int? labelOffset, string labelOffsetValue) = ReadChartElementIntWithValue(axis, "lblOffset");
            (int? tickLabelSkip, string tickLabelSkipValue) = ReadChartElementIntWithValue(axis, "tickLblSkip");
            (int? tickMarkSkip, string tickMarkSkipValue) = ReadChartElementIntWithValue(axis, "tickMarkSkip");
            (bool? isDeleted, string isDeletedValue) = ReadOptionalOoxmlBooleanElementWithValue(axis, "delete");
            (bool? noMultiLevelLabels, string noMultiLevelLabelsValue) = ReadOptionalOoxmlBooleanElementWithValue(axis, "noMultiLvlLbl");
            XElement? majorGridlines = axis.Element(ChartNamespace + "majorGridlines");
            XElement? minorGridlines = axis.Element(ChartNamespace + "minorGridlines");
            axes.Add(new PptxSceneChartAxis(
                id,
                ParseChartAxisKind(axis.Name.LocalName),
                axis.Name.LocalName,
                ParseChartAxisPosition(axisPosition),
                axisPosition,
                ReadChartElementValue(axis, "crossAx"),
                ParseChartAxisCrosses(crosses),
                crosses,
                crossesAt,
                crossesAtValue,
                ParseChartAxisCrossBetween(crossBetween),
                crossBetween,
                orientationKind,
                orientation,
                orientationKind == PptxSceneChartAxisOrientation.MaximumMinimum,
                isDeleted,
                isDeletedValue,
                axis.Element(ChartNamespace + "scaling") is not null,
                minimum,
                minimumValue,
                maximum,
                maximumValue,
                majorUnit,
                majorUnitValue,
                minorUnit,
                minorUnitValue,
                IsChartGridlineVisible(majorGridlines),
                IsChartGridlineVisible(minorGridlines),
                majorGridlines is not null,
                minorGridlines is not null,
                ReadChartAxisLine(axis, theme, colorMap),
                ReadChartGridlineLine(majorGridlines, theme, colorMap),
                ReadChartGridlineLine(minorGridlines, theme, colorMap),
                ReadChartStyleRoleLine(stylePart, "gridlineMajor"),
                ReadChartStyleRoleLine(stylePart, "gridlineMinor"),
                ReadChartTextStyleOverride(axis, theme, colorMap),
                ParseChartTickLabelPosition(tickLabelPosition),
                tickLabelPosition,
                ParseChartAxisTickMark(majorTickMark),
                majorTickMark,
                ParseChartAxisTickMark(minorTickMark),
                minorTickMark,
                labelOffset,
                labelOffsetValue,
                tickLabelSkip,
                tickLabelSkipValue,
                tickMarkSkip,
                tickMarkSkipValue,
                noMultiLevelLabels,
                noMultiLevelLabelsValue,
                ReadChartAxisNumberFormat(axis),
                ReadChartNumberFormat(axis),
                ReadChartTitleElement(axis.Element(ChartNamespace + "title"), theme, colorMap, null, "", null)));
        }

        return axes;
    }

    private static PptxSceneLineStyle ReadChartStyleRoleLine(PptxSceneChartStyle stylePart, string role)
    {
        PptxSceneChartStyleEntry entry = stylePart.Entries.FirstOrDefault(item => item.Role == role);
        if (entry.ShapeLine.HasLine)
        {
            return entry.ShapeLine;
        }

        return entry.Line;
    }

    internal static PptxSceneLineStyle ReadChartAxisLine(XElement axis, PptxTheme theme)
    {
        return ReadChartAxisLine(axis, theme, PptxColorMap.Default);
    }

    internal static PptxSceneLineStyle ReadChartAxisLine(XElement axis, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? shapeProperties = axis.Element(ChartNamespace + "spPr");
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneLineStyle(true, new RgbColor(0, 0, 0), 0d, 0d, [], null, null, null, null, null, null, null, true);
        }

        return ReadChartLine(shapeProperties, theme, colorMap);
    }

    internal static PptxSceneChartManualLayout ReadChartPlotAreaManualLayout(XDocument? chartXml)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        return ReadChartManualLayout(plotArea);
    }

    internal static PptxSceneChartManualLayout ReadChartManualLayout(XElement? container)
    {
        XElement? manualLayout = container
            ?.Element(ChartNamespace + "layout")
            ?.Element(ChartNamespace + "manualLayout");
        if (manualLayout is null)
        {
            return default;
        }

        (double? x, string xValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "x");
        (double? y, string yValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "y");
        (double? width, string widthValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "w");
        (double? height, string heightValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "h");
        string layoutTarget = ReadChartElementValue(manualLayout, "layoutTarget");
        string xMode = ReadChartElementValue(manualLayout, "xMode");
        string yMode = ReadChartElementValue(manualLayout, "yMode");
        string widthMode = ReadChartElementValue(manualLayout, "wMode");
        string heightMode = ReadChartElementValue(manualLayout, "hMode");
        return new PptxSceneChartManualLayout(
            true,
            x,
            xValue,
            y,
            yValue,
            width,
            widthValue,
            height,
            heightValue,
            ParseChartManualLayoutTarget(layoutTarget),
            layoutTarget,
            ParseChartManualLayoutMode(xMode),
            xMode,
            ParseChartManualLayoutMode(yMode),
            yMode,
            ParseChartManualLayoutMode(widthMode),
            widthMode,
            ParseChartManualLayoutMode(heightMode),
            heightMode);
    }

    private static (double? Value, string RawValue) ReadChartManualLayoutFactorWithValue(XElement manualLayout, string elementName)
    {
        string value = ReadChartElementValue(manualLayout, elementName);
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    internal static PptxSceneLineStyle ReadChartGridlineLine(XElement? gridlines, PptxTheme theme)
    {
        return ReadChartGridlineLine(gridlines, theme, PptxColorMap.Default);
    }

    internal static PptxSceneLineStyle ReadChartGridlineLine(XElement? gridlines, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? shapeProperties = gridlines?.Element(ChartNamespace + "spPr");
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneLineStyle(true, new RgbColor(0, 0, 0), 0d, 0d, [], null, null, null, null, null, null, null, true);
        }

        return ReadChartLine(shapeProperties, theme, colorMap);
    }

    private static double? ReadChartAxisScalingValue(XElement axis, string elementName)
    {
        (double? parsed, _) = ReadChartAxisScalingValueWithValue(axis, elementName);
        return parsed;
    }

    internal static (double? Value, string RawValue) ReadChartAxisScalingValueWithValue(XElement axis, string elementName)
    {
        string value = ReadChartElementValue(axis.Element(ChartNamespace + "scaling"), elementName);
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static double? ReadChartAxisUnitValue(XElement axis, string elementName)
    {
        (double? parsed, _) = ReadChartAxisUnitValueWithValue(axis, elementName);
        return parsed;
    }

    internal static (double? Value, string RawValue) ReadChartAxisUnitValueWithValue(XElement axis, string elementName)
    {
        string value = ReadChartElementValue(axis, elementName);
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) && parsed > 0d
            ? (parsed, value)
            : (null, value);
    }

    internal static bool IsChartGridlineVisible(XElement? gridlines)
    {
        XElement? line = gridlines?
            .Element(ChartNamespace + "spPr")
            ?.Element(DrawingNamespace + "ln");
        return gridlines is not null && line?.Element(DrawingNamespace + "noFill") is null;
    }

    private static string? ReadChartAxisNumberFormat(XElement axis)
    {
        string? format = (string?)axis
            .Element(ChartNamespace + "numFmt")
            ?.Attribute("formatCode");
        return string.IsNullOrWhiteSpace(format) ? null : format;
    }

    internal static bool IsOoxmlBooleanElementEnabled(XElement? element)
    {
        return OoxBoolean.ParseElement(element, false, null);
    }

    internal static bool IsOoxmlBooleanElementEnabled(XElement? element, bool defaultValue)
    {
        return OoxBoolean.ParseElement(element, defaultValue, null);
    }

}
