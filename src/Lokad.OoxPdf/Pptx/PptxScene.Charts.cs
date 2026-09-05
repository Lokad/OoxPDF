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
                ?.Element(ChartNamespace + "spPr"), theme, colorMap));
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

    private static bool? ReadOptionalOoxmlBooleanElement(XElement parent, string elementName)
    {
        XElement? element = parent.Element(ChartNamespace + elementName);
        return element is null ? null : IsOoxmlBooleanElementEnabled(element);
    }

    private static (bool? Value, string RawValue) ReadOptionalOoxmlBooleanElementWithValue(XElement? parent, string elementName)
    {
        XElement? element = parent?.Element(ChartNamespace + elementName);
        return element is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(element), (string?)element.Attribute("val") ?? string.Empty);
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
        PptxSceneFillStyle fill = !noFill && TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor fillColor, out double fillAlpha)
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

    internal static PptxSceneChartTextStyleOverride ReadChartTextStyleOverride(XElement? parent, PptxTheme theme)
    {
        return ReadChartTextStyleOverride(parent, theme, PptxColorMap.Default);
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartTitleTextStyleOverride(PptxSceneChart chart)
    {
        return ResolveChartElementTextStyleOverride(chart, chart.Title.TextStyle, "title");
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartLegendTextStyleOverride(PptxSceneChart chart)
    {
        return ResolveChartElementTextStyleOverride(chart, chart.Legend.TextStyle, "legend");
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartAxisTextStyleOverride(PptxSceneChart chart, PptxSceneChartAxis? axis, string? chartStyleRole)
    {
        return ResolveChartElementTextStyleOverride(chart, axis?.TextStyle ?? default, chartStyleRole);
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartElementTextStyleOverride(PptxSceneChart chart, PptxSceneChartTextStyleOverride elementTextStyle, string? chartStyleRole)
    {
        PptxSceneChartTextStyleOverride style = chart.TextStyle;
        if (!string.IsNullOrWhiteSpace(chartStyleRole))
        {
            style = MergeChartTextStyleOverride(style, ReadChartStyleRoleTextStyleOverride(chart.StylePart, chartStyleRole));
        }

        return MergeChartTextStyleOverride(style, elementTextStyle);
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartDataLabelTextStyleOverride(PptxSceneChart chart, PptxSceneChartDataLabels labels)
    {
        return ResolveChartElementTextStyleOverride(chart, labels.TextStyle, "dataLabel");
    }

    internal static PptxSceneChartTextStyleOverride ReadChartStyleRoleTextStyleOverride(PptxSceneChartStyle stylePart, string role)
    {
        PptxSceneChartStyleEntry entry = stylePart.Entries.FirstOrDefault(item => item.Role == role);
        return entry.TextStyle;
    }

    internal static PptxSceneChartTextStyleOverride MergeChartTextStyleOverride(PptxSceneChartTextStyleOverride style, PptxSceneChartTextStyleOverride next)
    {
        return new PptxSceneChartTextStyleOverride(
            next.FontFamily ?? style.FontFamily,
            next.FontFamily is null ? style.RequestedTypeface : next.RequestedTypeface,
            next.FontFamily is null ? style.TypefaceSource : next.TypefaceSource,
            next.FontSize ?? style.FontSize,
            next.CharacterSpacing ?? style.CharacterSpacing,
            next.Color ?? style.Color,
            next.Alpha ?? style.Alpha,
            next.Bold ?? style.Bold,
            next.Italic ?? style.Italic,
            next.Underline ?? style.Underline,
            next.Strike ?? style.Strike);
    }

    internal static PptxSceneChartTextStyleOverride ReadChartTextStyleOverride(XElement? parent, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? defaultRunProperties = parent?
            .Element(ChartNamespace + "txPr")?
            .Elements(DrawingNamespace + "p")
            .Select(paragraph => paragraph.Element(DrawingNamespace + "pPr")?.Element(DrawingNamespace + "defRPr"))
            .FirstOrDefault(element => element is not null);
        if (defaultRunProperties is null)
        {
            return default;
        }

        string? typeface = (string?)defaultRunProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        PptxThemeTypefaceResolution typefaceResolution = string.IsNullOrWhiteSpace(typeface)
            ? default
            : theme.ResolveTypefaceWithSource(typeface);
        string? fontFamily = typefaceResolution.Typeface;
        double? fontSize = defaultRunProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        double? characterSpacing = ReadOptionalChartCharacterSpacing(defaultRunProperties);
        RgbColor? color = TryReadSolidColorWithAlpha(defaultRunProperties.Element(DrawingNamespace + "solidFill"), theme, colorMap, out RgbColor parsedColor, out double alpha)
            ? parsedColor
            : null;
        bool? bold = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "b");
        bool? italic = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "i");
        bool? underline = ReadChartUnderline(defaultRunProperties);
        bool? strike = ReadChartStrike(defaultRunProperties);
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

    internal static PptxSceneChartTextBodyProperties ReadChartTextBodyProperties(XElement? parent)
    {
        XElement? bodyProperties = parent?
            .Element(ChartNamespace + "txPr")?
            .Element(DrawingNamespace + "bodyPr");
        string rotation = (string?)bodyProperties?.Attribute("rot") ?? string.Empty;
        return new PptxSceneChartTextBodyProperties(
            ParseOptionalOoxmlAngle(rotation),
            rotation,
            (string?)bodyProperties?.Attribute("vert") ?? string.Empty,
            (string?)bodyProperties?.Attribute("vertOverflow") ?? string.Empty);
    }

    private static double? ParseOptionalOoxmlAngle(string value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rawAngle)
            ? rawAngle / 60000d
            : null;
    }

    private static PptxSceneChartTextStyleOverride ReadChartTextRunStyle(XElement? runProperties, PptxTheme theme)
    {
        return ReadChartTextRunStyle(runProperties, theme, PptxColorMap.Default);
    }

    private static PptxSceneChartTextStyleOverride ReadChartTextRunStyle(XElement? runProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        if (runProperties is null)
        {
            return default;
        }

        string? typeface = (string?)runProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)runProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)runProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        PptxThemeTypefaceResolution typefaceResolution = string.IsNullOrWhiteSpace(typeface)
            ? default
            : theme.ResolveTypefaceWithSource(typeface);
        string? fontFamily = typefaceResolution.Typeface;
        double? fontSize = runProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        double? characterSpacing = ReadOptionalChartCharacterSpacing(runProperties);
        RgbColor? color = TryReadSolidColorWithAlpha(runProperties.Element(DrawingNamespace + "solidFill"), theme, colorMap, out RgbColor parsedColor, out double alpha)
            ? parsedColor
            : null;
        bool? bold = ReadOptionalOoxmlBooleanAttribute(runProperties, "b");
        bool? italic = ReadOptionalOoxmlBooleanAttribute(runProperties, "i");
        bool? underline = ReadChartUnderline(runProperties);
        bool? strike = ReadChartStrike(runProperties);
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

    private static bool? ReadChartUnderline(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("u");
        return value is null ? null : !value.Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadChartStrike(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("strike");
        return value is null ? null : !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
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

    private static PptxSceneChartTitle ReadChartTitle(XDocument? chartXml, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<PptxSceneChartPlot> plots)
    {
        XElement? chart = chartXml?
            .Descendants(ChartNamespace + "chart")
            .FirstOrDefault();
        if (chart is null)
        {
            return EmptyChartTitle(IsAutoDeleted: null, IsAutoDeletedValue: "");
        }

        (bool? isAutoDeleted, string isAutoDeletedValue) = ReadOptionalOoxmlBooleanElementWithValue(chart, "autoTitleDeleted");
        return ReadChartTitleElement(chart.Element(ChartNamespace + "title"), theme, colorMap, isAutoDeleted, isAutoDeletedValue, plots);
    }

    private static PptxSceneChartTitle ReadChartTitleElement(XElement? title, PptxTheme theme, bool? isAutoDeleted, string isAutoDeletedValue, IReadOnlyList<PptxSceneChartPlot>? plots)
    {
        return ReadChartTitleElement(title, theme, PptxColorMap.Default, isAutoDeleted, isAutoDeletedValue, plots);
    }

    private static PptxSceneChartTitle ReadChartTitleElement(XElement? title, PptxTheme theme, PptxColorMap colorMap, bool? isAutoDeleted, string isAutoDeletedValue, IReadOnlyList<PptxSceneChartPlot>? plots)
    {
        if (title is null)
        {
            if (isAutoDeleted != false || plots is null)
            {
                return EmptyChartTitle(isAutoDeleted, isAutoDeletedValue);
            }

            string? inferredText = InferAutoChartTitleText(plots);
            return string.IsNullOrWhiteSpace(inferredText)
                ? EmptyChartTitle(isAutoDeleted, isAutoDeletedValue)
                : new PptxSceneChartTitle(
                    inferredText,
                    [new PptxSceneChartTextRun(inferredText, default)],
                    isAutoDeleted,
                    isAutoDeletedValue,
                    IsAutoGenerated: true,
                    Overlay: null,
                    OverlayValue: string.Empty,
                    default,
                    new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
                    default,
                    default);
        }

        (bool? overlay, string overlayValue) = ReadOptionalOoxmlBooleanElementWithValue(title, "overlay");
        XElement? textElement = title?.Element(ChartNamespace + "tx");
        string? text = ReadChartText(textElement, false);

        return new PptxSceneChartTitle(
            string.IsNullOrWhiteSpace(text) ? null : text,
            ReadChartTextRuns(textElement, theme, colorMap),
            isAutoDeleted,
            isAutoDeletedValue,
            IsAutoGenerated: false,
            overlay,
            overlayValue,
            ReadChartManualLayout(title),
            ReadChartShapeStyle(title?.Element(ChartNamespace + "spPr"), theme, colorMap),
            ReadChartTextBodyProperties(title),
            ReadChartTextStyleOverride(title, theme, colorMap));

        string? InferAutoChartTitleText(IReadOnlyList<PptxSceneChartPlot> plots)
        {
            IReadOnlyList<PptxSceneChartSeries> series = plots
                .SelectMany(plot => plot.Series)
                .ToArray();
            if (series.Count != 1)
            {
                return null;
            }

            string? name = series[0].Name?.Trim();
            return string.IsNullOrWhiteSpace(name) ? null : name;
        }
    }

    private static PptxSceneChartTitle EmptyChartTitle(bool? IsAutoDeleted, string IsAutoDeletedValue)
    {
        return new PptxSceneChartTitle(
            null,
            [],
            IsAutoDeleted,
            IsAutoDeletedValue,
            IsAutoGenerated: false,
            Overlay: null,
            OverlayValue: string.Empty,
            default,
            new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
            default,
            default);
    }

    internal static PptxSceneChartLegend ReadChartLegend(XDocument? chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? legend = chartXml?
            .Descendants(ChartNamespace + "legend")
            .FirstOrDefault();
        if (legend is null)
        {
            return new PptxSceneChartLegend(
                PptxSceneChartLegendPosition.Right,
                "r",
                Overlay: null,
                OverlayValue: string.Empty,
                IsDefined: false,
                IsDeleted: null,
                IsDeletedValue: string.Empty,
                default,
                new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
                default,
                default);
        }

        string position = ReadOptionalChartValueAttribute(legend.Element(ChartNamespace + "legendPos")) ?? "r";
        (bool? overlay, string overlayValue) = ReadOptionalOoxmlBooleanElementWithValue(legend, "overlay");
        (bool? isDeleted, string isDeletedValue) = ReadOptionalOoxmlBooleanElementWithValue(legend, "delete");
        return new PptxSceneChartLegend(
            ParseChartLegendPosition(position),
            position,
            overlay,
            overlayValue,
            IsDefined: true,
            isDeleted,
            isDeletedValue,
            ReadChartManualLayout(legend),
            ReadChartShapeStyle(legend.Element(ChartNamespace + "spPr"), theme, colorMap),
            ReadChartTextBodyProperties(legend),
            ReadChartTextStyleOverride(legend, theme, colorMap));
    }

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
            if (TryReadSolidColorWithAlpha(wrapper, theme, colorMap, out RgbColor color, out _))
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
        bool isResolved = TryReadSolidColorWithAlpha(wrapper, theme, colorMap, out RgbColor color, out double alpha);
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
            TryReadSolidColorWithAlpha(fillStyle, theme, colorMap, fillReference, out RgbColor color, out double alpha))
        {
            return new PptxSceneFillStyle(true, color, alpha);
        }

        return fillReferenceIndex > 0 &&
            TryReadSolidColorWithAlpha(fillReference, theme, colorMap, out color, out alpha)
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
        RgbColor? color = TryReadSolidColorWithAlpha(defaultRunProperties?.Element(DrawingNamespace + "solidFill"), theme, colorMap, out RgbColor parsedColor, out double alpha) ||
            TryReadSolidColorWithAlpha(fontReference, theme, colorMap, out parsedColor, out alpha)
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

    internal static bool IsOoxmlBooleanElementEnabled(XElement? element)
    {
        return OoxBoolean.ParseElement(element, false, null);
    }

    internal static bool IsOoxmlBooleanElementEnabled(XElement? element, bool defaultValue)
    {
        return OoxBoolean.ParseElement(element, defaultValue, null);
    }

    private static bool? ReadOptionalOoxmlBooleanAttribute(XElement element, string attributeName)
    {
        return OoxBoolean.ParseOptionalAttribute(element, attributeName);
    }

}
