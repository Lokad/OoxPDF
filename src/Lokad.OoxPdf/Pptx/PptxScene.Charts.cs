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
        string? relationshipId = (string?)GetGraphicData(frame)
            ?.Element(ChartNamespace + "chart")
            ?.Attribute(RelationshipsNamespace + "id");
        string? targetPartName = relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship)
            ? relationship.ResolvedTarget
            : null;
        OoxPart? chartPart = targetPartName is null ? null : package.GetPart(targetPartName);
        XDocument? chartXml = chartPart is null ? null : package.LoadXml(chartPart, cancellationToken);
        IMemoryReservation? modelStorage = chartXml is null ? null : OoxMemory.ReserveModel(chartXml, cancellationToken);
        PptxSceneChart? result = null;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyDictionary<string, OoxRelationship> chartRelationships = chartPart is null
                ? new Dictionary<string, OoxRelationship>(StringComparer.Ordinal)
                : package.GetRelationshipDictionary(chartPart.Name, cancellationToken);
            PptxSceneChartExternalData externalData = chartPart is null
                ? default
                : ReadChartExternalData(package, chartRelationships, chartXml, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<RgbColor> paletteColors = chartPart is null
                ? []
                : ReadChartPaletteColors(package, chartRelationships, theme, colorMap, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            PptxSceneChartStyle stylePart = chartPart is null
                ? PptxSceneChartStyle.Undefined
                : ReadChartStylePart(package, chartRelationships, theme, colorMap, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            XElement? plotArea = chartXml?.Descendants(ChartNamespace + "plotArea").FirstOrDefault();
            IReadOnlyList<PptxSceneChartPlot> plots = ReadChartPlots(plotArea, theme, colorMap);
            cancellationToken.ThrowIfCancellationRequested();
            XElement? chartSpace = chartXml?.Root;
            string legacyStyleValue = ReadChartElementValue(chartSpace, "style");
            int? preferredGalleryStyle = OoxMarkupCompatibility.ReadPreferredChartGalleryStyle(chartSpace);
            int? galleryStyle = preferredGalleryStyle - 100;
            if (galleryStyle is null && int.TryParse(legacyStyleValue, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int legacyStyle) && legacyStyle is >= 1 and <= 48)
            {
                galleryStyle = legacyStyle;
            }

            // Series recipes use the selected tree. The preferred gallery may
            // instead come from an annotation on an unselected modern choice.
            XElement? directStyle = chartSpace?.Element(ChartStyle2010Namespace + "style");
            string seriesStyleValue = directStyle is null ? legacyStyleValue : ReadChartValueAttribute(directStyle);
            int? seriesStyleId = int.TryParse(seriesStyleValue, out int seriesStyle) ? seriesStyle : null;
            IReadOnlyList<PptxSceneChartAxis> axes = ReadChartAxes(plotArea, theme, colorMap, stylePart, plots, preferredGalleryStyle);
            return result = new PptxSceneChart(
                relationshipId,
                targetPartName,
                chartXml,
                colorMap,
                externalData,
                ReadChartOptions(chartXml),
                paletteColors,
                stylePart,
                legacyStyleValue,
                galleryStyle,
                seriesStyleId,
                plots,
                axes,
                ReadChartTitle(chartXml, theme, colorMap, plots),
                ReadChartLegend(chartXml, theme, colorMap),
                ReadChartTextStyleOverride(chartSpace, theme, colorMap),
                ReadChartManualLayout(plotArea),
                ReadChartShapeStyle(chartSpace?.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadChartShapeStyle(plotArea?.Element(ChartNamespace + "spPr"), theme, colorMap),
                chartRelationships);
        }
        finally
        {
            if (result is null) modelStorage?.Dispose();
            else if (modelStorage is not null) OoxMemory.Resize(modelStorage, modelStorage.Bytes / 2);
        }
    }

    private static PptxSceneChartOptions ReadChartOptions(XDocument? chartXml)
    {
        XElement? chartSpace = chartXml?.Root;
        XElement? chart = chartSpace?.Element(ChartNamespace + "chart");
        return new(
            ReadOptionalChartBoolean(chartSpace, "date1904"),
            ReadOptionalChartBoolean(chart, "plotVisOnly"),
            ParseChartDisplayBlanksAs(ReadChartElementValue(chart, "dispBlanksAs")));
    }

    private static IReadOnlyList<PptxSceneChartPlot> ReadChartPlots(XElement? plotArea, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plotArea is null)
        {
            return [];
        }

        var plots = new List<PptxSceneChartPlot>();
        foreach (XElement plot in plotArea.Elements().Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)))
        {
            string kind = plot.Name.LocalName;
            string grouping = ReadChartElementValue(plot, "grouping");
            string barDirection = ReadChartElementValue(plot, "barDir");
            string scatterStyle = ReadChartElementValue(plot, "scatterStyle");
            string radarStyle = ReadChartElementValue(plot, "radarStyle");
            PptxSceneChartPlotKind plotKind = ParseChartPlotKind(kind);
            string[] axisIds = plot
                .Elements(ChartNamespace + "axId")
                .Select(ReadChartValueAttribute)
                .Where(value => value.Length != 0)
                .ToArray();
            plots.Add(new PptxSceneChartPlot(
                plotKind,
                kind,
                axisIds,
                ReadChartSeries(plot, theme, colorMap, plotKind),
                grouping,
                barDirection,
                scatterStyle,
                radarStyle,
                ReadOptionalChartBoolean(plot, "varyColors"),
                ReadChartElementDoubleWithValue(plot, "gapWidth").Value,
                ReadChartElementDoubleWithValue(plot, "overlap").Value,
                ReadChartElementDoubleWithValue(plot, "holeSize").Value,
                ReadChartElementDoubleWithValue(plot, "firstSliceAng").Value,
                ReadChartDataLabels(plot, theme, colorMap),
                plot));
        }

        return plots;
    }

    internal static PptxSceneChartDataLabels ReadChartDataLabels(XElement plot, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? labels = plot.Element(ChartNamespace + "dLbls");
        if (labels is null)
        {
            return new PptxSceneChartDataLabels(
                ShowValue: null,
                ShowPercent: null,
                ShowCategoryName: null,
                ShowSeriesName: null,
                ShowLeaderLines: null,
                ShowLegendKey: null,
                ShowBubbleSize: null,
                LeaderLines: default,
                Position: null,
                Separator: string.Empty,
                NumberFormatInfo: default,
                Layout: default,
                TextStyle: new PptxSceneChartTextStyleOverride(null, null, null, null, null, null, null, null),
                TextBodyProperties: default,
                ShapeStyle: new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
                Overrides: [],
                IsDefined: false);
        }

        return new PptxSceneChartDataLabels(
                ReadOptionalChartBoolean(labels, "showVal"),
                ReadOptionalChartBoolean(labels, "showPercent"),
                ReadOptionalChartBoolean(labels, "showCatName"),
                ReadOptionalChartBoolean(labels, "showSerName"),
                ReadOptionalChartBoolean(labels, "showLeaderLines"),
                ReadOptionalChartBoolean(labels, "showLegendKey"),
                ReadOptionalChartBoolean(labels, "showBubbleSize"),
                ReadChartLeaderLines(labels, theme, colorMap),
                ParseChartDataLabelPosition(ReadChartElementValue(labels, "dLblPos")),
                labels.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                ReadChartNumberFormat(labels),
                ReadChartManualLayout(labels),
                ReadChartTextStyleOverride(labels, theme, colorMap),
                ReadChartTextBodyProperties(labels),
                ReadChartShapeStyle(labels.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadChartDataLabelOverrides(labels, theme, colorMap),
                IsDefined: true,
                IsDeleted: ReadOptionalChartBoolean(labels, "delete"));
    }

    private static PptxSceneChartLeaderLines ReadChartLeaderLines(XElement labels, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? leaderLines = labels.Element(ChartNamespace + "leaderLines");
        return leaderLines is null
            ? default
            : new PptxSceneChartLeaderLines(true, ReadChartLine(leaderLines.Element(ChartNamespace + "spPr"), theme, colorMap));
    }

    private static IReadOnlyList<PptxSceneChartDataLabelOverride> ReadChartDataLabelOverrides(XElement labels, PptxTheme theme, PptxColorMap colorMap)
    {
        var overrides = new List<PptxSceneChartDataLabelOverride>();
        foreach (XElement label in labels.Elements(ChartNamespace + "dLbl"))
        {
            if (!TryReadChartNonNegativeIndex(label, out int index))
            {
                continue;
            }

            XElement? text = label.Element(ChartNamespace + "tx");
            overrides.Add(new PptxSceneChartDataLabelOverride(
                index,
                ReadOptionalChartBoolean(label, "delete"),
                ReadOptionalChartBoolean(label, "showVal"),
                ReadOptionalChartBoolean(label, "showPercent"),
                ReadOptionalChartBoolean(label, "showCatName"),
                ReadOptionalChartBoolean(label, "showSerName"),
                ReadOptionalChartBoolean(label, "showLeaderLines"),
                ReadOptionalChartBoolean(label, "showLegendKey"),
                ReadOptionalChartBoolean(label, "showBubbleSize"),
                ReadChartLeaderLines(label, theme, colorMap),
                ReadChartText(text, false) ?? string.Empty,
                ReadChartTextRuns(text, theme, colorMap),
                ParseChartDataLabelPosition(ReadChartElementValue(label, "dLblPos")),
                label.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                ReadChartNumberFormat(label),
                ReadChartManualLayout(label),
                ReadChartTextStyleOverride(label, theme, colorMap),
                ReadChartTextBodyProperties(label),
                ReadChartShapeStyle(label.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadChartPointTextScope(label)));
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

    internal static PptxSceneChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        bool noFill = shapeProperties?.Element(DrawingNamespace + "noFill") is not null;
        PptxSceneFillStyle fill = !noFill && PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartShapeStyle(
            noFill,
            fill,
            noFill ? new PptxSceneGradientFill(false, false, 0d, []) : ReadLinearGradientFill(shapeProperties, theme, colorMap),
            noFill ? default : ReadChartPatternFill(shapeProperties, theme, colorMap),
            noFill ? default : ReadDrawingPictureFill(shapeProperties).Style,
            ReadChartLine(shapeProperties, theme, colorMap),
            TryReadGlow(shapeProperties, theme, colorMap, out PptxSceneGlow glow) ? glow : default,
            TryReadOuterShadow(shapeProperties, theme, colorMap, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            ReadEffects(shapeProperties));
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

    internal static (double? Value, string RawValue) ReadChartElementDoubleWithValue(XElement? element, string childName)
    {
        string value = ReadChartElementValue(element, childName);
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    internal static (int? Value, string RawValue) ReadChartElementIntWithValue(XElement? element, string childName)
    {
        string value = ReadChartElementValue(element, childName);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static int? ReadChartCachePointCount(XElement? cache)
    {
        string value = ReadChartValueAttribute(cache?.Element(ChartNamespace + "ptCount"));
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? parsed
            : null;
    }

    private static IReadOnlyList<PptxSceneChartAxis> ReadChartAxes(XElement? plotArea, PptxTheme theme, PptxColorMap colorMap,
        PptxSceneChartStyle stylePart, IReadOnlyList<PptxSceneChartPlot> plots, int? preferredGalleryStyle)
    {
        if (plotArea is null)
        {
            return [];
        }

        bool singleLinePlot = plots.Count == 1 && plots[0].PlotKind == PptxSceneChartPlotKind.Line;
        PptxSceneLineStyle majorStyleLine = ReadChartStyleRoleLine(stylePart, "gridlineMajor");
        PptxSceneLineStyle minorStyleLine = ReadChartStyleRoleLine(stylePart, "gridlineMinor");
        PptxSceneLineStyle axisStyleLine = default;
        if (!stylePart.IsDefined && singleLinePlot &&
            preferredGalleryStyle is 102 or 118 or 126)
        {
            // Independently qualified line-plot gallery roles. Other plot kinds
            // retain their fallback pending geometry qualification. Authored gridline
            // and axis strokes still take precedence. Other galleries and explicit
            // style parts retain their fallback pending independent qualification.
            RgbColor textColor = theme.TryResolveColor("tx1", colorMap, out RgbColor resolvedColor) ? resolvedColor : new RgbColor(0, 0, 0);
            majorStyleLine = new PptxSceneLineStyle(true, LinearLightColor.TintTowardWhite(textColor, 0.25d), 1d, 1d, [], null, null, null, null, null, null, null, true);
            minorStyleLine = new PptxSceneLineStyle(true, LinearLightColor.TintTowardWhite(textColor, 0.5d), 1d, 1d, [], null, null, null, null, null, null, null, true);
            axisStyleLine = majorStyleLine;
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
            XElement? scaling = axis.Element(ChartNamespace + "scaling");
            string orientation = ReadChartElementValue(scaling, "orientation");
            string tickLabelPosition = ReadChartElementValue(axis, "tickLblPos");
            string majorTickMark = ReadChartElementValue(axis, "majorTickMark");
            double? crossesAt = ReadChartElementDoubleWithValue(axis, "crossesAt").Value;
            var minimum = ReadChartElementDoubleWithValue(scaling, "min");
            var maximum = ReadChartElementDoubleWithValue(scaling, "max");
            var majorUnit = ReadChartAxisUnitValueWithValue(axis, "majorUnit");
            double? minorUnit = ReadChartAxisUnitValueWithValue(axis, "minorUnit").Value;
            int? labelOffset = ReadChartElementIntWithValue(axis, "lblOffset").Value;
            (int? tickLabelSkip, string tickLabelSkipValue) = ReadChartElementIntWithValue(axis, "tickLblSkip");
            bool? isDeleted = ReadOptionalChartBoolean(axis, "delete");
            XElement? majorGridlines = axis.Element(ChartNamespace + "majorGridlines");
            XElement? minorGridlines = axis.Element(ChartNamespace + "minorGridlines");
            bool hasLogarithmicScale = scaling?.Element(ChartNamespace + "logBase") is not null;
            axes.Add(PptxSceneChartAxis.Defined(
                id,
                axis.Name.LocalName,
                axisPosition,
                ReadChartElementValue(axis, "crossAx"),
                crosses,
                crossesAt,
                crossBetween,
                orientation,
                isDeleted,
                scaling is not null,
                IsInvalidAxisNumber(minimum) || IsInvalidAxisNumber(maximum) || IsInvalidAxisNumber(majorUnit),
                minimum.Value,
                maximum.Value,
                majorUnit.Value,
                minorUnit,
                IsChartGridlineVisible(majorGridlines),
                IsChartGridlineVisible(minorGridlines),
                majorGridlines is not null,
                ReadChartAxisLine(axis, theme, colorMap, axisStyleLine),
                ReadChartGridlineLine(majorGridlines, theme, colorMap, singleLinePlot ? majorStyleLine : default),
                ReadChartGridlineLine(minorGridlines, theme, colorMap, singleLinePlot && !hasLogarithmicScale ? minorStyleLine : default),
                majorStyleLine,
                minorStyleLine,
                ReadChartTextStyleOverride(axis, theme, colorMap),
                tickLabelPosition,
                majorTickMark,
                labelOffset,
                tickLabelSkip,
                tickLabelSkipValue.Length != 0,
                ReadChartNumberFormat(axis),
                ReadChartTitleElement(axis.Element(ChartNamespace + "title"), theme, colorMap),
                hasLogarithmicScale: hasLogarithmicScale) with
            {
                TextBodyProperties = ReadChartTextBodyProperties(axis)
            });
        }

        return axes;

        // Empty values retain automatic layout. Parsed values are also used by
        // the existing fallback paths when a linear declaration is rejected.
        static bool IsInvalidAxisNumber((double? Value, string RawValue) number) =>
            number.RawValue.Length != 0 && (number.Value is not double value || !double.IsFinite(value));
    }

    private static PptxSceneLineStyle ReadChartStyleRoleLine(PptxSceneChartStyle stylePart, string role)
    {
        PptxSceneChartStyleEntry entry = stylePart.Entries.FirstOrDefault(item => item.Role == role);
        return entry.ShapeStyle?.Line is { HasLine: true } inlineLine ? inlineLine : entry.Line;
    }

    private static PptxSceneLineStyle ReadChartAxisLine(XElement axis, PptxTheme theme, PptxColorMap colorMap, PptxSceneLineStyle inheritedLine)
    {
        PptxSceneLineStyle line = ReadChartLineWithInheritedProperties(axis.Element(ChartNamespace + "spPr"), theme, colorMap, inheritedLine);
        return line.HasLine ? line : inheritedLine;
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
            ParseChartManualLayoutTarget(ReadChartElementValue(manualLayout, "layoutTarget")),
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

    private static PptxSceneLineStyle ReadChartGridlineLine(XElement? gridlines, PptxTheme theme, PptxColorMap colorMap, PptxSceneLineStyle inheritedLine)
    {
        return ReadChartLineWithInheritedProperties(gridlines?.Element(ChartNamespace + "spPr"), theme, colorMap, inheritedLine);
    }

    private static PptxSceneLineStyle ReadChartLineWithInheritedProperties(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, PptxSceneLineStyle inheritedLine)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneLineStyle(true, new RgbColor(0, 0, 0), 0d, 0d, [], null, null, null, null, null, null, null, true);
        }

        if (!inheritedLine.HasLine || line is null)
        {
            return ReadChartLine(shapeProperties, theme, colorMap);
        }

        bool hasFill = line.Elements().Any(element => element.Name.Namespace == DrawingNamespace && element.Name.LocalName.EndsWith("Fill", StringComparison.Ordinal));
        if (hasFill && !PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, out _, out _))
        {
            // Unsupported or unresolved authored fills keep the existing fallback.
            return ReadChartLine(shapeProperties, theme, colorMap);
        }

        // Resolve omissions on a detached line; the authored XML remains intact.
        // Reuse the ordinary reader so width-dependent dash normalization has
        // one owner and explicit color/alpha/width/caps/joins retain precedence.
        var resolvedLine = new XElement(line);
        if (!hasFill)
        {
            var color = new XElement(DrawingNamespace + "srgbClr", new XAttribute("val", $"{inheritedLine.Color.Red:X2}{inheritedLine.Color.Green:X2}{inheritedLine.Color.Blue:X2}"));
            if (inheritedLine.Alpha < 1d)
            {
                color.Add(new XElement(DrawingNamespace + "alpha", new XAttribute("val", (int)Math.Round(inheritedLine.Alpha * 100000d))));
            }
            resolvedLine.AddFirst(new XElement(DrawingNamespace + "solidFill", color));
        }
        if (resolvedLine.Attribute("cap") is null) resolvedLine.SetAttributeValue("cap", inheritedLine.CapValue);
        if (resolvedLine.Attribute("cmpd") is null) resolvedLine.SetAttributeValue("cmpd", inheritedLine.CompoundValue);
        if (resolvedLine.Element(DrawingNamespace + "prstDash") is null && inheritedLine.DashPreset is not null)
        {
            resolvedLine.Add(new XElement(DrawingNamespace + "prstDash", new XAttribute("val", inheritedLine.DashPreset)));
        }
        if (!resolvedLine.Elements().Any(element => element.Name == DrawingNamespace + "round" || element.Name == DrawingNamespace + "bevel" || element.Name == DrawingNamespace + "miter") &&
            inheritedLine.JoinValue is not null)
        {
            resolvedLine.Add(new XElement(DrawingNamespace + inheritedLine.JoinValue));
        }

        PptxSceneLineStyle resolved = ReadChartLine(new XElement(ChartNamespace + "spPr", resolvedLine), theme, colorMap, inheritedLine.Width);
        // Inherited color must not activate a previously ignored invalid width.
        return resolved.Width < 0d ? ReadChartLine(shapeProperties, theme, colorMap) : resolved;
    }

    private static (double? Value, string RawValue) ReadChartAxisUnitValueWithValue(XElement axis, string elementName)
    {
        var (unit, value) = ReadChartElementDoubleWithValue(axis, elementName);
        return (unit is > 0d ? unit : null, value);
    }

    private static bool IsChartGridlineVisible(XElement? gridlines)
    {
        XElement? line = gridlines?
            .Element(ChartNamespace + "spPr")
            ?.Element(DrawingNamespace + "ln");
        return gridlines is not null && line?.Element(DrawingNamespace + "noFill") is null;
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
