using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    internal static IReadOnlyList<PptxSceneChartSeries> ReadChartSeries(XElement plot, PptxTheme theme, PptxColorMap colorMap, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled)
    {
        var series = new List<PptxSceneChartSeries>();
        foreach (XElement seriesElement in plot.Elements(ChartNamespace + "ser"))
        {
            int seriesIndex = series.Count;
            (int? index, string indexValue) = ReadChartElementIntWithValue(seriesElement, "idx");
            (int? order, string orderValue) = ReadChartElementIntWithValue(seriesElement, "order");
            (double? explosion, string explosionValue) = ReadChartElementDoubleWithValue(seriesElement, "explosion");
            (int? valuePointCount, string valuePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "val");
            (int? categoryPointCount, string categoryPointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "cat");
            (int? xValuePointCount, string xValuePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "xVal");
            (int? yValuePointCount, string yValuePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "yVal");
            (int? bubbleSizePointCount, string bubbleSizePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "bubbleSize");
            (bool? smooth, string smoothValue) = ReadChartSeriesSmooth(seriesElement);
            XElement? shapeProperties = seriesElement.Element(ChartNamespace + "spPr");
            series.Add(new PptxSceneChartSeries(
                index,
                indexValue,
                order,
                orderValue,
                ReadChartSeriesName(seriesElement),
                ReadChartSeriesDataSources(seriesElement),
                ReadChartSeriesValues(seriesElement),
                ReadChartSeriesNumberPoints(seriesElement, "val"),
                valuePointCount,
                valuePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "val"),
                ReadChartSeriesCategories(seriesElement),
                ReadChartSeriesStringPoints(seriesElement, "cat"),
                categoryPointCount,
                categoryPointCountValue,
                ReadChartSeriesStringLevels(seriesElement, "cat"),
                ReadChartSeriesNumbers(seriesElement, "xVal"),
                ReadChartSeriesNumberPoints(seriesElement, "xVal"),
                xValuePointCount,
                xValuePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "xVal"),
                ReadChartSeriesNumbers(seriesElement, "yVal"),
                ReadChartSeriesNumberPoints(seriesElement, "yVal"),
                yValuePointCount,
                yValuePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "yVal"),
                ReadChartSeriesNumbers(seriesElement, "bubbleSize"),
                ReadChartSeriesNumberPoints(seriesElement, "bubbleSize"),
                bubbleSizePointCount,
                bubbleSizePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "bubbleSize"),
                ReadChartSeriesFill(seriesElement, theme, colorMap),
                ReadChartSeriesPatternFill(seriesElement, theme, colorMap),
                ReadChartSeriesLine(seriesElement, theme, colorMap),
                ReadChartEffects(shapeProperties),
                ReadChartMarker(seriesElement, theme, colorMap, plotKind, chartMarkersEnabled, seriesIndex),
                ReadChartPointStyles(seriesElement, theme, colorMap),
                ReadRejectedChartNonNegativeIndexValues(seriesElement, "dPt"),
                explosion,
                explosionValue,
                smooth,
                smoothValue,
                ReadChartDataLabels(seriesElement, theme, colorMap)));
        }

        return series;
    }

    private static PptxSceneFillStyle ReadChartSeriesFill(XElement series, PptxTheme theme)
    {
        return ReadChartSeriesFill(series, theme, PptxColorMap.Default);
    }

    private static PptxSceneFillStyle ReadChartSeriesFill(XElement series, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? shapeProperties = series.Element(ChartNamespace + "spPr");
        return PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxSceneLineStyle ReadChartSeriesLine(XElement series, PptxTheme theme)
    {
        return ReadChartSeriesLine(series, theme, PptxColorMap.Default);
    }

    private static PptxSceneLineStyle ReadChartSeriesLine(XElement series, PptxTheme theme, PptxColorMap colorMap)
    {
        return ReadChartLine(series.Element(ChartNamespace + "spPr"), theme, colorMap);
    }

    private static PptxScenePatternFill ReadChartSeriesPatternFill(XElement series, PptxTheme theme)
    {
        return ReadChartSeriesPatternFill(series, theme, PptxColorMap.Default);
    }

    private static PptxScenePatternFill ReadChartSeriesPatternFill(XElement series, PptxTheme theme, PptxColorMap colorMap)
    {
        return ReadChartPatternFill(series.Element(ChartNamespace + "spPr"), theme, colorMap);
    }

    private static PptxScenePatternFill ReadChartPatternFill(XElement? shapeProperties, PptxTheme theme)
    {
        return ReadChartPatternFill(shapeProperties, theme, PptxColorMap.Default);
    }

    private static PptxScenePatternFill ReadChartPatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        if (patternFill is null)
        {
            return default;
        }

        RgbColor foreground = PptxColorResolver.TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, colorMap, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = PptxColorResolver.TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, colorMap, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        return new PptxScenePatternFill(
            HasPattern: true,
            HasPatternSource: true,
            HasUnsupportedPattern: false,
            Preset: (string?)patternFill.Attribute("prst") ?? "pct50",
            Foreground: foreground,
            Background: background,
            Alpha: 1d);
    }

    private static PptxSceneChartMarker ReadChartMarker(XElement series, PptxTheme theme, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, int seriesIndex)
    {
        return ReadChartMarker(series, theme, PptxColorMap.Default, plotKind, chartMarkersEnabled, seriesIndex);
    }

    private static PptxSceneChartMarker ReadChartMarker(XElement series, PptxTheme theme, PptxColorMap colorMap, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, int seriesIndex)
    {
        XElement? marker = series.Element(ChartNamespace + "marker");
        string symbol = ReadOptionalChartValueAttribute(marker?.Element(ChartNamespace + "symbol")) ??
            ReadDefaultChartMarkerSymbol(plotKind, chartMarkersEnabled, seriesIndex);
        string? sizeValue = ReadOptionalChartValueAttribute(marker?.Element(ChartNamespace + "size"));
        double size = PptxChartMarkerMetricRules.ResolveSize(
            sizeValue,
            plotKind,
            chartMarkersEnabled,
            marker is not null,
            marker?.Element(ChartNamespace + "spPr") is not null);
        XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
        PptxSceneFillStyle fill = PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartMarker(marker is not null, ParseChartMarkerSymbol(symbol), symbol, sizeValue, size, fill, ReadChartLine(shapeProperties, theme, colorMap));
    }

    private static string ReadDefaultChartMarkerSymbol(PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, int seriesIndex)
    {
        return PptxChartMarkerMetricRules.ResolveDefaultSymbol(plotKind, chartMarkersEnabled, seriesIndex);
    }

    private static IReadOnlyList<PptxSceneChartPointStyle> ReadChartPointStyles(XElement series, PptxTheme theme)
    {
        return ReadChartPointStyles(series, theme, PptxColorMap.Default);
    }

    private static IReadOnlyList<PptxSceneChartPointStyle> ReadChartPointStyles(XElement series, PptxTheme theme, PptxColorMap colorMap)
    {
        var styles = new List<PptxSceneChartPointStyle>();
        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (!TryReadChartNonNegativeIndex(point, out int index, out string indexValue))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            (double? explosion, string explosionValue) = ReadChartElementDoubleWithValue(point, "explosion");
            styles.Add(new PptxSceneChartPointStyle(
                index,
                indexValue,
                ReadChartPointFill(shapeProperties, theme, colorMap),
                ReadChartPointPatternFill(shapeProperties, theme, colorMap),
                ReadChartLine(shapeProperties, theme, colorMap),
                ReadChartEffects(shapeProperties),
                explosion,
                explosionValue));
        }

        return styles;
    }

    internal static bool TryReadChartNonNegativeIndex(XElement element, out int index, out string indexValue)
    {
        indexValue = ReadChartElementValue(element, "idx");
        return int.TryParse(indexValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out index) &&
            index >= 0;
    }

    private static IReadOnlyList<string> ReadRejectedChartNonNegativeIndexValues(XElement parent, string elementName)
    {
        var rejected = new List<string>();
        foreach (XElement element in parent.Elements(ChartNamespace + elementName))
        {
            if (!TryReadChartNonNegativeIndex(element, out _, out string indexValue))
            {
                rejected.Add(indexValue);
            }
        }

        return rejected;
    }

    private static PptxSceneFillStyle ReadChartPointFill(XElement? shapeProperties, PptxTheme theme)
    {
        return ReadChartPointFill(shapeProperties, theme, PptxColorMap.Default);
    }

    private static PptxSceneFillStyle ReadChartPointFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        return PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxScenePatternFill ReadChartPointPatternFill(XElement? shapeProperties, PptxTheme theme)
    {
        return ReadChartPointPatternFill(shapeProperties, theme, PptxColorMap.Default);
    }

    private static PptxScenePatternFill ReadChartPointPatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        return ReadChartPatternFill(shapeProperties, theme, colorMap);
    }

    private static PptxSceneLineStyle ReadChartLine(XElement? shapeProperties, PptxTheme theme)
    {
        return ReadChartLine(shapeProperties, theme, PptxColorMap.Default);
    }

    private static PptxSceneLineStyle ReadChartLine(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return default;
        }

        bool widthSpecified = line?.Attribute("w") is not null;
        return shapeProperties is not null &&
            PptxLineStyleReader.TryReadLineWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double lineWidth, out double alpha, fallbackLineWidth: null)
                ? new PptxSceneLineStyle(
                    true,
                    color,
                    lineWidth,
                    alpha,
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
                    ReadLineJoinValue(shapeProperties),
                    widthSpecified)
                : default;
    }

    internal static (bool? Value, string RawValue) ReadChartSeriesSmooth(XElement series)
    {
        XElement? smooth = series.Element(ChartNamespace + "smooth");
        return smooth is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(smooth), (string?)smooth.Attribute("val") ?? string.Empty);
    }

    private static string? ReadChartSeriesName(XElement series)
    {
        return ReadChartText(series.Element(ChartNamespace + "tx"), false);
    }

    private static PptxSceneChartSeriesDataSources ReadChartSeriesDataSources(XElement series)
    {
        return new PptxSceneChartSeriesDataSources(
            ReadChartDataSource(series.Element(ChartNamespace + "tx"), "strRef", "numRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "val"), "numRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "cat"), "strRef", "numRef", "multiLvlStrRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "xVal"), "numRef", "strRef", "multiLvlStrRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "yVal"), "numRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "bubbleSize"), "numRef"));
    }

    internal static PptxSceneChartDataSource ReadChartDataSource(XElement? container, params string[] referenceKinds)
    {
        if (container is null)
        {
            return default;
        }

        foreach (string referenceKind in referenceKinds)
        {
            XElement? reference = container.Element(ChartNamespace + referenceKind);
            if (reference is null)
            {
                continue;
            }

            XElement? cache = reference
                .Elements()
                .FirstOrDefault(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Cache", StringComparison.Ordinal));
            bool hasCachedPoints = cache?
                .Descendants(ChartNamespace + "pt")
                .Any() == true;
            return new PptxSceneChartDataSource(
                (string?)reference.Element(ChartNamespace + "f"),
                ParseChartDataSourceReferenceKind(referenceKind),
                referenceKind,
                ParseChartDataSourceCacheKind(cache?.Name.LocalName),
                cache?.Name.LocalName ?? string.Empty,
                hasCachedPoints);
        }

        return default;
    }

    internal static string? ReadChartText(XElement? text, bool trimLiteral)
    {
        string? literal = text?
            .Descendants(ChartNamespace + "v")
            .Select(value => trimLiteral ? value.Value.Trim() : value.Value)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!string.IsNullOrWhiteSpace(literal))
        {
            return literal;
        }

        string? richText = text?
            .Descendants(DrawingNamespace + "t")
            .Aggregate(string.Empty, (current, textElement) => current + textElement.Value);
        return string.IsNullOrWhiteSpace(richText) ? null : richText;
    }

    private static string? ReadChartHyperlinkClickId(XElement? runProperties)
    {
        return (string?)runProperties
            ?.Element(DrawingNamespace + "hlinkClick")
            ?.Attribute(RelationshipsNamespace + "id");
    }

    private static string? ReadChartHyperlinkClickAction(XElement? runProperties)
    {
        return (string?)runProperties
            ?.Element(DrawingNamespace + "hlinkClick")
            ?.Attribute("action");
    }

    internal static IReadOnlyList<PptxSceneChartTextRun> ReadChartTextRuns(XElement? text, PptxTheme theme)
    {
        return ReadChartTextRuns(text, theme, PptxColorMap.Default);
    }

    internal static IReadOnlyList<PptxSceneChartTextRun> ReadChartTextRuns(XElement? text, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? rich = text?.Element(ChartNamespace + "rich");
        if (rich is null)
        {
            string? literal = text?
                .Descendants(ChartNamespace + "v")
                .Select(value => value.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            return string.IsNullOrWhiteSpace(literal)
                ? []
                : [new PptxSceneChartTextRun(literal, default)];
        }

        var runs = new List<PptxSceneChartTextRun>();
        foreach (XElement run in rich.Descendants(DrawingNamespace + "r"))
        {
            string runText = run.Element(DrawingNamespace + "t")?.Value ?? string.Empty;
            if (runText.Length == 0)
            {
                continue;
            }

            XElement? runProperties = run.Element(DrawingNamespace + "rPr");
            runs.Add(new PptxSceneChartTextRun(runText, ReadChartTextRunStyle(runProperties, theme, colorMap), HyperlinkClickId: ReadChartHyperlinkClickId(runProperties), HyperlinkClickAction: ReadChartHyperlinkClickAction(runProperties)));
        }

        return runs;
    }

    private static IReadOnlyList<double> ReadChartSeriesValues(XElement series)
    {
        return ReadChartSeriesNumbers(series, "val");
    }

    private static IReadOnlyList<PptxSceneChartNumberPoint> ReadChartSeriesNumberPoints(XElement series, string elementName)
    {
        return ReadChartNumberPoints(series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "pt"), false);
    }

    internal static IReadOnlyList<PptxSceneChartNumberPoint> ReadChartNumberPoints(IEnumerable<XElement> sourcePoints, bool requireNonNegativeIndex)
    {
        var points = new List<PptxSceneChartNumberPoint>();
        int ordinal = 0;
        foreach (XElement point in sourcePoints)
        {
            (int? parsedIndex, string indexValue) = ReadChartPointIndexAttribute(point, requireNonNegativeIndex);
            int index = parsedIndex ?? ordinal;
            XElement? valueElement = point.Element(ChartNamespace + "v");
            string text = (string?)valueElement ?? string.Empty;
            double? value = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : null;
            points.Add(new PptxSceneChartNumberPoint(index, indexValue, parsedIndex is not null, value, text, valueElement is not null));
            ordinal++;
        }

        return points;
    }

    private static int? ReadChartSeriesPointCount(XElement series, string elementName)
    {
        (int? parsed, _) = ReadChartSeriesPointCountWithValue(series, elementName);
        return parsed;
    }

    private static (int? Value, string RawValue) ReadChartSeriesPointCountWithValue(XElement series, string elementName)
    {
        XElement? cache = series
            .Elements(ChartNamespace + elementName)
            .Descendants()
            .FirstOrDefault(child => child.Name.Namespace == ChartNamespace &&
                (child.Name.LocalName == "numLit" ||
                 child.Name.LocalName == "numCache" ||
                 child.Name.LocalName == "strLit" ||
                 child.Name.LocalName == "strCache" ||
                 child.Name.LocalName == "multiLvlStrCache"));
        return ReadChartCachePointCount(cache);
    }

    private static string? ReadChartSeriesNumberFormatCode(XElement series, string elementName)
    {
        XElement? cache = series
            .Elements(ChartNamespace + elementName)
            .Descendants()
            .FirstOrDefault(child => child.Name.Namespace == ChartNamespace &&
                (child.Name.LocalName == "numLit" ||
                 child.Name.LocalName == "numCache"));
        return (string?)cache?.Element(ChartNamespace + "formatCode");
    }

    private static IReadOnlyList<double> ReadChartSeriesNumbers(XElement series, string elementName)
    {
        return series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
            .Where(value => !double.IsNaN(value))
            .ToArray();
    }

    private static IReadOnlyList<string> ReadChartSeriesCategories(XElement series)
    {
        return series
            .Elements(ChartNamespace + "cat")
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .OfType<string>()
            .ToArray();
    }

    private static IReadOnlyList<PptxSceneChartStringPoint> ReadChartSeriesStringPoints(XElement series, string elementName)
    {
        return ReadChartStringPoints(series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "pt"), false);
    }

    private static IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> ReadChartSeriesStringLevels(XElement series, string elementName)
    {
        return series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "lvl")
            .Select(level => ReadChartStringPoints(level.Elements(ChartNamespace + "pt"), false))
            .Where(points => points.Count != 0)
            .ToArray();
    }

    internal static IReadOnlyList<PptxSceneChartStringPoint> ReadChartStringPoints(IEnumerable<XElement> sourcePoints, bool requireNonNegativeIndex)
    {
        var points = new List<PptxSceneChartStringPoint>();
        int ordinal = 0;
        foreach (XElement point in sourcePoints)
        {
            (int? parsedIndex, string indexValue) = ReadChartPointIndexAttribute(point, requireNonNegativeIndex);
            int index = parsedIndex ?? ordinal;
            XElement? valueElement = point.Element(ChartNamespace + "v");
            points.Add(new PptxSceneChartStringPoint(index, indexValue, parsedIndex is not null, valueElement?.Value ?? string.Empty, valueElement is not null));
            ordinal++;
        }

        return points;
    }
}
