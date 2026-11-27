using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    internal static IReadOnlyList<PptxSceneChartSeries> ReadChartSeries(XElement plot, PptxTheme theme, PptxColorMap colorMap, PptxSceneChartPlotKind plotKind)
    {
        var sources = plot.Elements(ChartNamespace + "ser").Select(element =>
        {
            (int? order, string raw) = ReadChartElementIntWithValue(element, "order");
            return (Element: element, Order: order, OrderValue: raw);
        }).ToArray();
        // Resolve display order before position-dependent defaults. Explicit
        // identity/style/data stay with their source; ambiguous orders retain
        // the complete XML-order fallback rather than sorting a partial set.
        if (sources.All(item => item.Order is >= 0) &&
            sources.Select(item => item.Order).Distinct().Count() == sources.Length)
            Array.Sort(sources, static (left, right) => Nullable.Compare(left.Order, right.Order));

        var series = new List<PptxSceneChartSeries>(sources.Length);
        foreach (var source in sources)
        {
            XElement seriesElement = source.Element;
            int seriesIndex = series.Count;
            (int? index, string indexValue) = ReadChartElementIntWithValue(seriesElement, "idx");
            (double? explosion, string explosionValue) = ReadChartElementDoubleWithValue(seriesElement, "explosion");
            XElement? shapeProperties = seriesElement.Element(ChartNamespace + "spPr");
            PptxSceneChartMarker seriesMarker = ReadChartMarker(seriesElement, theme, colorMap, plotKind, seriesIndex);
            series.Add(new PptxSceneChartSeries(
                index,
                indexValue,
                source.Order,
                source.OrderValue,
                ReadChartSeriesName(seriesElement),
                ReadChartSeriesFormulas(seriesElement),
                ReadChartSeriesNumberPoints(seriesElement, "val"),
                ReadChartSeriesPointCount(seriesElement, "val"),
                ReadChartSeriesNumberFormatCode(seriesElement, "val"),
                ReadChartSeriesStringPoints(seriesElement, "cat"),
                ReadChartSeriesPointCount(seriesElement, "cat"),
                ReadChartSeriesStringLevels(seriesElement, "cat"),
                ReadChartSeriesNumberPoints(seriesElement, "xVal"),
                ReadChartSeriesPointCount(seriesElement, "xVal"),
                ReadChartSeriesNumberFormatCode(seriesElement, "xVal"),
                ReadChartSeriesNumberPoints(seriesElement, "yVal"),
                ReadChartSeriesPointCount(seriesElement, "yVal"),
                ReadChartSeriesNumberFormatCode(seriesElement, "yVal"),
                ReadChartSeriesNumberPoints(seriesElement, "bubbleSize"),
                ReadChartSeriesPointCount(seriesElement, "bubbleSize"),
                ReadChartSeriesNumberFormatCode(seriesElement, "bubbleSize"),
                ReadChartSeriesFill(seriesElement, theme, colorMap),
                ReadChartPatternFill(shapeProperties, theme, colorMap),
                ReadChartSeriesLine(seriesElement, theme, colorMap),
                ReadEffects(shapeProperties),
                seriesMarker,
                ReadChartPointStyles(seriesElement, theme, colorMap, plotKind, seriesIndex, seriesMarker),
                explosion,
                explosionValue,
                ReadOptionalChartBoolean(seriesElement, "smooth"),
                ReadChartDataLabels(seriesElement, theme, colorMap))
            {
                NegativeInversion = ReadChartNegativeInversion(seriesElement)
            });
        }

        return series;
    }

    internal static PptxSceneChartNegativeInversion ReadChartNegativeInversion(XElement series)
    {
        bool? seriesValue = ReadOptionalChartBoolean(series, "invertIfNegative");
        Dictionary<int, bool>? points = null;
        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (TryReadChartNonNegativeIndex(point, out int index))
            {
                points ??= new Dictionary<int, bool>();
                // A point declaration has its own enabled default, including
                // declarations that only override fill or line. Office does not
                // inherit the series flag for an existing dPt with no flag.
                points[index] = ReadOptionalChartBoolean(point, "invertIfNegative") ?? true;
            }
        }
        return new PptxSceneChartNegativeInversion(seriesValue, points);
    }

    private static PptxSceneFillStyle ReadChartSeriesFill(XElement series, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? shapeProperties = series.Element(ChartNamespace + "spPr");
        return PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxSceneLineStyle ReadChartSeriesLine(XElement series, PptxTheme theme, PptxColorMap colorMap)
    {
        return ReadChartLine(series.Element(ChartNamespace + "spPr"), theme, colorMap);
    }

    private static PptxScenePatternFill ReadChartPatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        if (patternFill is null)
        {
            return default;
        }

        return ReadPatternFill(patternFill, (string?)patternFill.Attribute("prst") ?? "pct50", theme, colorMap);
    }

    private static PptxSceneChartMarker ReadChartMarker(XElement series, PptxTheme theme, PptxColorMap colorMap, PptxSceneChartPlotKind plotKind, int seriesIndex)
    {
        XElement? marker = series.Element(ChartNamespace + "marker");
        XElement? symbolElement = marker?.Element(ChartNamespace + "symbol");
        XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
        string symbol = ReadOptionalChartValueAttribute(symbolElement) ??
            PptxChartMarkerMetricRules.ResolveDefaultSymbol(plotKind, seriesIndex);
        string? sizeValue = ReadOptionalChartValueAttribute(marker?.Element(ChartNamespace + "size"));
        int scatterPointCount = plotKind == PptxSceneChartPlotKind.Scatter
            ? Math.Max(ReadChartSeriesPointCount(series, "xVal") ?? 0, ReadChartSeriesPointCount(series, "yVal") ?? 0)
            : 0;
        double size = PptxChartMarkerMetricRules.ResolveSize(
            sizeValue,
            plotKind,
            marker is not null,
            shapeProperties is not null,
            scatterPointCount,
            symbolElement is not null);
        PptxSceneFillStyle fill = PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartMarker(marker is not null, ParseChartMarkerSymbol(symbol), symbol, sizeValue, size, fill, ReadChartLine(shapeProperties, theme, colorMap))
        {
            SymbolSpecified = symbolElement is not null
        };
    }

    private static IReadOnlyList<PptxSceneChartPointStyle> ReadChartPointStyles(XElement series, PptxTheme theme, PptxColorMap colorMap, PptxSceneChartPlotKind plotKind, int seriesIndex, PptxSceneChartMarker seriesMarker)
    {
        var styles = new List<PptxSceneChartPointStyle>();
        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (!TryReadChartNonNegativeIndex(point, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            XElement? markerElement = point.Element(ChartNamespace + "marker");
            PptxSceneChartMarker? marker = markerElement is null ? null : ReadChartMarker(point, theme, colorMap, plotKind, seriesIndex);
            if (marker is not null)
            {
                marker = marker with
                {
                    Size = plotKind == PptxSceneChartPlotKind.Line && marker.SizeValue is null
                        ? marker.SymbolSpecified ? PptxChartMarkerMetricRules.UnsetMarkerSize : seriesMarker.Size : marker.Size,
                    SymbolKind = plotKind == PptxSceneChartPlotKind.Line && !marker.SymbolSpecified ? seriesMarker.SymbolKind : marker.SymbolKind,
                    Symbol = plotKind == PptxSceneChartPlotKind.Line && !marker.SymbolSpecified ? seriesMarker.Symbol : marker.Symbol,
                    Fill = ReadChartPointFill(markerElement!.Element(ChartNamespace + "spPr"), theme, colorMap)
                };
            }
            (double? explosion, string explosionValue) = ReadChartElementDoubleWithValue(point, "explosion");
            styles.Add(new PptxSceneChartPointStyle(
                index,
                ReadChartPointFill(shapeProperties, theme, colorMap),
                ReadChartPatternFill(shapeProperties, theme, colorMap),
                ReadChartLine(shapeProperties, theme, colorMap),
                ReadEffects(shapeProperties),
                explosion,
                explosionValue)
            {
                HasShapeProperties = shapeProperties is not null,
                Marker = marker
            });
        }

        return styles;
    }

    internal static bool TryReadChartNonNegativeIndex(XElement element, out int index)
    {
        return int.TryParse(ReadChartElementValue(element, "idx"), NumberStyles.Integer, CultureInfo.InvariantCulture, out index) &&
            index >= 0;
    }

    private static PptxSceneFillStyle ReadChartPointFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneFillStyle(true, default, 0d);
        }

        return PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxSceneLineStyle ReadChartLine(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, double? inheritedWidth = null)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return default(PptxSceneLineStyle) with { NoFill = true };
        }

        if (line is null) return default;
        bool widthSpecified = line.Attribute("w") is not null;
        return
            PptxLineStyleReader.TryReadLineWithAlpha(line, theme, colorMap, out RgbColor color, out double lineWidth, out double alpha, fallbackLineWidth: inheritedWidth)
                ? ReadSceneLineStyle(line, color, lineWidth, alpha, widthSpecified)
                : widthSpecified ? default(PptxSceneLineStyle) with { Width = lineWidth, WidthSpecified = true } : default;
    }

    private static string? ReadChartSeriesName(XElement series)
    {
        return ReadChartText(series.Element(ChartNamespace + "tx"), false);
    }

    private static PptxSceneChartSeriesFormulas ReadChartSeriesFormulas(XElement series)
    {
        return new PptxSceneChartSeriesFormulas(
            ReadChartFormula(series.Element(ChartNamespace + "tx"), "strRef", "numRef"),
            ReadChartFormula(series.Element(ChartNamespace + "val"), "numRef"),
            ReadChartFormula(series.Element(ChartNamespace + "cat"), "strRef", "numRef", "multiLvlStrRef"),
            ReadChartFormula(series.Element(ChartNamespace + "xVal"), "numRef", "strRef", "multiLvlStrRef"),
            ReadChartFormula(series.Element(ChartNamespace + "yVal"), "numRef"),
            ReadChartFormula(series.Element(ChartNamespace + "bubbleSize"), "numRef"));
    }

    private static string? ReadChartFormula(XElement? container, params string[] referenceKinds)
    {
        if (container is null)
        {
            return null;
        }

        foreach (string referenceKind in referenceKinds)
        {
            if (container.Element(ChartNamespace + referenceKind) is { } reference)
            {
                return (string?)reference.Element(ChartNamespace + "f");
            }
        }

        return null;
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

        string richText = string.Concat(EnumerateChartRichRuns(text?.Element(ChartNamespace + "rich"))
            .Select(source => (string?)source.Run.Element(DrawingNamespace + "t")));
        return string.IsNullOrWhiteSpace(richText) ? null : richText;
    }

    // Runs and cached fields share the authored line and paragraph order.
    private static IEnumerable<(XElement Run, int LineIndex)> EnumerateChartRichRuns(XElement? rich)
    {
        if (rich is null) yield break;

        int lineIndex = 0;
        foreach (XElement paragraph in rich.Elements(DrawingNamespace + "p"))
        {
            foreach (XElement element in paragraph.Elements())
            {
                if (element.Name == DrawingNamespace + "br")
                    lineIndex++;
                else if (element.Name == DrawingNamespace + "r" || element.Name == DrawingNamespace + "fld")
                    yield return (element, lineIndex);
            }
            lineIndex++;
        }
    }

    internal static IReadOnlyList<PptxSceneChartTextRun> ReadChartTextRuns(XElement? text, PptxTheme theme, PptxColorMap colorMap, bool inheritParagraphFormatting = false)
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
        foreach (var (run, lineIndex) in EnumerateChartRichRuns(rich))
        {
            string runText = run.Element(DrawingNamespace + "t")?.Value ?? string.Empty;
            if (runText.Length == 0)
            {
                continue;
            }

            XElement? runProperties = run.Element(DrawingNamespace + "rPr");
            // Titles inherit paragraph defaults. Data labels retain their native
            // point/container scope; direct run declarations remain authoritative.
            XElement? paragraphProperties = inheritParagraphFormatting
                ? run.Ancestors(DrawingNamespace + "p").FirstOrDefault()?
                    .Element(DrawingNamespace + "pPr")?.Element(DrawingNamespace + "defRPr")
                : null;
            PptxSceneChartTextStyleOverride style = MergeChartTextStyleOverride(
                ReadChartTextRunStyle(paragraphProperties, theme, colorMap),
                ReadChartTextRunStyle(runProperties, theme, colorMap));
            PptxSceneHyperlinkClick hyperlink = ReadHyperlinkClick(runProperties);
            runs.Add(new PptxSceneChartTextRun(runText, style, hyperlink.RelationshipId, hyperlink.Action, lineIndex));
        }

        return runs;
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
