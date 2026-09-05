using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly record struct ChartBarPlotOptions(
        PptxSceneChartGrouping Grouping,
        PptxSceneChartBarDirection BarDirection,
        ChartBooleanOption VaryColors,
        double GapWidth,
        double Overlap);

    private static ChartBarPlotOptions ReadSceneOrXmlChartBarOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxSceneChartGrouping defaultGrouping)
    {
        ChartBooleanOption ReadSceneOrXmlChartVaryColors()
        {
            if (plot is not null)
            {
                return new ChartBooleanOption(plot.VaryColors ?? true, plot.VaryColorsValue, plot.VaryColors is not null);
            }
    
            (bool? varyColors, string varyColorsValue) = PptxSceneBuilder.ReadChartPlotVaryColors(chartElement);
            return new ChartBooleanOption(varyColors ?? true, varyColorsValue, varyColors is not null);
        }

        double ReadSceneOrXmlChartGapWidth()
        {
            double ReadXmlChartGapWidth(XElement chartElement)
            {
                (double? gapWidth, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(chartElement, "gapWidth");
                return gapWidth is { } rawGapWidth
                    ? Math.Clamp(rawGapWidth, 0d, 500d)
                    : 150d;
            }
    
            return plot is not null
                ? plot.GapWidth ?? 150d
                : ReadXmlChartGapWidth(chartElement);
        }

        double ReadSceneOrXmlChartOverlap()
        {
            double ReadXmlChartOverlap(XElement chartElement)
            {
                (double? overlap, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(chartElement, "overlap");
                return overlap is { } rawOverlap
                    ? Math.Clamp(rawOverlap, -100d, 100d)
                    : 0d;
            }
    
            return plot is not null
                ? plot.Overlap ?? 0d
                : ReadXmlChartOverlap(chartElement);
        }

        return new ChartBarPlotOptions(
            ReadSceneOrXmlChartGrouping(plot, chartElement, defaultGrouping),
            ReadSceneOrXmlChartBarDirection(plot, chartElement),
            ReadSceneOrXmlChartVaryColors(),
            ReadSceneOrXmlChartGapWidth(),
            ReadSceneOrXmlChartOverlap());
    }

    private readonly record struct ChartLinePlotOptions(
        PptxSceneChartGrouping Grouping,
        bool Stacked,
        bool PercentStacked,
        IReadOnlyList<ChartBooleanOption> SmoothSeries,
        PptxSceneChartDisplayBlanksAs DisplayBlanksAs);

    private static ChartLinePlotOptions ReadSceneOrXmlChartLineOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, XElement chartElement, PptxSceneChartGrouping defaultGrouping)
    {
        PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(plot, chartElement, defaultGrouping);
        return new ChartLinePlotOptions(
            grouping,
            IsStackedChartGrouping(grouping),
            IsPercentStackedChartGrouping(grouping),
            ReadSceneOrXmlSmoothSeries(plot, chartElement),
            ReadSceneOrXmlChartDisplayBlanksAs(sceneChart, chartXml));
    }

    private readonly record struct ChartAreaPlotOptions(
        PptxSceneChartGrouping Grouping,
        bool Stacked,
        bool PercentStacked,
        PptxSceneChartDisplayBlanksAs DisplayBlanksAs);

    private static ChartAreaPlotOptions ReadSceneOrXmlChartAreaOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, XElement chartElement, PptxSceneChartGrouping defaultGrouping)
    {
        PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(plot, chartElement, defaultGrouping);
        return new ChartAreaPlotOptions(
            grouping,
            IsStackedChartGrouping(grouping),
            IsPercentStackedChartGrouping(grouping),
            ReadSceneOrXmlChartDisplayBlanksAs(sceneChart, chartXml));
    }

    private readonly record struct ChartScatterPlotOptions(
        PptxSceneChartScatterStyle ScatterStyle,
        bool ConnectLines,
        IReadOnlyList<ChartBooleanOption> SmoothSeries);

    private static ChartScatterPlotOptions ReadSceneOrXmlChartScatterOptions(PptxSceneChartPlot? plot, XElement chartElement)
    {
        PptxSceneChartScatterStyle scatterStyle = ReadSceneOrXmlChartScatterStyle(plot, chartElement);
        return new ChartScatterPlotOptions(
            scatterStyle,
            ResolveChartScatterLineConnection(scatterStyle),
            ReadSceneOrXmlSmoothSeries(plot, chartElement));
    }

    private readonly record struct ChartRadarPlotOptions(PptxSceneChartRadarStyle RadarStyle);

    private static ChartRadarPlotOptions ReadSceneOrXmlChartRadarOptions(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return new ChartRadarPlotOptions(ReadSceneOrXmlChartRadarStyle(plot, chartElement));
    }

    private static double NormalizeAngleDegrees(double angle)
    {
        double normalized = angle % 360d;
        return normalized < 0d ? normalized + 360d : normalized;
    }

    private readonly record struct ChartPolarPointOptions(
        IReadOnlyDictionary<int, ChartSeriesFill> PointFills,
        IReadOnlyDictionary<int, ChartSeriesStroke> PointStrokes,
        IReadOnlyDictionary<int, double> PointExplosions,
        double FirstSliceAngle);

    private static IReadOnlyList<PptxSceneChartSeries> ReadXmlChartSeries(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        PptxSceneChartPlotKind plotKind = chartElement.Name == ChartNamespace + "ser"
            ? PptxSceneChartPlotKind.Unknown
            : PptxSceneBuilder.ParseChartPlotKind(chartElement.Name.LocalName);
        bool markersEnabled = chartElement.Name != ChartNamespace + "ser" &&
            PptxSceneBuilder.IsOoxmlBooleanElementEnabled(chartElement.Element(ChartNamespace + "marker"));
        XElement plotElement = chartElement.Name == ChartNamespace + "ser"
            ? new XElement(ChartNamespace + "unknownChart", chartElement)
            : chartElement;
        return PptxSceneBuilder.ReadChartSeries(plotElement, theme, colorMap, plotKind, markersEnabled);
    }

    private static ChartPolarPointOptions ReadSceneOrXmlChartPolarPointOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap, ChartWorkbookData? workbook)
    {
        double ReadSceneOrXmlFirstSliceAngle()
        {
            if (plot is not null)
            {
                return plot.FirstSliceAngle is { } sceneAngle
                    ? NormalizeAngleDegrees(sceneAngle)
                    : 0d;
            }
    
            (double? firstSliceAngle, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(chartElement, "firstSliceAng");
            return firstSliceAngle is { } rawFirstSliceAngle
                ? NormalizeAngleDegrees(rawFirstSliceAngle)
                : 0d;
        }

        IReadOnlyDictionary<int, ChartSeriesFill> ReadSceneOrXmlChartPointFills()
        {
            if (plot is not null)
            {
                return plot.Series.Count > 0 ? ReadSceneChartPointFills(plot.Series[0]) : new Dictionary<int, ChartSeriesFill>();
            }
    
            IReadOnlyList<PptxSceneChartSeries> series = ReadXmlChartSeries(chartElement, theme, colorMap);
            return series.Count == 0 ? new Dictionary<int, ChartSeriesFill>() : ReadSceneChartPointFills(series[0]);
        }

        IReadOnlyDictionary<int, ChartSeriesStroke> ReadSceneOrXmlChartPointStrokes()
        {
            if (plot is not null)
            {
                return plot.Series.Count > 0 ? ReadSceneChartPointStrokes(plot.Series[0]) : new Dictionary<int, ChartSeriesStroke>();
            }
    
            IReadOnlyList<PptxSceneChartSeries> series = ReadXmlChartSeries(chartElement, theme, colorMap);
            return series.Count == 0 ? new Dictionary<int, ChartSeriesStroke>() : ReadSceneChartPointStrokes(series[0]);
        }

        return new ChartPolarPointOptions(
            ReadSceneOrXmlChartPointFills(),
            ReadSceneOrXmlChartPointStrokes(),
            ReadSceneOrXmlChartPointExplosions(plot, chartElement, workbook),
            ReadSceneOrXmlFirstSliceAngle());
    }

    private readonly record struct ChartDoughnutPlotOptions(
        ChartPolarPointOptions PolarPoints,
        double HoleSize);

    private static ChartDoughnutPlotOptions ReadSceneOrXmlChartDoughnutOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap, ChartWorkbookData? workbook)
    {
        double ReadSceneDoughnutHoleSize(PptxSceneChartPlot? plot, XElement doughnutChart)
        {
            if (plot is not null)
            {
                return plot.HoleSize is { } rawHoleSize
                    ? Math.Clamp(rawHoleSize / 100d, PptxChartMetricRules.DoughnutHoleMinimumRatio, PptxChartMetricRules.DoughnutHoleMaximumRatio)
                    : PptxChartMetricRules.DoughnutHoleFallbackRatio;
            }
    
            (double? holeSize, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(doughnutChart, "holeSize");
            return holeSize is { } xmlHoleSize
                ? Math.Clamp(xmlHoleSize / 100d, PptxChartMetricRules.DoughnutHoleMinimumRatio, PptxChartMetricRules.DoughnutHoleMaximumRatio)
                : PptxChartMetricRules.DoughnutHoleFallbackRatio;
        }

        return new ChartDoughnutPlotOptions(
            ReadSceneOrXmlChartPolarPointOptions(plot, chartElement, theme, colorMap, workbook),
            ReadSceneDoughnutHoleSize(plot, chartElement));
    }
}
