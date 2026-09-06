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
    private static void RenderAreaChart(
        PdfGraphicsBuilder graphics,
        PptxTheme theme,
        PptxColorMap colorMap,
        IReadOnlyList<RgbColor>? chartPalette,
        ChartLayoutBox plotAreaBox,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        ChartAreaPlotOptions areaOptions,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        IReadOnlyList<ChartSeriesStroke?> seriesStrokes,
        ChartValueAxisRenderOptions valueAxisOptions,
        ChartAxesStyle axesStyle,
        ChartShapeStyle plotAreaStyle,
        ChartValueExtents valueExtents)
    {
        bool stacked = areaOptions.Stacked;
        bool percentStacked = areaOptions.PercentStacked;
        bool majorGridlines = valueAxisOptions.MajorGridlines;
        bool minorGridlines = valueAxisOptions.MinorGridlines;
        ChartGridlineStyle gridlineStyle = valueAxisOptions.GridlineStyle;
        ChartAxisUnits axisUnits = valueAxisOptions.Units;
        double? valueAxisCrossingValue = valueAxisOptions.CrossingValue;
        bool valueAxisReversed = valueAxisOptions.Reversed;
        PptxSceneChartDisplayBlanksAs displayBlanksAs = areaOptions.DisplayBlanksAs;
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        {
            if (minorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor));
            }

            if (majorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major));
            }

            RenderAreaChartSeries(graphics, theme, colorMap, chartPalette, plotBox, denseSeries, stacked, percentStacked, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, displayBlanksAs);

            ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
            ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
            double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotY, plotHeight, valueAxisReversed);
            if (axesStyle.CategoryAxisVisible && categoryAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, categoryAxisStroke);
                RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(plotX, valueAxisCrossingY, plotX + plotWidth, valueAxisCrossingY));
            }

            if (axesStyle.ValueAxisVisible)
            {
                if (valueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, valueAxisStroke);
                    double axisX = axesStyle.ValueAxisRightSide ? plotX + plotWidth : plotX;
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }

                if (axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke && secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }
            }
        }
    }

    private static void RenderAreaChartSeries(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, bool stacked, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents valueExtents, bool valueAxisReversed, PptxSceneChartDisplayBlanksAs displayBlanksAs)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double[] lower = new double[pointCount];
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            if (displayBlanksAs == PptxSceneChartDisplayBlanksAs.Gap)
            {
                int start = 0;
                while (start < pointCount)
                {
                    while (start < pointCount && (start >= values.Count || values[start]?.Value is null))
                    {
                        start++;
                    }

                    int end = start;
                    while (end < pointCount && end < values.Count && values[end]?.Value is not null)
                    {
                        end++;
                    }

                    if (end - start >= 2)
                    {
                        RenderAreaChartSeriesSegment(graphics, theme, colorMap, chartPalette, plotBox, values, start, end, lower, stacked, percentStacked, seriesIndex, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, series);
                    }

                    start = Math.Max(end, start + 1);
                }

                continue;
            }

            RenderAreaChartSeriesSegment(graphics, theme, colorMap, chartPalette, plotBox, values, 0, pointCount, lower, stacked, percentStacked, seriesIndex, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, series);
        }
    }

    private static void RenderAreaChartSeriesSegment(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberPoint?> values, int startIndex, int endIndex, double[] lower, bool stacked, bool percentStacked, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents valueExtents, bool valueAxisReversed, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> allSeries)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        int pointCount = lower.Length;
        int segmentPointCount = Math.Max(0, endIndex - startIndex);
        if (segmentPointCount < 2)
        {
            return;
        }

        var upperPoints = new (double X, double Y)[segmentPointCount];
        var lowerPoints = new (double X, double Y)[segmentPointCount];
        for (int i = startIndex; i < endIndex; i++)
        {
            double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
            double value = i < values.Count && values[i]?.Value is { } indexedValue ? indexedValue : 0d;
            double lowerValue = stacked ? lower[i] : 0d;
            double positiveTotal = GetCategoryPositiveTotal(allSeries, i, percentStacked);
            double normalizedValue = NormalizeStackedValue(value, positiveTotal, percentStacked);
            double upperValue = stacked ? lower[i] + normalizedValue : value;
            int segmentIndex = i - startIndex;
            upperPoints[segmentIndex] = (pointX, ChartValueToPlotCoordinate(valueExtents, upperValue, plotY, plotHeight, valueAxisReversed));
            lowerPoints[segmentIndex] = (pointX, ChartValueToPlotCoordinate(valueExtents, lowerValue, plotY, plotHeight, valueAxisReversed));
            if (stacked)
            {
                lower[i] = upperValue;
            }
        }

        var polygon = new (double X, double Y)[segmentPointCount * 2];
        for (int i = 0; i < segmentPointCount; i++)
        {
            polygon[i] = upperPoints[i];
            polygon[polygon.Length - i - 1] = lowerPoints[i];
        }

        ChartSeriesFill fill = ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, 1d);
        if (fill.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
        }

        graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
        RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.FillPolygon(polygon));
        if (fill.Alpha < 1d)
        {
            graphics.RestoreState();
        }

        ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, ChartLineDefaultStrokeWidth);
        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokePolygon(polygon));

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }
}
