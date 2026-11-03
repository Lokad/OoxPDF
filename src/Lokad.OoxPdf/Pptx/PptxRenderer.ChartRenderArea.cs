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
        ChartColorPalette chartPalette,
        ChartLayoutBox plotAreaBox,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        ChartAreaPlotOptions areaOptions,
        bool categoryPointsOnTickMarks,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        IReadOnlyList<ChartSeriesStroke?> seriesStrokes,
        ChartValueAxisRenderOptions valueAxisOptions,
        ChartAxesStyle axesStyle,
        ChartShapeStyle plotAreaStyle,
        ChartValueExtents valueExtents,
        bool categoryAxisReversed = false, PptxRenderContext? context = null)
    {
        bool majorGridlines = valueAxisOptions.MajorGridlines;
        bool minorGridlines = valueAxisOptions.MinorGridlines;
        ChartGridlineStyle gridlineStyle = valueAxisOptions.GridlineStyle;
        ChartAxisUnits axisUnits = valueAxisOptions.Units;
        double? valueAxisCrossingValue = valueAxisOptions.CrossingValue;
        bool valueAxisReversed = valueAxisOptions.Reversed;
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotY, plotHeight, valueAxisReversed);
        void DrawGridlines(bool major, ChartSeriesStroke? stroke)
        {
            // Source intervals share the line grid's major/minor lattice and
            // subdivision ownership. Other area scales retain their fallback.
            if (valueAxisOptions.TickPolicy == ChartAxisTickPolicy.SourceMinimumIntervals)
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents,
                    axisUnits, valueAxisCrossingValue, valueAxisReversed, major, stroke,
                    valueAxisOptions.HasMajorGridlineElement, valueAxisOptions.TickPolicy, context: context);
            else
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents,
                    major ? axisUnits.MajorUnit : axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major, stroke, context: context);
        }
        IReadOnlyList<IReadOnlyList<double?>> denseSeries = DensifyChartValueSeries(series);
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        {
            if (minorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawGridlines(major: false, gridlineStyle.Minor));
            }

            if (majorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawGridlines(major: true, gridlineStyle.Major));
            }

            RenderAreaChartSeries(graphics, theme, colorMap, chartPalette, plotBox, denseSeries, areaOptions, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, categoryPointsOnTickMarks, valueAxisCrossingY, categoryAxisReversed);

            ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
            ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
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

    private static void RenderAreaChartSeries(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<double?>> series, ChartAreaPlotOptions areaOptions, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents valueExtents, bool valueAxisReversed, bool categoryPointsOnTickMarks, double crossingY, bool categoryAxisReversed)
    {
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double[] lower = new double[pointCount];
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double?> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            // Office stacks source slots through missing values, even when Gap
            // is explicitly requested. Standard areas retain separate segments.
            if (!areaOptions.Stacked && areaOptions.DisplayBlanksAs == PptxSceneChartDisplayBlanksAs.Gap)
            {
                int start = 0;
                while (start < pointCount)
                {
                    while (start < pointCount && (start >= values.Count || values[start] is null))
                    {
                        start++;
                    }

                    int end = start;
                    while (end < pointCount && end < values.Count && values[end] is not null)
                    {
                        end++;
                    }

                    if (end - start >= 2)
                    {
                        RenderAreaChartSeriesSegment(graphics, theme, colorMap, chartPalette, plotBox, values, start, end, lower, areaOptions, seriesIndex, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, series, categoryPointsOnTickMarks, crossingY, categoryAxisReversed);
                    }

                    start = Math.Max(end, start + 1);
                }

            }
            else
            {
                RenderAreaChartSeriesSegment(graphics, theme, colorMap, chartPalette, plotBox, values, 0, pointCount, lower, areaOptions, seriesIndex, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, series, categoryPointsOnTickMarks, crossingY, categoryAxisReversed);
            }
            // Stacking belongs to source slots, including isolated values that
            // do not form a painted polygon on their own.
            if (areaOptions.Stacked)
            {
                for (int i = 0; i < values.Count; i++)
                    lower[i] += NormalizeStackedValue(values[i] ?? 0d, GetCategoryPositiveTotal(series, i, areaOptions.PercentStacked), areaOptions.PercentStacked);
            }
        }
    }

    private static void RenderAreaChartSeriesSegment(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, ChartPlotBox plotBox, IReadOnlyList<double?> values, int startIndex, int endIndex, double[] lower, ChartAreaPlotOptions areaOptions, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents valueExtents, bool valueAxisReversed, IReadOnlyList<IReadOnlyList<double?>> allSeries, bool categoryPointsOnTickMarks, double crossingY, bool categoryAxisReversed)
    {
        double plotX = plotBox.X;
        double plotWidth = plotBox.Width;
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
            double pointX = ChartCategoryToPlotCoordinate(i, pointCount, plotX, plotWidth, categoryPointsOnTickMarks, categoryAxisReversed);
            double value = i < values.Count && values[i] is { } indexedValue ? indexedValue : 0d;
            var span = ResolveAreaPointVerticalSpan(plotBox, valueExtents, areaOptions,
                valueAxisReversed, seriesIndex, value, lower[i],
                GetCategoryPositiveTotal(allSeries, i, areaOptions.PercentStacked), crossingY);
            int segmentIndex = i - startIndex;
            upperPoints[segmentIndex] = (pointX, span.UpperY);
            lowerPoints[segmentIndex] = (pointX, span.LowerY);
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

        // Unstyled area series carry no outline (Office draws fills only); explicit
        // strokes keep the inherited-width behavior below.
        bool defaultAreaOutline = seriesIndex >= seriesStrokes.Count || seriesStrokes[seriesIndex] is null;
        if (!defaultAreaOutline)
        {
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

    // Painting and data labels use the same physical span. Source accumulation
    // remains separate from whether a polygon or a label is visible.
    private static (double LowerY, double UpperY, double UpperValue) ResolveAreaPointVerticalSpan(
        ChartPlotBox plotBox, ChartValueExtents extents, ChartAreaPlotOptions options,
        bool reversed, int seriesIndex, double value, double lower, double positiveTotal, double crossingY)
    {
        double upper = options.Stacked ? lower + NormalizeStackedValue(value, positiveTotal, options.PercentStacked) : value;
        double lowerY = !options.Stacked || seriesIndex == 0
            ? crossingY : ChartValueToPlotCoordinate(extents, lower, plotBox.Y, plotBox.Height, reversed);
        return (lowerY, ChartValueToPlotCoordinate(extents, upper, plotBox.Y, plotBox.Height, reversed), upper);
    }
}
