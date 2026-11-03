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
    private static (ChartValueExtents DataX, ChartValueExtents DataY, ChartValueExtents PathX, ChartValueExtents PathY)
        MeasureScatterChartExtents(IReadOnlyList<ScatterSeries> series, ChartScatterPlotOptions? connections,
            IReadOnlyList<bool>? hidden, CancellationToken cancellationToken)
    {
        ScatterPoint first = series.SelectMany(item => item.Points).First();
        double minX = first.X, maxX = first.X, minY = first.Y, maxY = first.Y;
        foreach (ScatterSeries item in series)
        foreach (ScatterPoint point in item.Points)
        {
            cancellationToken.ThrowIfCancellationRequested();
            minX = Math.Min(minX, point.X); maxX = Math.Max(maxX, point.X);
            minY = Math.Min(minY, point.Y); maxY = Math.Max(maxY, point.Y);
        }
        ChartValueExtents x = new(minX, maxX), y = new(minY, maxY);
        return (x, y, CurveExtents(x, horizontal: true), CurveExtents(y, horizontal: false));

        ChartValueExtents CurveExtents(ChartValueExtents data, bool horizontal)
        {
            double span = data.Max - data.Min;
            if (connections is not { ConnectLines: true } options || !(span > 0d) || !double.IsFinite(span)) return data;
            // Normalize before solving derivatives, so large source coordinates
            // do not overflow polynomial coefficients. No per-point buffer.
            double lower = 0d, upper = 1d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                if (!IsSmoothSeries(seriesIndex, options.SmoothSeries) || IsLineHiddenSeries(seriesIndex, hidden)) continue;
                IReadOnlyList<ScatterPoint> points = series[seriesIndex].Points;
                for (int i = 0; i + 1 < points.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    double Value(int index) => ((horizontal ? points[index].X : points[index].Y) - data.Min) / span;
                    double start = Value(i), end = Value(i + 1);
                    (double c1, double c2) = SmoothChartControls(Value(Math.Max(0, i - 1)), start, end, Value(Math.Min(points.Count - 1, i + 2)));
                    ExtendCubicChartExtents(start, c1, c2, end, ref lower, ref upper);
                }
            }
            double minimum = data.Min + lower * span, maximum = data.Min + upper * span;
            return double.IsFinite(minimum) && double.IsFinite(maximum) ? new(minimum, maximum) : data;
        }
    }

    private static void ExtendCubicChartExtents(double start, double c1, double c2, double end, ref double minimum, ref double maximum)
    {
        double a = -start + 3d * c1 - 3d * c2 + end;
        double b = 2d * (start - 2d * c1 + c2);
        double c = c1 - start;
        if (a == 0d)
        {
            if (b != 0d) Extend(-c / b, ref minimum, ref maximum);
            return;
        }
        double discriminant = b * b - 4d * a * c;
        if (discriminant < 0d) return;
        double q = -.5d * (b + Math.CopySign(Math.Sqrt(discriminant), b));
        Extend(q / a, ref minimum, ref maximum);
        if (q != 0d) Extend(c / q, ref minimum, ref maximum);

        void Extend(double t, ref double min, ref double max)
        {
            if (!(t > 0d && t < 1d)) return;
            double u = 1d - t;
            double value = u * u * u * start + 3d * u * u * t * c1 + 3d * u * t * t * c2 + t * t * t * end;
            min = Math.Min(min, value); max = Math.Max(max, value);
        }
    }

    private static bool IsLineHiddenSeries(int index, IReadOnlyList<bool>? hidden) => hidden is not null && index < hidden.Count && hidden[index];

    private static void RenderScatterChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, ChartPlotBox plotBox, IReadOnlyList<ScatterSeries> series, bool connectLines, bool bubble, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<ChartBooleanOption> smoothSeries, IReadOnlyList<bool> seriesLineHidden, ChartValueExtents xExtents, ChartValueExtents yExtents, ChartValueAxisRenderOptions xAxisOptions, ChartValueAxisRenderOptions yAxisOptions)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        double maxBubbleSize = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.Size));

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        double xAxisY = ChartValueToPlotCoordinate(yExtents, yAxisOptions.CrossingValue, plotY, plotHeight, yAxisOptions.Reversed);
        double yAxisX = ChartValueToPlotCoordinate(xExtents, xAxisOptions.CrossingValue, plotX, plotWidth, xAxisOptions.Reversed);
        graphics.StrokeLine(plotX, xAxisY, plotX + plotWidth, xAxisY);
        graphics.StrokeLine(yAxisX, plotY, yAxisX, plotY + plotHeight);

        {
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                ChartSeriesFill fill = ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, 1d);
                // Unstyled scatter connects share the line-chart 2.25pt default width
                // (Office smooth connects render 2.25); marker
                // fills and outlines keep their own sizes through the marker path.
                ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, ChartLineDefaultStrokeWidth);
                if (fill.Alpha < 1d || stroke.Alpha < 1d)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(fill.Alpha, stroke.Alpha);
                }

                // Unstyled series strokes default to round caps/joins; explicitly styled
                // lines keep DrawingML attr defaults (combo contract).
                bool defaultSeriesStroke = seriesIndex >= seriesStrokes.Count || seriesStrokes[seriesIndex] is null;
                ChartSeriesStroke scatterStroke = defaultSeriesStroke ? stroke with { Cap = stroke.Cap ?? 1, Join = stroke.Join ?? 1, Color = ApplyUnstyledLineStrokeTint(stroke.Color) } : stroke;
                SetChartStroke(graphics, scatterStroke);
                graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
                var points = new List<(double X, double Y)>(series[seriesIndex].Points.Count);
                foreach (ScatterPoint point in series[seriesIndex].Points)
                {
                    (double pointX, double pointY, _) = ResolveScatterPointGeometry(plotBox, point, bubble, xExtents, yExtents, maxBubbleSize, xAxisOptions.Reversed, yAxisOptions.Reversed);
                    points.Add((pointX, pointY));
                }

                if (connectLines && !IsLineHiddenSeries(seriesIndex, seriesLineHidden))
                {
                    StrokeScatterChartPathInPlotClip(graphics, plotBox, points, IsSmoothSeries(seriesIndex, smoothSeries));
                }

                foreach (ScatterPoint point in series[seriesIndex].Points)
                {
                    (double pointX, double pointY, double radius) = ResolveScatterPointGeometry(plotBox, point, bubble, xExtents, yExtents, maxBubbleSize, xAxisOptions.Reversed, yAxisOptions.Reversed);
                    if (bubble)
                    {
                        FillBubbleInPlotClip(pointX, pointY, radius);
                    }
                    else
                    {
                        DrawChartMarkerInPlotClip(graphics, plotBox, pointX, pointY, ChartMarker(seriesIndex, markerStyles), fill.Color, defaultSeriesStroke ? ApplyUnstyledLineStrokeTint(stroke.Color) : stroke.Color);
                    }

                }

                if (fill.Alpha < 1d || stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }
        }

        void FillBubbleInPlotClip(double pointX, double pointY, double radius)
        {
            RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.FillEllipseEvenOdd(pointX - radius, pointY - radius, radius * 2d, radius * 2d));
        }

    }

    private static void StrokeScatterChartPathInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<(double X, double Y)> points, bool smooth)
    {
        RenderInChartPlotAreaClip(graphics, plotBox, () =>
        {
            if (smooth)
            {
                StrokeSmoothChartPath(graphics, points);
            }
            else
            {
                StrokeStraightChartPath(graphics, points);
            }
        });
    }

}
