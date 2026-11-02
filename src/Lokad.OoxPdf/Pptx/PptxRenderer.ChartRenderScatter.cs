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
    private static ChartValueExtents GetScatterXValueExtents(IReadOnlyList<ScatterSeries> series)
    {
        double maxX = series.SelectMany(item => item.Points).Max(point => point.X);
        double minX = series.SelectMany(item => item.Points).Min(point => point.X);
        return new ChartValueExtents(minX, maxX);
    }

    private static ChartValueExtents GetScatterYValueExtents(IReadOnlyList<ScatterSeries> series)
    {
        double maxY = series.SelectMany(item => item.Points).Max(point => point.Y);
        double minY = series.SelectMany(item => item.Points).Min(point => point.Y);
        return new ChartValueExtents(minY, maxY);
    }

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
                // (Office smooth connects render 2.25, geometry-identical to ours); marker
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

        static bool IsLineHiddenSeries(int seriesIndex, IReadOnlyList<bool> seriesLineHidden)
        {
            return seriesIndex < seriesLineHidden.Count && seriesLineHidden[seriesIndex];
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
