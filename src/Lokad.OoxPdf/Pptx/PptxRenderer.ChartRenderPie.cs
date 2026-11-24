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
    private static IReadOnlyList<ChartIndexedPieSlice> BuildChartIndexedPieSlices(ChartIndexedNumberVector values, out int slotCount)
    {
        IReadOnlyDictionary<int, ChartIndexedNumberPoint> workbookPoints = values.WorkbookPointsForPlotVisibility(values.PlotVisibleOnly)
            .GroupBy(point => point.Index)
            .ToDictionary(group => group.Key, group => group.First());
        // RV20: sparse last-wins slot index instead of the dense array: same
        // slots (same resolution, range, validity and precedence) with
        // provenance kept once in the source. Charge the resolved slot count
        // exactly as the dense materialization did.
        slotCount = values.DensePointCount();
        if (slotCount <= 0)
        {
            return [];
        }

        return BuildDensePointIndex(values)
            .Values
            .Where(point => point.Value is > 0d)
            .OrderBy(point => point.Index)
            .Select(point => new ChartIndexedPieSlice(
                point.Index,
                (point.Value ?? 0d),
                point,
                workbookPoints.TryGetValue(point.Index, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null))
            .ToArray();
    }

    private static IEnumerable<(ChartIndexedPieSlice Slice, double Start, double Sweep, double Midpoint)> EnumeratePieSliceAngles(
        IReadOnlyList<ChartIndexedPieSlice> slices, double total, double firstSliceAngle)
    {
        // Advance clockwise from the authored start for every painted wedge.
        double start = Math.PI / 2d - firstSliceAngle * Math.PI / 180d;
        foreach (ChartIndexedPieSlice slice in slices)
        {
            double sweep = -slice.Value / total * Math.PI * 2d;
            yield return (slice, start, sweep, start + sweep / 2d);
            start += sweep;
        }
    }

    private static void RenderPieOrDoughnutSlices(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, double total, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, double firstSliceAngle)
    {
        if (total <= 0d)
        {
            return;
        }

        ChartPolarGeometry geometry = layout.Geometry;
        double innerRadius = geometry.Radius * Math.Clamp(holeSize, 0d, 0.95d);
        foreach (var (slice, start, sweep, midpoint) in EnumeratePieSliceAngles(slices, total, firstSliceAngle))
        {
            double explosionOffset = pointExplosions.TryGetValue(slice.Index, out double explosion) ? geometry.Radius * explosion : 0d;
            double sliceCenterX = geometry.CenterX + Math.Cos(midpoint) * explosionOffset;
            double sliceCenterY = geometry.CenterY + Math.Sin(midpoint) * explosionOffset;
            ChartSeriesFill fill = pointFills.TryGetValue(slice.Index, out ChartSeriesFill explicitFill)
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, slice.Index, slices.Count), 1d, null, null);
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            AppendPieOrDoughnutSlicePath(graphics, sliceCenterX, sliceCenterY, geometry.Radius, innerRadius, start, sweep);
            graphics.FillCurrentPathEvenOdd();
            if (fill.Alpha < 1d)
            {
                graphics.RestoreState();
            }

            if (pointStrokes.TryGetValue(slice.Index, out ChartSeriesStroke stroke))
            {
                if (stroke.Alpha < 1d)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(1d, stroke.Alpha);
                }

                SetChartStroke(graphics, stroke);
                AppendPieOrDoughnutSlicePath(graphics, sliceCenterX, sliceCenterY, geometry.Radius, innerRadius, start, sweep);
                graphics.StrokeCurrentPath();
                if (stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }
        }
    }

    private static void AppendPieOrDoughnutSlicePath(PdfGraphicsBuilder graphics, double centerX, double centerY, double outerRadius, double innerRadius, double startAngle, double sweepAngle)
    {
        graphics.MoveTo(centerX + Math.Cos(startAngle) * outerRadius, centerY + Math.Sin(startAngle) * outerRadius);
        AppendCircularArc(graphics, centerX, centerY, outerRadius, startAngle, sweepAngle);
        if (innerRadius <= 0d)
        {
            graphics.LineTo(centerX, centerY);
        }
        else
        {
            double endAngle = startAngle + sweepAngle;
            graphics.LineTo(centerX + Math.Cos(endAngle) * innerRadius, centerY + Math.Sin(endAngle) * innerRadius);
            AppendCircularArc(graphics, centerX, centerY, innerRadius, endAngle, -sweepAngle);
        }
        graphics.ClosePath();
    }

    private static void AppendCircularArc(PdfGraphicsBuilder graphics, double centerX, double centerY, double radius, double startAngle, double sweepAngle)
    {
        int segmentCount = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweepAngle) / (Math.PI / 2d)));
        double segmentSweep = sweepAngle / segmentCount;
        for (int segment = 0; segment < segmentCount; segment++)
        {
            double segmentStart = startAngle + segmentSweep * segment;
            AppendEllipseArcSegment(
                graphics,
                centerX,
                centerY,
                radius,
                radius,
                // AppendEllipseArcSegment takes visual angles (y-down): negate the math
                // angles (y-up) so the emitted arc matches the MoveTo/LineTo endpoints.
                -segmentStart * 180d / Math.PI,
                -segmentSweep * 180d / Math.PI,
                moveToStart: false);
        }
    }

    private static void EmitChartDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, string id, OoxPdfSeverity severity, string message, string? partName, int slideIndex, string fallback)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            id,
            severity,
            message,
            partName,
            PageIndex: null,
            SlideIndex: slideIndex,
            Feature: "chart",
            Fallback: fallback));
    }
}
