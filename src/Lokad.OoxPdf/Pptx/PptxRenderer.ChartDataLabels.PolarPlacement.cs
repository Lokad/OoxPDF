using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // One existing label owns its furniture, text, clip and links through placement.
    private sealed record PolarLabelPaint(ChartLayoutBox Box, ChartShapeStyle Shape,
        ChartLayoutBox? KeyBox, ChartSeriesFill? KeyFill, bool Authored,
        int RunStart, int RunCount, int LinkStart, int LinkCount, Action? Leader);

    private static async ValueTask<double[]> PlacePolarLabelRowsAsync(
        IReadOnlyList<PolarLabelPaint> paints, IReadOnlyList<TextRun> runs,
        ChartPlotBox placementBox, PresentationFontResolver? resolver, PptxRenderContext context)
    {
        var offsets = new double[paints.Count];
        if (paints.Count < 2) return offsets;
        var bounds = new ChartLayoutBox?[paints.Count];
        var measurer = new ChartTextMeasurer(resolver);
        for (int index = 0; index < paints.Count; index++)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            PolarLabelPaint paint = paints[index];
            if (!paint.Shape.IsEmpty) Include(ref bounds[index], paint.Box);
            if (paint.KeyBox is { } key) Include(ref bounds[index], key);
            for (int part = paint.RunStart; part < paint.RunStart + paint.RunCount; part++)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                TextRun run = runs[part];
                if (run.Alpha <= 0d || string.IsNullOrWhiteSpace(run.Text)) continue;
                var style = new ChartTextStyle(run.FontFamily, run.FontSize, run.CharacterSpacing,
                    run.Color, run.Alpha, run.Bold, run.Italic, run.Underline, run.Strike, null, null);
                double width = await measurer.MeasureAsync(run.Text, style).ConfigureAwait(false);
                var emission = style with { FontSize = OfficePdfTextEmissionProfile.FontSize(style.FontSize) };
                if (await measurer.TryMeasureInkBoundsAsync(run.Text, emission).ConfigureAwait(false) is not { } ink)
                    return offsets; // Existing font diagnostics own unmeasurable paint; retain its placement.
                double free = Math.Max(0d, run.Width - width);
                double x = run.X + (run.Alignment == TextAlignment.Center ? free / 2d : run.Alignment == TextAlignment.Right ? free : 0d);
                double left = Math.Max(x, run.ClipX), right = Math.Min(x + width, run.ClipX + run.ClipWidth);
                double bottom = Math.Max(run.Y + run.BaselineOffset + ink.Bottom, run.ClipY);
                double top = Math.Min(run.Y + run.BaselineOffset + ink.Top, run.ClipY + run.ClipHeight);
                if (right > left && top > bottom) Include(ref bounds[index], new(left, bottom, right - left, top - bottom));
            }
        }

        // Source-Y order is stable. A current row contains disjoint X intervals,
        // so its predecessor/successor suffice; there is no quadratic pair scan.
        var row = new SortedSet<(double X, int Index)>();
        double rowTop = double.PositiveInfinity, rowBottom = double.PositiveInfinity;
        foreach (int index in Enumerable.Range(0, paints.Count).Where(i => bounds[i].HasValue)
            .OrderByDescending(i => bounds[i]!.Value.Y + bounds[i]!.Value.Height).ThenBy(i => i))
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            ChartLayoutBox box = bounds[index]!.Value;
            double top = box.Y + box.Height;
            if (row.Count == 0 || top <= rowBottom)
            {
                row.Clear();
                rowTop = top;
                rowBottom = box.Y;
            }
            (double X, int Index) previous = row.GetViewBetween((double.NegativeInfinity, int.MinValue),
                (box.X, int.MaxValue)).Reverse().FirstOrDefault((double.NaN, -1));
            (double X, int Index) next = row.GetViewBetween((box.X, int.MinValue),
                (double.PositiveInfinity, int.MaxValue)).FirstOrDefault((double.NaN, -1));
            bool conflict = previous.Index >= 0 && bounds[previous.Index]!.Value.X + bounds[previous.Index]!.Value.Width > box.X ||
                next.Index >= 0 && next.X < box.X + box.Width;
            if (conflict)
            {
                rowTop = Math.Min(top, rowBottom - PptxChartMetricRules.DataLabelVerticalPadding);
                rowBottom = rowTop - box.Height;
                row.Clear();
            }
            double targetTop = Math.Min(top, rowTop), dy = targetTop - top;
            if (dy != 0d && (paints[index].Authored || box.Y + dy < placementBox.Y || box.Width > placementBox.Width))
            {
                Array.Clear(offsets);
                context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic("PPTX_CHART_LABEL_OVERLAP", OoxPdfSeverity.Warning,
                    "Automatic polar labels do not fit separate rows without leaving their container or moving authored geometry.",
                    PartName: context.SlidePartName, SlideIndex: context.SlideNumber,
                    Feature: "Polar data labels", Fallback: "Source label arrangement retained"));
                return offsets;
            }
            offsets[index] = dy;
            rowBottom = Math.Min(rowBottom, box.Y + dy);
            row.Add((box.X, index));
        }
        return offsets;

        static void Include(ref ChartLayoutBox? current, ChartLayoutBox value)
        {
            if (current is not { } box) { current = value; return; }
            double x = Math.Min(box.X, value.X), y = Math.Min(box.Y, value.Y);
            current = new(x, y, Math.Max(box.X + box.Width, value.X + value.Width) - x,
                Math.Max(box.Y + box.Height, value.Y + value.Height) - y);
        }
    }
}
