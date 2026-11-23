using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private sealed record DefaultAxisTitleBand(ChartLayoutBox Box, ChartTextStyle Style,
        IReadOnlyList<MeasuredChartTextLine> Lines, double Height);

    private sealed record DefaultAxisTitleBands(ChartFrameBox Body, bool FitBody,
        IReadOnlyDictionary<int, DefaultAxisTitleBand> Titles);

    private static async ValueTask<ChartLayout> PrepareDefaultAxisTitleBandsAsync(
        ChartLayout layout, PptxTheme theme, PptxSceneChart chart, PresentationFontResolver fontResolver)
    {
        List<(int Index, PptxSceneChartAxisPosition Side, ChartTextStyle Style,
            IReadOnlyList<MeasuredChartTextLine> Lines, double Height)>? measured = null;
        for (int index = 0; index < chart.Axes.Count; index++)
        {
            PptxSceneChartAxis axis = chart.Axes[index];
            if (GetDefaultAxisTitleStatus(axis) != DefaultAxisTitleStatus.Supported)
            {
                continue;
            }

            bool multiline = axis.Title.HasMultipleTextLines;
            if (!multiline && axis.PositionKind is not (PptxSceneChartAxisPosition.Top or PptxSceneChartAxisPosition.Bottom))
                continue;

            ChartTextStyle style = ReadDefaultAxisTitleStyle(theme, chart, axis);
            IReadOnlyList<ChartTextRunOverride> runs = ToChartTextRuns(axis.Title.TextRuns);
            string text = axis.Title.Text!;
            if (!multiline && !await HorizontalAxisTitleWouldClipAsync(layout, axis.PositionKind, text,
                    runs, style, fontResolver).ConfigureAwait(false))
                continue;

            var (lines, height) = await MeasureChartRichTextLinesAsync(text, runs, style, fontResolver).ConfigureAwait(false);
            (measured ??= []).Add((index, axis.PositionKind, style, lines, height));
        }

        if (measured is null) return layout;

        const double padding = PptxChartMetricRules.ChartTitleLinePadding;
        ChartFrameBox frame = layout.Frame;
        if (frame.Width <= 0d || frame.Height <= 0d) return layout;
        bool fitBody = !layout.ManualPlotLayoutApplied && !layout.Legend.Layout.HasGeometryValues;
        double header = 0d;
        if (measured.Any(title => title.Side == PptxSceneChartAxisPosition.Top) &&
            !string.IsNullOrWhiteSpace(chart.Title.Text) && !chart.Title.Layout.HasLayout)
        {
            var title = await MeasureChartTitleAsync(theme, chart, fontResolver).ConfigureAwait(false);
            header = Math.Min(Math.Max(0d, frame.Height - 1d), title.Height + PptxChartMetricRules.ChartContentFrameInset + 2d * padding);
        }

        double widthBudget = Math.Max(0d, frame.Width - 1d);
        double heightBudget = Math.Max(0d, frame.Height - header - 1d);
        double left = Requested(PptxSceneChartAxisPosition.Left);
        double right = Requested(PptxSceneChartAxisPosition.Right);
        double bottom = Requested(PptxSceneChartAxisPosition.Bottom);
        double top = Requested(PptxSceneChartAxisPosition.Top);
        if (!fitBody)
        {
            // Authored plot and legend placement constrain the bands. Keep their
            // geometry and clip the caption if a fixed margin is too small.
            ChartLayoutBox plot = layout.PlotAreaBox;
            left = Math.Min(left, Math.Max(0d, plot.X - frame.X));
            right = Math.Min(right, Math.Max(0d, frame.X + frame.Width - plot.X - plot.Width));
            bottom = Math.Min(bottom, Math.Max(0d, plot.Y - frame.Y));
            top = Math.Min(top, Math.Max(0d, frame.Y + frame.Height - header - plot.Y - plot.Height));
        }
        (left, right) = FitPair(left, right, widthBudget);
        (bottom, top) = FitPair(bottom, top, heightBudget);

        ChartFrameBox body = new(frame.X + left, frame.Y + bottom,
            frame.Width - left - right, frame.Height - bottom - top - header);
        double leftEdge = frame.X;
        double rightEdge = frame.X + frame.Width;
        double bottomEdge = frame.Y;
        double topEdge = frame.Y + frame.Height - header;
        var titles = new Dictionary<int, DefaultAxisTitleBand>(measured.Count);
        // Side bands own the corners. Top and bottom captions stop at the
        // body edges, leaving the full longitudinal span for rotated text.
        foreach (var title in measured)
        {
            double request = title.Height + 2d * padding;
            ChartLayoutBox box;
            switch (title.Side)
            {
                case PptxSceneChartAxisPosition.Left:
                    double leftWidth = Math.Min(left, request);
                    box = new(leftEdge, frame.Y, leftWidth, frame.Height);
                    left -= leftWidth;
                    leftEdge += leftWidth;
                    break;
                case PptxSceneChartAxisPosition.Right:
                    double rightWidth = Math.Min(right, request);
                    rightEdge -= rightWidth;
                    box = new(rightEdge, frame.Y, rightWidth, frame.Height);
                    right -= rightWidth;
                    break;
                case PptxSceneChartAxisPosition.Bottom:
                    double bottomHeight = Math.Min(bottom, request);
                    box = new(body.X, bottomEdge, body.Width, bottomHeight);
                    bottom -= bottomHeight;
                    bottomEdge += bottomHeight;
                    break;
                default:
                    double topHeight = Math.Min(top, request);
                    topEdge -= topHeight;
                    box = new(body.X, topEdge, body.Width, topHeight);
                    top -= topHeight;
                    break;
            }
            titles.Add(title.Index, new(box, title.Style, title.Lines, title.Height));
        }
        return ApplyDefaultAxisTitleBands(layout, new(body, fitBody, titles));

        double Requested(PptxSceneChartAxisPosition side) =>
            measured.Where(title => title.Side == side).Sum(title => title.Height + 2d * padding);

        static (double First, double Second) FitPair(double first, double second, double available)
        {
            double sum = first + second;
            if (sum <= available) return (first, second);
            double scale = available / sum;
            return (first * scale, second * scale);
        }
    }

    private static async ValueTask<bool> HorizontalAxisTitleWouldClipAsync(ChartLayout layout, PptxSceneChartAxisPosition side,
        string text, IReadOnlyList<ChartTextRunOverride> runs, ChartTextStyle style, PresentationFontResolver fontResolver)
    {
        double baseline = ComputeDefaultHorizontalAxisTitleBaselineY(layout, side);
        var measure = new ChartTextMeasurer(fontResolver);
        if (runs.Count == 0) return await WouldClipAsync(text.Trim(), style).ConfigureAwait(false);

        foreach (ChartTextRunOverride run in runs)
            if (await WouldClipAsync(run.Text, style.Merge(run.TextStyle)).ConfigureAwait(false)) return true;
        return false;

        async ValueTask<bool> WouldClipAsync(string value, ChartTextStyle runStyle)
        {
            // Keep established single-line placement when its visible ink fits the chart.
            var emission = runStyle with { FontSize = OfficePdfTextEmissionProfile.FontSize(runStyle.FontSize) };
            return await measure.TryMeasureInkBoundsAsync(value, emission).ConfigureAwait(false) is { } ink &&
                (baseline + ink.Bottom < layout.Frame.Y || baseline + ink.Top > layout.Frame.Y + layout.Frame.Height);
        }
    }

    private static ChartLayout ApplyDefaultAxisTitleBands(ChartLayout layout, DefaultAxisTitleBands? bands)
    {
        if (bands is null) return layout;
        layout = layout with { AxisTitleBands = bands };
        if (!bands.FitBody || layout.ManualPlotLayoutApplied) return layout;

        ChartFrameBox body = bands.Body;
        double sx = body.Width / layout.Frame.Width;
        double sy = body.Height / layout.Frame.Height;
        ChartPlotBox plot = layout.PlotBox;
        ChartLayoutBox area = layout.PlotAreaBox;
        return layout with
        {
            Frame = body,
            PlotBox = new(body.X + (plot.X - layout.Frame.X) * sx, body.Y + (plot.Y - layout.Frame.Y) * sy,
                plot.Width * sx, plot.Height * sy),
            PlotAreaBox = new(body.X + (area.X - layout.Frame.X) * sx, body.Y + (area.Y - layout.Frame.Y) * sy,
                area.Width * sx, area.Height * sy)
        };
    }

    private static async ValueTask RenderDefaultAxisTitleBandAsync(PdfGraphicsBuilder graphics,
        PptxSceneChart chart, PptxSceneChartAxis axis, DefaultAxisTitleBand band,
        PresentationFontResolver fontResolver, List<PdfFontResource> chartFonts, PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations, HashSet<string> reportedHyperlinkIds)
    {
        ChartLayoutBox box = band.Box;
        if (box.Width <= 0d || box.Height <= 0d) return;

        bool vertical = axis.PositionKind is PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right;
        double centerX = box.X + box.Width / 2d;
        double centerY = box.Y + box.Height / 2d;
        double inlineExtent = vertical ? box.Height : box.Width;
        double crossExtent = vertical ? box.Width : box.Height;
        double lineTop = centerY + Math.Min(crossExtent, band.Height) / 2d;
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? links = chart.Relationships is null ? null : [];
        await AddChartTitleLinesAsync(runs, axis.Title.Text!, band.Lines, band.Style, centerX - inlineExtent / 2d,
            inlineExtent, lineTop, null, box, fontResolver, links).ConfigureAwait(false);

        ChartShapeStyle shape = ToChartShapeStyle(axis.Title.ShapeStyle);
        if (!shape.IsEmpty)
        {
            ChartLayoutBox measured = await MeasureChartTitlePaintBoxAsync(runs, band.Style, fontResolver).ConfigureAwait(false);
            var paint = new PdfRectangle(measured.X, measured.Y, measured.Width, measured.Height);
            if (vertical)
                paint = TransformedRectangleBounds(paint, CenteredTransformMatrix(-90d, centerX, centerY, false, false));
            RenderInChartPlotAreaClip(graphics, new ChartPlotBox(box.X, box.Y, box.Width, box.Height),
                () => RenderChartShapeStyle(graphics, paint.X, paint.Y, paint.Width, paint.Height, shape));
        }

        if (vertical)
            RotateChartAxisTitle(runs, centerX, centerY);
        var axisPaint = await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CAT", fontResolver,
            context.DiagnosticSink, links).ConfigureAwait(false);
        AddChartTextRunHyperlinkAnnotations(links, chart, context, axisPaint, linkAnnotations, reportedHyperlinkIds);
    }
}
