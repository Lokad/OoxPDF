using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static async ValueTask<(ChartLayout Layout, ChartLegendBox? Box)> AdmitCartesianRightLegendAsync(
        ChartLayout layout, PptxSceneChart chart, ChartAxisSource valueAxis,
        IReadOnlyList<ChartIndexedTextPoint?> labels, bool onTickMarks, bool reversed,
        IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle legendStyle,
        PresentationFontResolver resolver, ChartLegendPlacement placement)
    {
        if (layout.ManualPlotLayoutApplied || !string.IsNullOrWhiteSpace(layout.Title) ||
            !layout.Legend.Visible || layout.Legend.Overlay || layout.Legend.PositionKind != PptxSceneChartLegendPosition.Right ||
            layout.Legend.Layout.HasGeometryValues || chart.Plots.Count != 1 ||
            chart.Axes.Count(axis => axis.AxisKind == PptxSceneChartAxisKind.Value) != 1 ||
            valueAxis.SceneAxis is not { HasLogarithmicScale: false } ||
            ReadSceneOrXmlPairedCategoryAxis(chart, valueAxis.SceneAxis, valueAxis.XmlAxis).SceneAxis is null ||
            layout.CategoryLabels is not { RotationDegrees: 0d } captions ||
            captions.Labels.Values.Any(parts => parts.Length != 1) || labels.Count == 0 || entries.Count == 0)
            return (layout, null);

        // Office controls establish horizontal, unwrapped caption admission.
        // Rotated/wrapped content retains its existing layout and legend path.
        // Preserve the already measured legend position while the inner plot
        // admits numeric and category fragments. Side stroke boxes otherwise
        // follow the plot right edge during painting.
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        await measure.PrepareAsync(entries.Select(entry => entry.Name), legendStyle).ConfigureAwait(false);
        ChartLegendBox legend = ResolveChartLegendBox(layout.Frame, layout.PlotBox, entries,
            layout.Legend, legendStyle, measure, placement);
        double margin = PptxChartMetricRules.ColumnPlotRightReserve;
        double outerLeft = layout.Frame.X + margin;
        double outerRight = layout.PlotBox.X + layout.PlotBox.Width;
        if (placement == ChartLegendPlacement.AreaRightLegend)
        {
            // The inherited area reserve already contains half its last raw
            // category. Remove that embedded band before admitting actual frames.
            int last = (labels.Count - 1) / captions.TickLabelSkip * captions.TickLabelSkip;
            if (labels[last] is { Text.Length: > 0 } label)
                outerRight += await measure.MeasureAsync(label.Text, captions.Style).ConfigureAwait(false) / 2d;
        }
        double available = outerRight - outerLeft;
        double numericBand = IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis)
            ? Math.Max(0d, layout.PlotBox.X - outerLeft) : 0d;
        bool numericRight = ResolveSceneOrXmlValueAxisLabelsRightSide(chart, valueAxis.SceneAxis, valueAxis.XmlAxis,
            ResolveSceneOrXmlValueAxisRightSide(chart, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide: false));
        double numericLeft = numericRight ? 0d : numericBand;
        double numericRightBand = numericRight ? numericBand : 0d;
        if (available - numericLeft - numericRightBand < 1d) return (layout, null);

        (double Left, double Right) RequiredBands(double width)
        {
            double left = numericLeft, right = numericRightBand;
            foreach (var (index, parts) in captions.Labels)
            {
                resolver.CancellationToken.ThrowIfCancellationRequested();
                double ratio = ChartCategoryToPlotCoordinate(index, labels.Count, 0d, 1d, onTickMarks, reversed);
                double half = parts[0].Width / 2d;
                left = Math.Max(left, half - ratio * width);
                right = Math.Max(right, half - (1d - ratio) * width);
                // Both ordinal range ends reserve their endpoint advance;
                // for a sole category the first and last source are the same.
                if (onTickMarks && index == labels.Count - 1)
                {
                    left = Math.Max(left, reversed ? half : 0d);
                    right = Math.Max(right, reversed ? 0d : half);
                }
            }
            return (left, right);
        }

        // The intersection of these linear frame constraints is convex. Zero
        // width supplies a feasible lower bound; overfull fragments keep the
        // existing fallback. Bisection finds the widest admitted inner plot.
        var minimum = RequiredBands(0d);
        if (minimum.Left + minimum.Right > available) return (layout, null);
        double lower = 0d, upper = available - numericLeft - numericRightBand;
        var upperBands = RequiredBands(upper);
        if (upper + upperBands.Left + upperBands.Right > available)
        {
            for (int step = 0; step < 40; step++)
            {
                double width = (lower + upper) / 2d;
                var bands = RequiredBands(width);
                if (width + bands.Left + bands.Right <= available) lower = width;
                else upper = width;
            }
        }
        else lower = upper;
        if (lower < 1d) return (layout, null);
        double slotWidth = lower / labels.Count * Math.Min(captions.TickLabelSkip, labels.Count) *
            PptxChartMetricRules.CategoryAxisVerticalWidthFactor;
        if (captions.Labels.Values.Any(parts => parts[0].Width > slotWidth)) return (layout, null);
        var admitted = RequiredBands(lower);
        ChartPlotBox plot = layout.PlotBox with { X = outerLeft + admitted.Left, Width = lower };
        return (layout with { PlotBox = plot, PlotAreaBox = new(plot.X, plot.Y, plot.Width, plot.Height) }, legend);
    }

    private static ChartAxisPlacement? ReadHorizontalBarCategoryPlacement(PptxSceneChart chart,
        ChartAxisSource category, ChartAxisSource value, ChartValueExtents extents, bool reversed)
    {
        // A reciprocal ordinal/numeric pair owns the source crossing. Keep
        // unpaired, composed and unsupported scales on their declared fallback.
        if (chart.Plots.Count != 1 || chart.Axes.Count(a => a.AxisKind == PptxSceneChartAxisKind.Value) != 1 ||
            category.SceneAxis is not { AxisKind: PptxSceneChartAxisKind.Category } ordinal ||
            value.SceneAxis is not { HasScaling: true, HasLogarithmicScale: false } numeric ||
            string.IsNullOrEmpty(ordinal.Id) || string.IsNullOrEmpty(numeric.Id) ||
            ordinal.CrossAxisId != numeric.Id || numeric.CrossAxisId != ordinal.Id) return null;
        return ReadChartAxisPlacement(category, extents, reversed, horizontal: false);
    }

    private static ChartAxisPlacement? ReadChartAxisPlacement(ChartAxisSource axis, ChartValueExtents extents, bool reversed, bool horizontal = true)
    {
        if (axis.SceneAxis is null && axis.XmlAxis is null) return null;
        if (ReadSceneOrXmlValueAxisCrossingValue(axis.SceneAxis, axis.XmlAxis, extents) is not { } crossing || !double.IsFinite(crossing))
            return null;
        double? explicitCrossing = axis.SceneAxis is not null ? axis.SceneAxis.CrossesAt :
            PptxSceneBuilder.ReadChartElementDoubleWithValue(axis.XmlAxis, "crossesAt").Value;
        PptxSceneChartAxisCrosses crosses = axis.SceneAxis?.CrossesKind ??
            PptxSceneBuilder.ParseChartAxisCrosses(PptxSceneBuilder.ReadChartElementValue(axis.XmlAxis, "crosses"));
        bool maximum = explicitCrossing is null && crosses == PptxSceneChartAxisCrosses.Maximum;
        bool numeric = axis.SceneAxis is { } scene ? scene.AxisKind == PptxSceneChartAxisKind.Value
            : axis.XmlAxis?.Name == ChartNamespace + "valAx";
        PptxSceneChartAxisCrossBetween between = axis.SceneAxis?.CrossBetweenKind ??
            PptxSceneBuilder.ParseChartAxisCrossBetween(PptxSceneBuilder.ReadChartElementValue(axis.XmlAxis, "crossBetween"));
        // On numeric pairs, explicit between uses half a value unit, even when
        // the source range, data or major unit changes. Omitted/midpoint modes
        // keep the numeric crossing; min/max enums still own their range edge.
        bool edge = explicitCrossing is null && crosses is PptxSceneChartAxisCrosses.Minimum or PptxSceneChartAxisCrosses.Maximum;
        if (numeric && between == PptxSceneChartAxisCrossBetween.Between && !edge) crossing -= .5d;
        double ratio = GetChartValuePlotRatio(extents, crossing, reversed);
        // Office admits reversed upper numeric next-to captions to the low band;
        // the stroke keeps its numeric coordinate and explicit low/high still win.
        if (explicitCrossing is not null && reversed && ratio == 0d) ratio = 1d;
        var labels = ResolveChartAxisLabelPlacement(axis.SceneAxis, axis.XmlAxis, ratio, reversed, maximum);
        bool axisTop = explicitCrossing is not null
            ? horizontal
                ? ResolveSceneOrXmlCategoryAxisTopSide(axis.SceneAxis, axis.XmlAxis, defaultTopSide: false)
                : ResolveSceneOrXmlCategoryAxisRightSide(axis.SceneAxis, axis.XmlAxis, defaultRightSide: false)
            : maximum != reversed;
        return new(crossing, labels.LabelPlotRatio, labels.LabelsHighSide, axisTop);
    }

    private static (double Top, double Bottom) ResolveCartesianCategoryReserves(double frameHeight,
        double headerReserve, double categoryReserve, double numericReserve, ChartAxisPlacement placement)
    {
        double top = headerReserve, bottom = categoryReserve;
        double ratio = placement.LabelPlotRatio;
        if (placement.LabelsHighSide)
        {
            bottom = numericReserve;
            double titleReserve = headerReserve - numericReserve;
            top = ratio > 0d ? Math.Max(headerReserve,
                (titleReserve + categoryReserve - (1d - ratio) * (frameHeight - bottom)) / ratio) : headerReserve;
        }
        else if (ratio > 0d)
        {
            bottom = ratio < 1d ? Math.Max(numericReserve,
                (categoryReserve - ratio * (frameHeight - top)) / (1d - ratio)) : numericReserve;
        }
        return (top, bottom);
    }

    private static (double LabelPlotRatio, bool LabelsHighSide) ResolveChartAxisLabelPlacement(
        PptxSceneChartAxis? axis, XElement? xml, double crossingRatio, bool reversed, bool maximumCrossing = false)
    {
        PptxSceneChartTickLabelPosition position = ResolveChartTickLabelPosition(axis is not null
            ? axis.TickLabelPositionKind
            : PptxSceneBuilder.ParseChartTickLabelPosition(PptxSceneBuilder.ReadChartElementValue(xml, "tickLblPos")));
        // Low/high refer to value direction, so reversing the scale swaps both
        // the physical edge and the label side. An authored maximum crossing
        // selects the high side for next-to captions; a numeric crossesAt at
        // the same bound keeps the ordinary next-to side.
        double ratio = position == PptxSceneChartTickLabelPosition.Low ? (reversed ? 1d : 0d) :
            position == PptxSceneChartTickLabelPosition.High ? (reversed ? 0d : 1d) : crossingRatio;
        bool highSide = position == PptxSceneChartTickLabelPosition.High ||
            position == PptxSceneChartTickLabelPosition.NextTo && maximumCrossing;
        return (ratio, highSide != reversed);
    }

    // Resolve spacing before placing the plot, then paint these same measured
    // labels. Automatic angles use the outer plot strip; authored angles use
    // the inner plot. Office PDF/API controls distinguish these two owners.
    private static async ValueTask<ChartLayout> PrepareCategoryLabelLayoutAsync(
        ChartLayout layout, PptxTheme theme, PptxSceneChart chart, XDocument xml,
        ChartAxisSource axis, IReadOnlyList<ChartIndexedTextPoint?> labels,
        bool onTickMarks, bool topSide, double axisRatio,
        PresentationFontResolver resolver, ChartLegendBox? legendBox = null)
    {
        if (labels.Count == 0 || !IsSceneOrXmlChartAxisLabelVisible(axis.SceneAxis, axis.XmlAxis))
            return layout;
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, chart, axis.SceneAxis, xml, axis.XmlAxis,
            PptxChartMetricRules.CategoryAxisFallbackFontSize, "categoryAxis");
        PptxSceneChartTextBodyProperties body = axis.SceneAxis?.TextBodyProperties ?? PptxSceneBuilder.ReadChartTextBodyProperties(axis.XmlAxis);
        // Empty axis body properties inherit Office's automatic orientation.
        // An authored horizontal orientation with omitted rotation means zero.
        bool automaticAngle = body.RotationValue == "-60000000" ||
            string.IsNullOrEmpty(body.RotationValue) && string.IsNullOrEmpty(body.OrientationValue);
        double angle = body.RotationDegrees is >= -90d and <= 90d ? body.RotationDegrees.Value : 0d;
        string skipSpelling = axis.SceneAxis?.TickLabelSkipValue ?? PptxSceneBuilder.ReadChartElementIntWithValue(axis.XmlAxis, "tickLblSkip").RawValue;
        bool automaticSkip = string.IsNullOrEmpty(skipSpelling);
        int skip = ResolveSceneOrXmlCategoryAxisTickLabelSkip(axis.SceneAxis, axis.XmlAxis);
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        // Office's chart font getter quantizes hundredths of a point to twips.
        // Keep emission at its existing size; this is the spacing measurement.
        ChartTextStyle policyStyle = style with { FontSize = Math.Round(style.FontSize * 20d, MidpointRounding.AwayFromZero) / 20d };
        var line = await measure.MeasureLineBoxAsync(style).ConfigureAwait(false);
        // Tick spacing includes the horizontal-header line gap. Selecting the
        // OS/2 ink box instead admits too many Arial/Times labels (Office controls).
        double policyHeight = (await measure.MeasureLineBoxAsync(policyStyle, horizontalHeader: true).ConfigureAwait(false)).Height;
        double widestWord = 0d;
        foreach (var label in labels)
        {
            resolver.CancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(label?.Text)) continue;
            foreach (string word in label.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                widestWord = Math.Max(widestWord, await measure.MeasureAsync(word, policyStyle).ConfigureAwait(false));
        }
        double availableWidth = automaticAngle && !layout.ManualPlotLayoutApplied
            ? Math.Max(1d, layout.Frame.Width - 2d * PptxChartMetricRules.ColumnPlotRightReserve)
            : layout.PlotBox.Width;
        double pitch = availableWidth / Math.Max(1, onTickMarks ? labels.Count - 1 : labels.Count);
        int RequiredSkip(double degrees)
        {
            double span = degrees == 0d ? widestWord : policyHeight / Math.Sin(Math.Abs(degrees) * Math.PI / 180d);
            return (int)Math.Clamp(Math.Ceiling(span / pitch), 1d, int.MaxValue);
        }
        if (automaticAngle)
        {
            angle = 0d;
            int needed = RequiredSkip(angle);
            foreach (double candidate in new[] { -45d, -90d })
            {
                int candidateSkip = RequiredSkip(candidate);
                if (automaticSkip ? candidateSkip < needed : needed > skip)
                {
                    angle = candidate;
                    needed = candidateSkip;
                }
            }
        }
        if (automaticSkip) skip = RequiredSkip(angle);

        double innerPitch = layout.PlotBox.Width / labels.Count;
        double flatWidth = Math.Max(1d, innerPitch * Math.Min(skip, labels.Count) * PptxChartMetricRules.CategoryAxisVerticalWidthFactor);
        var measured = new Dictionary<int, ChartTextRunLayout[]>();
        double widest = 0d;
        int maxLines = 1;
        for (int i = 0; i < labels.Count; i++)
        {
            resolver.CancellationToken.ThrowIfCancellationRequested();
            if (i % skip != 0 || string.IsNullOrWhiteSpace(labels[i]?.Text)) continue;
            string text = labels[i]!.Text;
            double width = await measure.MeasureAsync(text, style).ConfigureAwait(false);
            var lines = new List<ChartTextRunLayout>();
            if (angle == 0d && automaticSkip && width > flatWidth)
            {
                string current = string.Empty;
                foreach (string word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                {
                    string joined = current.Length == 0 ? word : current + " " + word;
                    double joinedWidth = await measure.MeasureAsync(joined, style).ConfigureAwait(false);
                    if (current.Length > 0 && joinedWidth > flatWidth)
                    {
                        lines.Add(new(current, style, await measure.MeasureAsync(current, style).ConfigureAwait(false)));
                        current = word;
                    }
                    else current = joined;
                }
                if (current.Length > 0) lines.Add(new(current, style, await measure.MeasureAsync(current, style).ConfigureAwait(false)));
            }
            else lines.Add(new(text, style, width));
            measured.Add(i, lines.ToArray());
            maxLines = Math.Max(maxLines, lines.Count);
            widest = Math.Max(widest, width);
        }
        double sin = Math.Sin(Math.Abs(angle) * Math.PI / 180d);
        double cos = Math.Cos(angle * Math.PI / 180d);
        double requested = angle == 0d ? (maxLines - 1) * line.Height : Math.Max(0d, widest * sin + line.Height * cos - line.Height);
        // A caption band must not become taller than the plot it describes.
        // For extra height e, line.Height + e <= plot.Height - e. Office
        // controls support this admission constraint, not an exact allocation.
        bool atEdge = topSide ? axisRatio == 1d : axisRatio == 0d;
        double freeEdge = topSide ? layout.Frame.Y + layout.Frame.Height : layout.Frame.Y;
        if (legendBox is { } legend && layout.Legend.Visible && !layout.Legend.Overlay &&
            layout.Legend.PositionKind == (topSide ? PptxSceneChartLegendPosition.Top : PptxSceneChartLegendPosition.Bottom))
            freeEdge = topSide ? legend.ClipY : legend.ClipY + legend.ClipHeight;
        double axisY = layout.PlotBox.Y + layout.PlotBox.Height * axisRatio;
        double existingExtra = layout.ManualPlotLayoutApplied || !atEdge
            ? Math.Max(0d, (topSide ? freeEdge - axisY : axisY - freeEdge) - ComputeCartesianLabelStripBottomReserve(style.FontSize)) : 0d;
        double admitted = layout.ManualPlotLayoutApplied || !atEdge ? 0d :
            Math.Min(requested, Math.Max(0d, (layout.PlotBox.Height - line.Height) / 2d));
        if (admitted > 0d)
        {
            double shift = topSide ? 0d : admitted;
            layout = layout with
            {
                PlotBox = layout.PlotBox with { Y = layout.PlotBox.Y + shift, Height = layout.PlotBox.Height - admitted },
                PlotAreaBox = layout.PlotAreaBox with { Y = layout.PlotAreaBox.Y + shift, Height = Math.Max(1d, layout.PlotAreaBox.Height - admitted) }
            };
        }
        double availableExtra = admitted + existingExtra;
        if (angle != 0d && availableExtra < requested)
        {
            double cap = Math.Max(0d, (line.Height + availableExtra - line.Height * cos) / sin);
            for (int index = 0; index < labels.Count; index++)
            {
                if (measured.TryGetValue(index, out var parts)) measured[index] = [await FitAsync(parts[0].Text, cap).ConfigureAwait(false)];
            }
        }
        else if (angle == 0d && availableExtra < requested)
        {
            int admittedLines = 1 + (int)Math.Floor(availableExtra / line.Height);
            for (int index = 0; index < labels.Count; index++)
            {
                if (!measured.TryGetValue(index, out var parts) || parts.Length <= admittedLines) continue;
                var fitted = parts.Take(admittedLines).ToArray();
                fitted[^1] = await FitAsync(string.Join(" ", parts.Skip(admittedLines - 1).Select(part => part.Text)), flatWidth).ConfigureAwait(false);
                measured[index] = fitted;
            }
        }
        return layout with { CategoryLabels = new(style, angle, skip, line.Ascent, line.Descent, line.Height, measured) };

        async ValueTask<ChartTextRunLayout> FitAsync(string text, double cap)
        {
            double fullWidth = await measure.MeasureAsync(text, style).ConfigureAwait(false);
            if (fullWidth <= cap) return new(text, style, fullWidth);
            string suffix = body.VerticalOverflowValue == "ellipsis" ? "\u2026" : string.Empty;
            if (await measure.MeasureAsync(suffix, style).ConfigureAwait(false) > cap) return new(string.Empty, style, 0d);
            int low = 0, high = text.Length, best = 0;
            while (low <= high)
            {
                resolver.CancellationToken.ThrowIfCancellationRequested();
                int middle = low + (high - low) / 2;
                int cut = middle > 0 && char.IsHighSurrogate(text[middle - 1]) ? middle - 1 : middle;
                if (await measure.MeasureAsync(text[..cut] + suffix, style).ConfigureAwait(false) <= cap)
                {
                    best = cut;
                    low = middle + 1;
                }
                else high = middle - 1;
            }
            string fitted = text[..best] + suffix;
            return new(fitted, style, await measure.MeasureAsync(fitted, style).ConfigureAwait(false));
        }
    }
}
