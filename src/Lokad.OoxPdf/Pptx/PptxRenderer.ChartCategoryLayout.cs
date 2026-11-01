using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static (double LabelPlotRatio, bool LabelsTopSide) ResolveCategoryLabelPlacement(
        PptxSceneChartAxis? axis, XElement? xml, double crossingRatio, bool reversed)
    {
        PptxSceneChartTickLabelPosition position = ResolveChartTickLabelPosition(axis is not null
            ? axis.TickLabelPositionKind
            : PptxSceneBuilder.ParseChartTickLabelPosition(PptxSceneBuilder.ReadChartElementValue(xml, "tickLblPos")));
        // Low/high refer to value direction, so reversing the scale swaps both
        // the physical edge and the label side. NextTo follows the crossing.
        double ratio = position == PptxSceneChartTickLabelPosition.Low ? (reversed ? 1d : 0d) :
            position == PptxSceneChartTickLabelPosition.High ? (reversed ? 0d : 1d) : crossingRatio;
        return (ratio, (position == PptxSceneChartTickLabelPosition.High) != reversed);
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
            foreach (string word in label.Value.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
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
            string text = labels[i]!.Value.Text;
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
