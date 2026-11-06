using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool CanUseMeasuredColumnLabelFrame(ChartDataLabelOptions options)
    {
        var layout = options.Layout;
        bool sized = layout.Width.HasValue || layout.Height.HasValue;
        // Known factor extents and positions share measurement and placement.
        // Malformed values, edge extents and unknown body properties retain
        // their independently qualified fallback.
        bool measuredPlacement = !layout.HasLayout ||
            layout.HasGeometryValues &&
            (layout.X.HasValue || string.IsNullOrEmpty(layout.XValue)) &&
            (layout.Y.HasValue || string.IsNullOrEmpty(layout.YValue)) &&
            (layout.Width.HasValue || string.IsNullOrEmpty(layout.WidthValue)) &&
            (layout.Height.HasValue || string.IsNullOrEmpty(layout.HeightValue)) &&
            double.IsFinite(layout.X.GetValueOrDefault()) && double.IsFinite(layout.Y.GetValueOrDefault()) &&
            double.IsFinite(layout.Width.GetValueOrDefault()) && double.IsFinite(layout.Height.GetValueOrDefault()) &&
            layout.XMode is null or "" or "factor" or "edge" &&
            layout.YMode is null or "" or "factor" or "edge" &&
            (!sized || layout.WidthMode is null or "" or "factor" &&
                layout.HeightMode is null or "" or "factor" &&
                options.TextBodyProperties.VerticalOverflowValue is null or "" or "ellipsis");
        return measuredPlacement && options.CustomTextRuns.Count == 0 &&
            options.TextBodyProperties.WrapValue is null or "" or "square" or "none" &&
            options.TextBodyProperties.RotationDegrees.GetValueOrDefault() == 0d &&
            options.TextBodyProperties.OrientationValue is null or "" or "horz";
    }

    private static async ValueTask AddMeasuredColumnLabelAsync(
        List<TextRun> runs, PdfGraphicsBuilder graphics, string text, ChartDataLabelOptions options,
        ChartTextStyle style, ChartTextMeasurer measurer, ChartFrameBox frame,
        double centerX, double barEnd, ChartSeriesFill? keyFill, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        double horizontalPadding = PptxChartMetricRules.DataLabelHorizontalPadding;
        double verticalPadding = PptxChartMetricRules.DataLabelVerticalPadding;
        var layout = options.Layout;
        bool sized = layout.Width.HasValue || layout.Height.HasValue;
        if (sized && (layout.Width is < 0d or > 1d || layout.Height is < 0d or > 1d ||
            layout.XModeKind == PptxSceneChartManualLayoutMode.Edge && layout.X is < 0d or > 1d ||
            layout.YModeKind == PptxSceneChartManualLayoutMode.Edge && layout.Y is < 0d or > 1d))
        {
            // Native negative/oversized extents and invalid edge origins reject
            // the declaration, preserving the complete automatic label.
            layout = default;
            sized = false;
        }
        double swatch = style.FontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
        double keyReserve = keyFill is not null ? 2d * (swatch + PptxChartMetricRules.BarLegendKeyUnitCenterOffset) : 0d;
        // An absent height has no paint area. An explicit zero width hides both
        // text and key; an absent width can instead leave a key-only reference.
        if (sized && (layout.Height.GetValueOrDefault() == 0d || layout.Width == 0d ||
            !layout.Width.HasValue && keyFill is null)) return;
        bool keyOnly = sized && !layout.Width.HasValue;
        double sourceWidth = layout.Width.GetValueOrDefault() * frame.Width;
        // Automatic capacity belongs to the chart frame. Authored text capacity
        // is separate from the key reserve used to admit its default reference.
        double cap = Math.Max(1d, sized ? sourceWidth - 2d * horizontalPadding
            : frame.Width / 5d - 2d * horizontalPadding);
        bool wrapText = options.TextBodyProperties.WrapValue != "none";
        List<(string Text, double Width)> lines = keyOnly ? []
            : await MeasureChartTextLinesAsync(text, cap, style, measurer, cancellationToken, wrapText).ConfigureAwait(false);
        var metrics = await measurer.MeasureLineBoxAsync(style).ConfigureAwait(false);
        // Native Arial/Times controls distinguish line advance (hhea plus gap)
        // from the Windows baseline descent. Calibri's two heights coincide.
        double lineAdvance = (await measurer.MeasureLineBoxAsync(style, horizontalHeader: true).ConfigureAwait(false)).Height;
        // Forced overfull text elements keep their full ink advance; the
        // fitted source frame still owns the box and each line's alignment.
        double textWidth = keyOnly ? 0d : Math.Min(cap, lines.Max(line => line.Width));
        double width = sized ? keyOnly ? keyReserve : sourceWidth
            : textWidth + 2d * horizontalPadding + keyReserve;
        double height = sized ? layout.Height.GetValueOrDefault() * frame.Height
            : lines.Count * lineAdvance + 2d * verticalPadding;
        double? lastAlignmentWidth = null;
        if (sized && lines.Count > 0)
        {
            // Positive-height native boxes retain at least one source line,
            // even when its ink exceeds the requested height.
            int limit = Math.Max(1, (int)Math.Min(lines.Count,
                Math.Max(0d, Math.Floor((height - 2d * verticalPadding) / lineAdvance))));
            if (limit < lines.Count)
            {
                var fitted = await FitColumnLabelLineAsync(lines[limit - 1].Text,
                    sourceWidth - 2d * horizontalPadding, style, measurer, cancellationToken).ConfigureAwait(false);
                lines[limit - 1] = (fitted.Text, fitted.Width);
                lines.RemoveRange(limit, lines.Count - limit);
                // One-line elision retains its prefix reference; a final line
                // in a wrapped block includes the marker in its line advance.
                lastAlignmentWidth = limit == 1 ? fitted.PrefixWidth : fitted.Width;
            }
        }
        // Admit the complete occupied frame before painting its shape and text.
        // Positive outside-end labels sit above the bar, with the existing box gap.
        double maxX = Math.Max(frame.X, frame.X + frame.Width - width);
        double maxY = Math.Max(frame.Y, frame.Y + frame.Height - height);
        double referenceWidth = width + (sized && !keyOnly ? keyReserve : 0d);
        double x = Math.Clamp(centerX - referenceWidth / 2d, frame.X,
            Math.Max(frame.X, frame.X + frame.Width - referenceWidth)) +
            (sized && !keyOnly ? keyReserve / 2d : 0d);
        double y = Math.Clamp(barEnd + PptxChartMetricRules.BarLegendKeyOutEndBoxPad, frame.Y, maxY);
        // Native inward controls apply factor offsets to the admitted automatic
        // frame; edge coordinates use the full chart. Admit the resulting frame
        // before shape, swatch and paragraph painting share its geometry.
        var offsets = sized ? layout with { Width = null, Height = null,
            WidthModeKind = PptxSceneChartManualLayoutMode.Unknown, HeightModeKind = PptxSceneChartManualLayoutMode.Unknown } : layout;
        if (offsets.HasLayout && TryBuildManualLayoutBox(offsets, frame,
            new ChartLayoutBox(x, y, width, height), out ChartLayoutBox placed,
            clampToFrame: false, missingPositionModesAreFactor: true))
        {
            x = Math.Clamp(placed.X, frame.X, maxX);
            y = Math.Clamp(placed.Y, frame.Y, maxY);
        }
        RenderChartShapeStyle(graphics, x, y, width, height, options.ShapeStyle);
        double textX = x + horizontalPadding + keyReserve;
        double baseline = y + verticalPadding + metrics.Descent + (lines.Count - 1) * lineAdvance;
        if (sized) baseline += (height - lines.Count * lineAdvance - 2d * verticalPadding) / 2d;
        double keyBaseline = sized ? y + height - verticalPadding - lineAdvance + metrics.Descent : baseline;
        if (keyFill is { } fill)
            RenderFillDataLabelLegendKey(graphics, new ChartLayoutBox(
                ComputeBarLegendKeySwatchX(textX, style.FontSize), ComputeBarLegendKeySwatchY(keyBaseline, style.FontSize), swatch, swatch), style.FontSize, fill);
        for (int index = 0; index < lines.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var line = lines[index];
            // Unwrapped paragraphs keep their full advance and the complete
            // fitted paragraph span, including key space. Wrapped fragments
            // align within the text portion remaining after the key.
            double paragraphWidth = sized ? sourceWidth - 2d * horizontalPadding
                : textWidth + (wrapText ? 0d : keyReserve);
            double lineWidth = index == lines.Count - 1 && lastAlignmentWidth is { } aligned ? aligned
                : sized || !wrapText ? line.Width : Math.Min(cap, line.Width);
            double lineX = textX + (paragraphWidth - lineWidth) / 2d;
            runs.Add(CreateChartTextRun(line.Text, lineX, baseline, Math.Max(1d, line.Width), metrics.Height,
                frame.X, frame.Y, frame.Width, frame.Height, style, TextAlignment.Left));
            baseline -= lineAdvance;
        }
    }

    private static async ValueTask<(string Text, double Width, double PrefixWidth)> FitColumnLabelLineAsync(
        string text, double width, ChartTextStyle style, ChartTextMeasurer measurer, CancellationToken cancellationToken)
    {
        const string marker = "\u2026";
        await measurer.PrepareAsync(marker, style).ConfigureAwait(false);
        string prefix = text.TrimEnd(' ', '\t');
        while (prefix.Length > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            double fitted = measurer.Measure(prefix + marker, style);
            if (fitted <= width) return (prefix + marker, fitted, measurer.Measure(prefix, style));
            int word = prefix.LastIndexOfAny([' ', '\t']);
            if (word < 0) break;
            prefix = prefix[..word].TrimEnd(' ', '\t');
        }
        // If no whole word fits, split only at text-element boundaries. Keep a
        // first source element without a marker when the box cannot fit both.
        int[] elements = StringInfo.ParseCombiningCharacters(prefix);
        int low = 1, high = elements.Length, best = 0;
        while (low <= high)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int middle = low + (high - low) / 2;
            int cut = middle < elements.Length ? elements[middle] : prefix.Length;
            if (measurer.Measure(prefix[..cut] + marker, style) <= width) { best = cut; low = middle + 1; }
            else high = middle - 1;
        }
        string value = best > 0 ? prefix[..best] + marker
            : elements.Length > 0 ? prefix[..(elements.Length > 1 ? elements[1] : prefix.Length)]
            : measurer.Measure(marker, style) <= width ? marker : string.Empty;
        double advance = measurer.Measure(value, style);
        return (value, advance, best > 0 ? measurer.Measure(prefix[..best], style) : advance);
    }

    private static async ValueTask<List<(string Text, double Width)>> MeasureChartTextLinesAsync(
        string text, double cap, ChartTextStyle style, ChartTextMeasurer measurer, CancellationToken cancellationToken, bool wrapText = true)
    {
        await measurer.PrepareAsync(text, style).ConfigureAwait(false);
        var lines = new List<(string Text, double Width)>();
        void Add(string value)
        {
            value = value.TrimEnd(' ', '\t');
            lines.Add((value, Math.Max(0d, measurer.Measure(value, style))));
        }
        foreach (string paragraph in text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n'))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!wrapText)
            {
                lines.Add((paragraph, Math.Max(0d, measurer.Measure(paragraph, style))));
                continue;
            }
            string current = string.Empty;
            int cursor = 0;
            while (cursor < paragraph.Length)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int spaces = cursor;
                while (cursor < paragraph.Length && paragraph[cursor] is ' ' or '\t') cursor++;
                int start = cursor;
                while (cursor < paragraph.Length && paragraph[cursor] is not (' ' or '\t')) cursor++;
                if (start == cursor) break;
                string word = paragraph[start..cursor];
                string separator = current.Length == 0 ? string.Empty : paragraph[spaces..start];
                string candidate = current + separator + word;
                if (measurer.Measure(candidate, style) <= cap)
                {
                    current = candidate;
                    continue;
                }
                if (current.Length > 0) Add(current);
                // Split long tokens at text-element boundaries, keeping surrogate
                // pairs and combining sequences intact. Measure complete prefixes
                // so kerning and character spacing agree with emitted text.
                int[] elements = StringInfo.ParseCombiningCharacters(word);
                int first = 0;
                while (first < elements.Length)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int low = first + 1, high = elements.Length;
                    while (low < high)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int middle = low + (high - low + 1) / 2;
                        int end = middle < elements.Length ? elements[middle] : word.Length;
                        if (measurer.Measure(word[elements[first]..end], style) <= cap) low = middle;
                        else high = middle - 1;
                    }
                    int next = low < elements.Length ? elements[low] : word.Length;
                    current = word[elements[first]..next];
                    first = low;
                    if (first < elements.Length) Add(current);
                }
            }
            Add(current);
        }
        return lines;
    }
}
