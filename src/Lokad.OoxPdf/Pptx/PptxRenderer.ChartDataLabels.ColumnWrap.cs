using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool CanUseMeasuredColumnLabelFrame(ChartDataLabelOptions options) =>
        !options.Layout.HasLayout && options.CustomTextRuns.Count == 0 &&
        options.TextBodyProperties.WrapValue is null or "" or "square" or "none" &&
        options.TextBodyProperties.RotationDegrees.GetValueOrDefault() == 0d &&
        options.TextBodyProperties.OrientationValue is null or "" or "horz";

    private static async ValueTask AddAutomaticColumnLabelAsync(
        List<TextRun> runs, PdfGraphicsBuilder graphics, string text, ChartDataLabelOptions options,
        ChartTextStyle style, ChartTextMeasurer measurer, ChartFrameBox frame,
        double centerX, double barEnd, ChartSeriesFill? keyFill, CancellationToken cancellationToken)
    {
        double horizontalPadding = PptxChartMetricRules.DataLabelHorizontalPadding;
        double verticalPadding = PptxChartMetricRules.DataLabelVerticalPadding;
        // Independent Office controls retain this frame-owned cap when category
        // count, plot width, font and key presence change. Key space is additional.
        double cap = Math.Max(1d, frame.Width / 5d - 2d * horizontalPadding);
        bool wrapText = options.TextBodyProperties.WrapValue != "none";
        var lines = await WrapChartTextAsync(text, wrapText ? cap : double.MaxValue, style, measurer, cancellationToken).ConfigureAwait(false);
        var metrics = await measurer.MeasureLineBoxAsync(style).ConfigureAwait(false);
        // Native Arial/Times controls distinguish line advance (hhea plus gap)
        // from the Windows baseline descent. Calibri's two heights coincide.
        double lineAdvance = (await measurer.MeasureLineBoxAsync(style, horizontalHeader: true).ConfigureAwait(false)).Height;
        // Forced overfull text elements keep their full ink advance; the
        // fitted source frame still owns the box and each line's alignment.
        double textWidth = Math.Min(cap, lines.Max(line => line.Width));
        double swatch = style.FontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
        double keyReserve = keyFill is not null ? 2d * (swatch + PptxChartMetricRules.BarLegendKeyUnitCenterOffset) : 0d;
        double width = textWidth + 2d * horizontalPadding + keyReserve;
        double height = lines.Count * lineAdvance + 2d * verticalPadding;
        // Admit the complete occupied frame before painting its shape and text.
        // Positive outside-end labels sit above the bar, with the existing box gap.
        double x = Math.Clamp(centerX - width / 2d, frame.X, Math.Max(frame.X, frame.X + frame.Width - width));
        double y = Math.Clamp(barEnd + PptxChartMetricRules.BarLegendKeyOutEndBoxPad,
            frame.Y, Math.Max(frame.Y, frame.Y + frame.Height - height));
        RenderChartShapeStyle(graphics, x, y, width, height, options.ShapeStyle);
        double textX = x + horizontalPadding + keyReserve;
        double baseline = y + verticalPadding + metrics.Descent + (lines.Count - 1) * lineAdvance;
        if (keyFill is { } fill)
            RenderFillDataLabelLegendKey(graphics, new ChartLayoutBox(
                ComputeBarLegendKeySwatchX(textX, style.FontSize), ComputeBarLegendKeySwatchY(baseline, style.FontSize), swatch, swatch), style.FontSize, fill);
        foreach (var line in lines)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Unwrapped paragraphs keep their full advance and the complete
            // fitted paragraph span, including key space. Wrapped fragments
            // align within the text portion remaining after the key.
            double paragraphWidth = textWidth + (wrapText ? 0d : keyReserve);
            double lineWidth = wrapText ? Math.Min(cap, line.Width) : line.Width;
            double lineX = textX + (paragraphWidth - lineWidth) / 2d;
            runs.Add(CreateChartTextRun(line.Text, lineX, baseline, Math.Max(1d, line.Width), metrics.Height,
                frame.X, frame.Y, frame.Width, frame.Height, style, TextAlignment.Left));
            baseline -= lineAdvance;
        }
    }

    private static async ValueTask<List<(string Text, double Width)>> WrapChartTextAsync(
        string text, double cap, ChartTextStyle style, ChartTextMeasurer measurer, CancellationToken cancellationToken)
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
