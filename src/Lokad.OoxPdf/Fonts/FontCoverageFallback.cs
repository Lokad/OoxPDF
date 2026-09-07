using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal readonly record struct FontCoverageSpan(int FontIndex, int Start, int Length);

internal static class FontCoverageFallback
{
    public static IReadOnlyList<FontCoverageSpan> SplitByCoverage(string text, IReadOnlyList<OpenTypeFont?> candidates, CancellationToken cancellationToken)
    {
        var spans = new List<FontCoverageSpan>();
        if (string.IsNullOrEmpty(text) || candidates.Count == 0)
        {
            return spans;
        }
        int position = 0;
        int spanStart = 0;
        int spanFont = -2;
        foreach (Rune rune in text.EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            int index = -1;
            for (int i = 0; i < candidates.Count; i++)
            {
                OpenTypeFont? candidate = candidates[i];
                if (candidate is not null && candidate.MapCodePoint(rune.Value) != 0)
                {
                    index = i;
                    break;
                }
            }
            if (index != spanFont)
            {
                if (spanFont != -2)
                {
                    spans.Add(new FontCoverageSpan(spanFont, spanStart, position - spanStart));
                }
                spanStart = position;
                spanFont = index;
            }
            position += rune.Utf16SequenceLength;
        }
        if (spanFont != -2 && position > spanStart)
        {
            spans.Add(new FontCoverageSpan(spanFont, spanStart, position - spanStart));
        }
        return spans;
    }
}
