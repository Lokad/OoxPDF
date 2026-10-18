namespace Lokad.OoxPdf.Docx;

internal static class DocxInlineLineGeometry
{
    internal static bool TryGetOccupiedBounds(DocxTextLineLayout line, out double bottom, out double top)
    {
        bottom = top = 0d;
        double? height = line.InlineFrame?.Height ?? line.BodyLineBoxHeightPoints ?? line.OccupiedLineBox?.Height;
        double? inset = line.InlineFrame?.BaselineInset ?? line.BodyLineBoxBaselineInsetPoints ?? line.OccupiedLineBox?.BaselineInset;
        if (height is not > 0d || inset is null || !double.IsFinite(height.Value) || !double.IsFinite(inset.Value) ||
            inset.Value < 0d || inset.Value > height.Value || !double.IsFinite(line.BaselineY)) return false;
        var visible = VerticalExtents(line);
        top = line.BaselineY + Math.Max(inset.Value, visible.Above);
        bottom = line.BaselineY - Math.Max(height.Value - inset.Value, visible.Below);
        return double.IsFinite(top) && double.IsFinite(bottom);
    }

    // The legacy paragraph boundary already includes the ordinary text descent.
    // Charge only extra occupied depth, so mixed inline frames are not counted twice.
    internal static double ExtraBelowTextFrame(DocxTextLineLayout line)
    {
        if (line.SourceParagraph is not { } paragraph || !DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph) ||
            line.InlineFrame is not { TextBaselineInset: { } textAbove } frame ||
            line.SingleLineHeight is not { } textHeight || !double.IsFinite(textAbove) || textAbove < 0d ||
            !double.IsFinite(textHeight) || textHeight < textAbove ||
            !double.IsFinite(frame.Height) || !double.IsFinite(frame.BaselineInset) ||
            frame.BaselineInset < 0d || frame.Height < frame.BaselineInset) return 0d;
        return Math.Max(0d, frame.Height - frame.BaselineInset - (textHeight - textAbove));
    }

    internal static bool TryGetExactBodyLineBox(DocxTextLineLayout line, out DocxOccupiedLineBox box)
    {
        box = default;
        if (line.Story?.Kind != DocxStoryKind.Body || line.LineHeightSource != DocxLineHeightSource.ExactLineSpacing ||
            line.BodyLineBoxHeightPoints is not { } height || line.BodyLineBoxBaselineInsetPoints is not { } inset ||
            !double.IsFinite(height) || height <= 0d || !double.IsFinite(inset) || inset < 0d || inset > height)
            return false;
        box = new(height, inset);
        return true;
    }

    internal static (double Above, double Below) VerticalExtents(DocxTextLineLayout line)
    {
        double above = line.InlineFrame?.BaselineInset ?? 0d;
        double below = line.InlineFrame is { } frame ? Math.Max(0d, frame.Height - frame.BaselineInset) : 0d;
        if (line.InlineFrame is null && TryGetExactBodyLineBox(line, out var authored))
        {
            above = authored.BaselineInset;
            below = authored.Height - authored.BaselineInset;
        }
        // Source endpoints constrain horizontal extent. Selection and connector
        // share the occupied line frame, enlarged only by visible math bounds.
        foreach (var extent in line.InlineSourceExtents)
            if (extent.IsMath)
            { above = Math.Max(above, extent.AboveBaseline); below = Math.Max(below, extent.BelowBaseline); }
        return (above, below);
    }

}
