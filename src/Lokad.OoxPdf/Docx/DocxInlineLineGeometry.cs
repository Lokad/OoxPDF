namespace Lokad.OoxPdf.Docx;

internal static class DocxInlineLineGeometry
{
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

    internal static (double Above, double Below) VerticalExtents(DocxTextLineLayout line)
    {
        double above = line.InlineFrame?.BaselineInset ?? 0d;
        double below = line.InlineFrame is { } frame ? Math.Max(0d, frame.Height - frame.BaselineInset) : 0d;
        // Source endpoints constrain horizontal extent. Selection and connector
        // share the occupied line frame, enlarged only by visible math bounds.
        foreach (var extent in line.InlineSourceExtents)
            if (extent.IsMath)
            { above = Math.Max(above, extent.AboveBaseline); below = Math.Max(below, extent.BelowBaseline); }
        return (above, below);
    }

}
