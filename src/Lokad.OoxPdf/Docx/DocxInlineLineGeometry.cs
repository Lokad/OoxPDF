namespace Lokad.OoxPdf.Docx;

internal static class DocxInlineLineGeometry
{
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
