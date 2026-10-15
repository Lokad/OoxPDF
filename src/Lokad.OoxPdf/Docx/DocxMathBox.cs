using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxMathGlyphPlacement(
    int SourceCodePoint, int RenderCodePoint, PdfFontResource Font, string EncodedGlyph,
    double FontSize, double X, double BaselineY, double Advance, RgbColor Color)
{
    public ushort OriginalGlyph { get; init; }
    public double TextShear { get; init; }
}

internal sealed record DocxMathBox(
    double Width, double AboveBaseline, double BelowBaseline,
    IReadOnlyList<DocxMathGlyphPlacement> Glyphs)
{
    public double? LogicalAdvance { get; init; }
    public double AdvanceWidth => LogicalAdvance ?? Width;
    // Enclosing operators include a script argument's nominal descent without
    // changing the ink frame already used for paragraph flow and painting.
    public double? EnclosureBelowBaseline { get; init; }
    // A composite fraction/stack already owns its upper/lower clearances.
    // A delimiter uses that complete height; literals derive an axis frame.
    public double? EnclosureHeight { get; init; }
    public DocxMathMetricOwner? ParentMetrics { get; init; }
    public DocxMathRowLayout.RowEdges? Edges { get; init; }
    public IReadOnlyList<DocxMathRulePlacement> Rules { get; init; } = [];
}

internal sealed record DocxMathMetricOwner(Lokad.OoxPdf.Fonts.OpenTypeFont Font, double Size)
{
    // Invisible controls own a font frame without a painted glyph. A parent
    // used only for mathematical constants retains its existing frame profile.
    public bool IncludeTypographicFrame { get; init; }
}

internal sealed record DocxMathRulePlacement(double X, double Y, double Width, double Thickness, RgbColor Color);

internal static class DocxMathBoxPainter
{
    public static void Paint(DocxMathBox box, PdfGraphicsBuilder graphics, double x, double baselineY, CancellationToken token)
    {
        foreach (DocxMathRulePlacement rule in box.Rules)
        {
            token.ThrowIfCancellationRequested();
            graphics.SetFillRgb(rule.Color.Red, rule.Color.Green, rule.Color.Blue);
            graphics.FillRectangle(x + rule.X, baselineY + rule.Y, rule.Width, rule.Thickness);
        }
        foreach (DocxMathGlyphPlacement glyph in box.Glyphs)
        {
            token.ThrowIfCancellationRequested();
            double glyphX = x + glyph.X, glyphY = baselineY + glyph.BaselineY;
            if (glyph.TextShear != 0d)
            {
                graphics.SaveState();
                graphics.Transform(1d, 0d, glyph.TextShear, 1d, glyphX, glyphY);
                glyphX = glyphY = 0d;
            }
            graphics.DrawGlyphText(glyph.Font.ResourceName, glyph.FontSize, glyphX, glyphY,
                glyph.Color.Red, glyph.Color.Green, glyph.Color.Blue, glyph.EncodedGlyph,
                italic: false, characterSpacing: 0d, textRenderingMode: 0,
                strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
            if (glyph.TextShear != 0d) { graphics.RestoreState(); }
        }
    }
}
