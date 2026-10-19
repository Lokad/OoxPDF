using System.Globalization;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static bool TryMeasureSkewedFraction(DocxEquation equation, DocxMathFraction fraction,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources))
        { fallback = "UnsupportedMathFractionSource"; return false; }
        var resource = ResolveControlResource(fraction, sources, resolveFont, resolveControlFont);
        var font = resource?.Font.Font;
        double size = ResolveControlSize(fraction, equation, sources);
        if (font is null || !double.IsFinite(size) || size <= 0d || !font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathSkewedFractionMetrics"; return false; }
        // Both arguments keep text style. Unknown/deeper/empty arguments retain
        // complete box fallback rather than fabricating a diagonal placeholder.
        if (!TryMeasureLiteralArgument(fraction.Numerator, resolveFont, font, size, 0, token,
                out var numerator, out fallback, out bool emptyNumerator, exportFontSize: true, operatorSpacingFontSize: size) ||
            !TryMeasureLiteralArgument(fraction.Denominator, resolveFont, font, size, 0, token,
                out var denominator, out fallback, out bool emptyDenominator, exportFontSize: true, operatorSpacingFontSize: size) ||
            numerator is null || denominator is null || emptyNumerator || emptyDenominator ||
            numerator.Glyphs.Count == 0 || denominator.Glyphs.Count == 0)
        { fallback ??= "UnsupportedMathSkewedFractionArgument"; return false; }
        double scale = size / font.UnitsPerEm;
        double axis = constants.GetValue(OpenTypeMathConstant.AxisHeight) * scale;
        double horizontalGap = constants.GetValue(OpenTypeMathConstant.SkewedFractionHorizontalGap) * scale;
        double verticalGap = constants.GetValue(OpenTypeMathConstant.SkewedFractionVerticalGap) * scale;
        if (horizontalGap < 0d || verticalGap < 0d)
        { fallback = "InvalidMathSkewedFractionConstants"; return false; }
        // Font-defined ink gaps meet at the math axis. The slash is centered
        // in the horizontal gap without adding a second advance of its own.
        double numeratorY = axis + verticalGap / 2d + numerator.BelowBaseline;
        double denominatorY = axis - verticalGap / 2d - denominator.AboveBaseline;
        double denominatorX = numerator.AdvanceWidth + horizontalGap;
        double above = Math.Max(0d, numeratorY + numerator.AboveBaseline);
        double below = Math.Max(0d, denominator.BelowBaseline - denominatorY);
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size), glyphScale = glyphSize / font.UnitsPerEm;
        if (!font.TryGetNearestMathVerticalVariant(font.MapCodePoint(0x2044), (above + below) / glyphScale, token, out var variant) ||
            !font.TryGetGlyphBounds(variant.Glyph, out var bounds) || !resource!.Font.TryGetEncodedCid(variant.Glyph, out ushort cid))
        { fallback = "UnsupportedMathSkewedFractionVariant"; return false; }
        double slashAdvance = font.GetAdvanceWidth(variant.Glyph) * glyphScale;
        double slashX = numerator.AdvanceWidth + (horizontalGap - slashAdvance) / 2d;
        double slashY = axis - (bounds.YMax + bounds.YMin) / 2d * glyphScale;
        double advance = denominatorX + denominator.AdvanceWidth;
        double width = Math.Max(advance, Math.Max(numerator.Width, Math.Max(denominatorX + denominator.Width, slashX + bounds.XMax * glyphScale)));
        above = Math.Max(above, slashY + bounds.YMax * glyphScale);
        below = Math.Max(below, -slashY - bounds.YMin * glyphScale);
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) ||
            !double.IsFinite(slashX) || width <= 0d || slashX + bounds.XMin * glyphScale < 0d)
        { fallback = "InvalidMathSkewedFractionExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(numerator.Glyphs.Count + denominator.Glyphs.Count + 1);
        Append(numerator, 0d, numeratorY);
        Append(denominator, denominatorX, denominatorY);
        RgbColor color = numerator.Glyphs[0].Color;
        if (fraction.ControlText is { } control) { color = default; RgbColor.TryParse(control.ColorHex, out color); }
        glyphs.Add(new(0x2044, 0x2044, resource, cid.ToString("X4", CultureInfo.InvariantCulture), glyphSize,
            slashX, slashY, slashAdvance, color) { OriginalGlyph = variant.Glyph, ControlOwner = fraction });
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, ParentMetrics = new(font, size), EnclosureHeight = above + below,
            Edges = new(DocxMathRowLayout.AtomKind.Inner, DocxMathRowLayout.AtomKind.Inner, false) };
        fallback = null;
        return true;

        void Append(DocxMathBox child, double x, double y)
        {
            foreach (var glyph in child.Glyphs)
            { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + x, BaselineY = glyph.BaselineY + y }); }
        }
    }
}
