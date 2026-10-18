using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private sealed record RadicalShape(ushort Glyph, string EncodedGlyph, double AdvanceHeight, double AdvanceWidth,
        double XMin, double YMin, double XMax, double YMax, IReadOnlyList<DocxMathGlyphPartPlacement> Parts);

    private static bool TryMeasureRadicalShape(PdfFontResource resource, double height, double glyphScale,
        CancellationToken token, [NotNullWhen(true)] out RadicalShape? shape)
    {
        shape = null;
        var font = resource.Font.Font;
        ushort canonical = font.MapCodePoint(0x221A);
        if (font.TryGetMathVerticalVariant(canonical, height, token, out var selected))
        {
            // An encoding/bounds failure in a selected variant remains a failure;
            // an assembly only supplies a height that ready-made shapes cannot.
            if (!font.TryGetGlyphBounds(selected.Glyph, out var bounds) ||
                !resource.Font.TryGetEncodedCid(selected.Glyph, out ushort cid)) return false;
            shape = new(selected.Glyph, cid.ToString("X4", CultureInfo.InvariantCulture), selected.Advance,
                font.GetAdvanceWidth(selected.Glyph), bounds.XMin, bounds.YMin, bounds.XMax, bounds.YMax, []);
            return true;
        }
        if (!font.TryGetMathVerticalAssembly(canonical, token, out var assembly) || assembly is null ||
            !DocxMathGlyphAssembly.TryArrange(assembly, height, token, out var arranged, out double advanceHeight) || arranged is null) return false;
        var parts = new DocxMathGlyphPartPlacement[arranged.Count];
        double xMin = double.PositiveInfinity, yMin = double.PositiveInfinity,
            xMax = double.NegativeInfinity, yMax = double.NegativeInfinity, advance = 0d;
        for (int index = 0; index < arranged.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            var part = arranged[index];
            if (!font.TryGetGlyphBounds(part.Part.Glyph, out var bounds) ||
                !resource.Font.TryGetEncodedCid(part.Part.Glyph, out ushort cid)) return false;
            string encoded = cid.ToString("X4", CultureInfo.InvariantCulture);
            parts[index] = new(part.Part.Glyph, encoded, 0d, part.Offset * glyphScale);
            advance = Math.Max(advance, font.GetAdvanceWidth(part.Part.Glyph));
            xMin = Math.Min(xMin, bounds.XMin); xMax = Math.Max(xMax, bounds.XMax);
            yMin = Math.Min(yMin, part.Offset + bounds.YMin); yMax = Math.Max(yMax, part.Offset + bounds.YMax);
        }
        shape = new(parts[0].OriginalGlyph, parts[0].EncodedGlyph, advanceHeight, advance,
            xMin, yMin, xMax, yMax, Array.AsReadOnly(parts));
        return true;
    }
}
