using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private sealed record VerticalGlyphShape(ushort Glyph, string EncodedGlyph, double AdvanceHeight, double AdvanceWidth,
        double XMin, double YMin, double XMax, double YMax, IReadOnlyList<DocxMathGlyphPartPlacement> Parts);

    private static bool TryMeasureVerticalGlyphShape(PdfFontResource resource, int point, double height, double glyphScale,
        bool nearest, CancellationToken token, [NotNullWhen(true)] out VerticalGlyphShape? shape)
    {
        shape = null;
        var font = resource.Font.Font;
        ushort canonical = font.MapCodePoint(point);
        bool ready = nearest
            ? font.TryGetNearestMathVerticalVariant(canonical, height, token, out var selected)
            : font.TryGetMathVerticalVariant(canonical, height, token, out selected);
        if (ready)
        {
            // An encoding/bounds failure in a selected variant remains a failure;
            // an assembly only supplies a height that ready-made shapes cannot.
            return TryMeasureFixedGlyphShape(resource, selected.Glyph, selected.Advance, out shape);
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

    private static bool TryMeasureFixedGlyphShape(PdfFontResource resource, ushort glyph, double advanceHeight,
        [NotNullWhen(true)] out VerticalGlyphShape? shape)
    {
        shape = null;
        var font = resource.Font.Font;
        if (glyph == 0 || !font.TryGetGlyphBounds(glyph, out var bounds) ||
            !resource.Font.TryGetEncodedCid(glyph, out ushort cid)) return false;
        shape = new(glyph, cid.ToString("X4", CultureInfo.InvariantCulture), advanceHeight,
            font.GetAdvanceWidth(glyph), bounds.XMin, bounds.YMin, bounds.XMax, bounds.YMax, []);
        return true;
    }
}
