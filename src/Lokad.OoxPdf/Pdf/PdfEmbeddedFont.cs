using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfEmbeddedFont
{
    private PdfEmbeddedFont(OpenTypeFont font, string baseFontName, IReadOnlyDictionary<ushort, int> unicodeByCid)
    {
        Font = font;
        BaseFontName = baseFontName;
        UnicodeByCid = unicodeByCid;
    }

    public OpenTypeFont Font { get; }

    public string BaseFontName { get; }

    public IReadOnlyDictionary<ushort, int> UnicodeByCid { get; }

    public string ResourceKey => BaseFontName;

    public static PdfEmbeddedFont Create(OpenTypeFont font, IEnumerable<int> codePoints, CancellationToken cancellationToken = default)
    {
        var unicodeByCid = new SortedDictionary<ushort, int>();
        foreach (int codePoint in codePoints)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = font.MapCodePoint(codePoint);
            if (glyph != 0)
            {
                unicodeByCid[glyph] = codePoint;
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        string fontHash = Convert.ToHexString(SHA256.HashData(font.Bytes.Span)).Substring(0, 8);
        return new PdfEmbeddedFont(font, "LOKAD+" + SanitizeName(font.FamilyName) + "-" + fontHash, unicodeByCid);
    }

    public static PdfEmbeddedFont Merge(IEnumerable<PdfEmbeddedFont> fonts, CancellationToken cancellationToken = default)
    {
        PdfEmbeddedFont[] items = fonts.ToArray();
        if (items.Length == 0)
        {
            throw new ArgumentException("At least one font is required.", nameof(fonts));
        }

        if (items.Length == 1)
        {
            return items[0];
        }

        var unicodeByCid = new SortedDictionary<ushort, int>();
        foreach (PdfEmbeddedFont font in items)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach ((ushort cid, int codePoint) in font.UnicodeByCid)
            {
                cancellationToken.ThrowIfCancellationRequested();
                unicodeByCid[cid] = codePoint;
            }
        }

        return new PdfEmbeddedFont(items[0].Font, items[0].BaseFontName, unicodeByCid);
    }

    public string BuildToUnicodeCMap(CancellationToken cancellationToken = default)
    {
        var builder = new StringBuilder();
        builder.AppendLine("/CIDInit /ProcSet findresource begin");
        builder.AppendLine("12 dict begin");
        builder.AppendLine("begincmap");
        builder.AppendLine("/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> def");
        builder.AppendLine("/CMapName /Adobe-Identity-UCS def");
        builder.AppendLine("/CMapType 2 def");
        builder.AppendLine("1 begincodespacerange");
        builder.AppendLine("<0000> <FFFF>");
        builder.AppendLine("endcodespacerange");
        builder.Append(CultureInfo.InvariantCulture, $"{UnicodeByCid.Count} beginbfchar\n");
        foreach ((ushort cid, int codePoint) in UnicodeByCid)
        {
            cancellationToken.ThrowIfCancellationRequested();
            builder.Append('<').Append(cid.ToString("X4", CultureInfo.InvariantCulture)).Append("> <");
            if (codePoint <= 0xFFFF)
            {
                builder.Append(codePoint.ToString("X4", CultureInfo.InvariantCulture));
            }
            else
            {
                int scalar = codePoint - 0x10000;
                int high = 0xD800 + (scalar >> 10);
                int low = 0xDC00 + (scalar & 0x3FF);
                builder.Append(high.ToString("X4", CultureInfo.InvariantCulture));
                builder.Append(low.ToString("X4", CultureInfo.InvariantCulture));
            }

            builder.AppendLine(">");
        }

        builder.AppendLine("endbfchar");
        builder.AppendLine("endcmap");
        builder.AppendLine("CMapName currentdict /CMap defineresource pop");
        builder.AppendLine("end");
        builder.AppendLine("end");
        return builder.ToString();
    }

    public string BuildWidthArray(CancellationToken cancellationToken = default)
    {
        if (Font.GlyphCount == 0)
        {
            return "[]";
        }

        var builder = new StringBuilder("[");
        bool first = true;
        for (int glyphId = 0; glyphId < Font.GlyphCount; glyphId++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort cid = (ushort)glyphId;
            ushort advance = Font.GetAdvanceWidth(cid);
            if (advance == 0)
            {
                continue;
            }

            if (!first)
            {
                builder.Append(' ');
            }

            first = false;
            int width = (int)Math.Round(advance * 1000d / Font.UnitsPerEm);
            builder.Append(CultureInfo.InvariantCulture, $"{cid} [{width}]");
        }

        builder.Append(']');
        return builder.ToString();
    }

    public string EncodeGlyphHex(string text)
    {
        var builder = new StringBuilder(text.Length * 4);
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = Font.MapCodePoint(rune.Value);
            if (glyph != 0)
            {
                builder.Append(glyph.ToString("X4", CultureInfo.InvariantCulture));
            }
        }

        return builder.ToString();
    }

    public string? EncodeGlyphPositioningArray(string text)
    {
        return EncodeGlyphPositioningArray(text, 0d, 1d);
    }

    public string? EncodeGlyphPositioningArray(string text, double characterSpacingPoints, double fontSize, bool forcePositioningArray = false, bool kerningEnabled = true)
    {
        var glyphs = new List<ushort>();
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = Font.MapCodePoint(rune.Value);
            if (glyph != 0)
            {
                glyphs.Add(glyph);
            }
        }

        if (glyphs.Count == 0)
        {
            return null;
        }

        double trackingAdjustment = Math.Abs(characterSpacingPoints) <= 0.001d || fontSize <= 0d
            ? 0d
            : -characterSpacingPoints * 1000d / fontSize;
        bool hasPositioning = false;
        var builder = new StringBuilder("[");
        var glyphChunk = new StringBuilder();
        for (int i = 0; i < glyphs.Count; i++)
        {
            if (i > 0)
            {
                double adjustment = trackingAdjustment;
                short kerning = kerningEnabled ? Font.GetKerning(glyphs[i - 1], glyphs[i]) : (short)0;
                if (kerning != 0)
                {
                    adjustment += -kerning * 1000d / Font.UnitsPerEm;
                }

                if (Math.Abs(adjustment) > 0.001d)
                {
                    AppendGlyphChunk(builder, glyphChunk);
                    builder.Append(' ').Append(adjustment.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ');
                    hasPositioning = true;
                }
            }

            glyphChunk.Append(glyphs[i].ToString("X4", CultureInfo.InvariantCulture));
        }

        AppendGlyphChunk(builder, glyphChunk);
        builder.Append(']');
        return hasPositioning || forcePositioningArray ? builder.ToString() : null;
    }

    private static void AppendGlyphChunk(StringBuilder builder, StringBuilder glyphChunk)
    {
        if (glyphChunk.Length == 0)
        {
            return;
        }

        builder.Append('<').Append(glyphChunk).Append('>');
        glyphChunk.Clear();
    }

    public double MeasureTextPoints(string text, double fontSize)
    {
        return MeasureTextPoints(text, fontSize, kerningEnabled: true);
    }

    public double MeasureTextPoints(string text, double fontSize, bool kerningEnabled)
    {
        double units = 0;
        ushort previousGlyph = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = Font.MapCodePoint(rune.Value);
            if (kerningEnabled && previousGlyph != 0 && glyph != 0)
            {
                units += Font.GetKerning(previousGlyph, glyph);
            }

            units += Font.GetAdvanceWidth(glyph);
            previousGlyph = glyph;
        }

        return units * fontSize / Font.UnitsPerEm;
    }

    public static string SanitizeName(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (char c in value)
        {
            builder.Append(char.IsLetterOrDigit(c) ? c : '-');
        }

        return builder.Length == 0 ? "Font" : builder.ToString();
    }
}
