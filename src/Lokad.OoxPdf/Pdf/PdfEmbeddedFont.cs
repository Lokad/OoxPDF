using System.Globalization;
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

    public static PdfEmbeddedFont Create(OpenTypeFont font, IEnumerable<int> codePoints)
    {
        var unicodeByCid = new SortedDictionary<ushort, int>();
        foreach (int codePoint in codePoints)
        {
            ushort glyph = font.MapCodePoint(codePoint);
            if (glyph != 0)
            {
                unicodeByCid[glyph] = codePoint;
            }
        }

        return new PdfEmbeddedFont(font, "LOKAD+" + SanitizeName(font.FamilyName), unicodeByCid);
    }

    public string BuildToUnicodeCMap()
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

    public string BuildWidthArray()
    {
        if (UnicodeByCid.Count == 0)
        {
            return "[]";
        }

        var builder = new StringBuilder("[");
        bool first = true;
        foreach (ushort cid in UnicodeByCid.Keys.Order())
        {
            if (!first)
            {
                builder.Append(' ');
            }

            first = false;
            int width = (int)Math.Round(Font.GetAdvanceWidth(cid) * 1000d / Font.UnitsPerEm);
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

    public double MeasureTextPoints(string text, double fontSize)
    {
        double units = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = Font.MapCodePoint(rune.Value);
            units += Font.GetAdvanceWidth(glyph);
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
