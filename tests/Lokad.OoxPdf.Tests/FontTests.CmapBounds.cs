using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontTests
{
    public static void CmapCoverageRejectsGlyphsOutsideTheDeclaredFont()
    {
        foreach (ushort format in new ushort[] { 4, 12 })
        foreach (ushort glyph in new ushort[] { 0, TestFontBuilder.GlyphCount - 1, TestFontBuilder.GlyphCount, 65534 })
        {
            OpenTypeFont font = OpenTypeFont.Load(CreateCmapBoundaryFont(format, glyph));
            int codePoint = format == 4 ? 'V' : 0x1F600;
            ushort expected = glyph < TestFontBuilder.GlyphCount ? glyph : (ushort)0;
            TestAssert.Equal(expected, font.MapCodePoint(codePoint));
            TestAssert.Equal(TestFontBuilder.LoadTestFont().MapCodePoint('B'), font.MapCodePoint('B'));
        }
    }

    // Mutate an existing boundary segment or group. The format-4 recovery
    // fixture can also remove the question-mark mapping.
    internal static byte[] CreateCmapBoundaryFont(ushort format, ushort glyph, bool omitQuestionMark = false)
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        int cmap = TestFontBuilder.GetTableRange(bytes, "cmap").Offset;
        if (format == 4)
        {
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(cmap + 2), 1);
            int subtable = cmap + checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(cmap + 8)));
            int segments = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(subtable + 6)) / 2;
            for (int i = 0; i < segments; i++)
            {
                int codePoint = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(subtable + 14 + i * 2));
                if (codePoint != 'V' && !(omitQuestionMark && codePoint == '?')) continue;
                int mappedGlyph = codePoint == 'V' ? glyph : 0;
                int delta = (mappedGlyph - codePoint) & 0xFFFF;
                BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(subtable + 16 + segments * 4 + i * 2), (ushort)delta);
            }
            return bytes;
        }
        else if (format == 12)
        {
            int subtable = cmap + checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(cmap + 16)));
            int groups = checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(subtable + 12)));
            for (int i = 0; i < groups; i++)
            {
                int group = subtable + 16 + i * 12;
                if (BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(group)) != 0x1F600) continue;
                BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(group + 8), glyph);
                return bytes;
            }
        }
        throw new InvalidOperationException("Expected the synthetic cmap boundary segment.");
    }
}
