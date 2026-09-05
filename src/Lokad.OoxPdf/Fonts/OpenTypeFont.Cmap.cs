using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    // One cmap subtable parser. Implementors read a single subtable at their stored
    // offset: Map returns the glyph id for a code point, or 0 when unmapped.
    // bytes is always the whole font file, so stored offsets stay absolute.
    private abstract record CmapFormat
    {
        public abstract ushort Map(int codePoint, byte[] bytes);
    }

    private sealed record CmapFormat4(ushort[] EndCodes, ushort[] StartCodes, short[] IdDeltas, ushort[] IdRangeOffsets, int GlyphArrayOffset) : CmapFormat
    {
        public static CmapFormat4 Read(byte[] bytes, int offset)
        {
            ushort segCount = (ushort)(U16(bytes, offset + 6) / 2);
            int endCodesOffset = offset + 14;
            int startCodesOffset = endCodesOffset + segCount * 2 + 2;
            int idDeltasOffset = startCodesOffset + segCount * 2;
            int idRangeOffsetsOffset = idDeltasOffset + segCount * 2;
            var endCodes = new ushort[segCount];
            var startCodes = new ushort[segCount];
            var idDeltas = new short[segCount];
            var idRangeOffsets = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
            {
                endCodes[i] = U16(bytes, endCodesOffset + i * 2);
                startCodes[i] = U16(bytes, startCodesOffset + i * 2);
                idDeltas[i] = unchecked((short)U16(bytes, idDeltasOffset + i * 2));
                idRangeOffsets[i] = U16(bytes, idRangeOffsetsOffset + i * 2);
            }

            return new CmapFormat4(endCodes, startCodes, idDeltas, idRangeOffsets, idRangeOffsetsOffset + segCount * 2);
        }

        public override ushort Map(int codePoint, byte[] bytes)
        {
            if (codePoint > ushort.MaxValue)
            {
                return 0;
            }

            ushort c = (ushort)codePoint;
            for (int i = 0; i < EndCodes.Length; i++)
            {
                if (c < StartCodes[i] || c > EndCodes[i])
                {
                    continue;
                }

                if (IdRangeOffsets[i] == 0)
                {
                    return (ushort)((c + IdDeltas[i]) & 0xFFFF);
                }

                int idRangeOffsetAddress = GlyphArrayOffset - IdRangeOffsets.Length * 2 + i * 2;
                int glyphOffset = idRangeOffsetAddress + IdRangeOffsets[i] + (c - StartCodes[i]) * 2;
                ushort glyph = U16(bytes, glyphOffset);
                return glyph == 0 ? (ushort)0 : (ushort)((glyph + IdDeltas[i]) & 0xFFFF);
            }

            return 0;
        }
    }

    private sealed record CmapFormat12(CmapGroup[] Groups) : CmapFormat
    {
        public static CmapFormat12 Read(byte[] bytes, int offset)
        {
            uint count = U32(bytes, offset + 12);
            var groups = new CmapGroup[count];
            int groupOffset = offset + 16;
            for (int i = 0; i < groups.Length; i++)
            {
                groups[i] = new CmapGroup(U32(bytes, groupOffset), U32(bytes, groupOffset + 4), U32(bytes, groupOffset + 8));
                groupOffset += 12;
            }

            return new CmapFormat12(groups);
        }

        public override ushort Map(int codePoint, byte[] bytes)
        {
            uint c = (uint)codePoint;
            foreach (CmapGroup group in Groups)
            {
                if (c >= group.StartCode && c <= group.EndCode)
                {
                    uint glyph = group.StartGlyph + c - group.StartCode;
                    return glyph <= ushort.MaxValue ? (ushort)glyph : (ushort)0;
                }
            }

            return 0;
        }
    }

    private readonly record struct CmapGroup(uint StartCode, uint EndCode, uint StartGlyph);
}
