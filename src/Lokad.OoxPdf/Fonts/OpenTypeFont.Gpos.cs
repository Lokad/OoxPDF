using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    private static void ReadLegacyKerningPairs(byte[] bytes, Dictionary<string, TableRecord> tables, Dictionary<uint, short> pairs)
    {
        if (!tables.TryGetValue("kern", out TableRecord kern) || kern.Length < 4)
        {
            return;
        }

        ushort tableCount = U16(bytes, kern.Offset + 2);
        int subtableOffset = kern.Offset + 4;
        for (int table = 0; table < tableCount && subtableOffset + 6 <= kern.Offset + kern.Length; table++)
        {
            ushort length = U16(bytes, subtableOffset + 2);
            ushort coverage = U16(bytes, subtableOffset + 4);
            int format = coverage >> 8;
            bool horizontal = (coverage & 0x0001) != 0;
            if (length >= 14 && format == 0 && horizontal)
            {
                ushort pairCount = U16(bytes, subtableOffset + 6);
                int pairOffset = subtableOffset + 14;
                int pairEnd = Math.Min(subtableOffset + length, kern.Offset + kern.Length);
                for (int i = 0; i < pairCount && pairOffset + 6 <= pairEnd; i++)
                {
                    ushort left = U16(bytes, pairOffset);
                    ushort right = U16(bytes, pairOffset + 2);
                    short value = I16(bytes, pairOffset + 4);
                    if (value != 0)
                    {
                        pairs[((uint)left << 16) | right] = value;
                    }

                    pairOffset += 6;
                }
            }

            if (length == 0)
            {
                break;
            }

            subtableOffset += length;
        }
    }

    private static void ReadGposPairAdjustments(byte[] bytes, Dictionary<string, TableRecord> tables, Dictionary<uint, short> pairs)
    {
        if (!tables.TryGetValue("GPOS", out TableRecord gpos) || gpos.Length < 10)
        {
            return;
        }

        int tableEnd = gpos.Offset + gpos.Length;
        int lookupList = gpos.Offset + U16(bytes, gpos.Offset + 8);
        if (lookupList + 2 > tableEnd)
        {
            return;
        }

        HashSet<ushort> kernLookupIndices = ReadGposKernLookupIndices(bytes, gpos.Offset, tableEnd);
        if (kernLookupIndices.Count == 0)
        {
            return;
        }

        ushort lookupCount = U16(bytes, lookupList);
        for (int i = 0; i < lookupCount && lookupList + 2 + i * 2 + 2 <= tableEnd; i++)
        {
            if (!kernLookupIndices.Contains((ushort)i))
            {
                continue;
            }

            int lookup = lookupList + U16(bytes, lookupList + 2 + i * 2);
            if (lookup + 6 > tableEnd)
            {
                continue;
            }

            ushort lookupType = U16(bytes, lookup);
            ushort subtableCount = U16(bytes, lookup + 4);
            for (int j = 0; j < subtableCount && lookup + 6 + j * 2 + 2 <= tableEnd; j++)
            {
                int subtable = lookup + U16(bytes, lookup + 6 + j * 2);
                ReadGposPairAdjustmentLookupSubtable(bytes, subtable, lookupType, tableEnd, pairs);
            }
        }
    }

    private static readonly HashSet<string> KernScriptTags = new(StringComparer.Ordinal)
    {
        "latn",
        "cyrl",
        "grek",
        "DFLT",
    };

    private static HashSet<ushort> ReadGposKernLookupIndices(byte[] bytes, int gposOffset, int tableEnd)
    {
        var lookupIndices = new HashSet<ushort>();
        int featureList = gposOffset + U16(bytes, gposOffset + 6);
        if (featureList + 2 > tableEnd)
        {
            return lookupIndices;
        }

        HashSet<int>? scriptFeatures = ReadGposKernScriptFeatureIndices(bytes, gposOffset, tableEnd);
        ushort featureCount = U16(bytes, featureList);
        for (int i = 0; i < featureCount && featureList + 2 + i * 6 + 6 <= tableEnd; i++)
        {
            int record = featureList + 2 + i * 6;
            string tag = Encoding.ASCII.GetString(bytes, record, 4);
            if (!tag.Equals("kern", StringComparison.Ordinal))
            {
                continue;
            }

            if (scriptFeatures is not null && !scriptFeatures.Contains(i))
            {
                continue;
            }

            int feature = featureList + U16(bytes, record + 4);
            if (feature + 4 > tableEnd)
            {
                continue;
            }

            ushort lookupIndexCount = U16(bytes, feature + 2);
            for (int j = 0; j < lookupIndexCount && feature + 4 + j * 2 + 2 <= tableEnd; j++)
            {
                lookupIndices.Add(U16(bytes, feature + 4 + j * 2));
            }
        }

        return lookupIndices;
    }

    private static HashSet<int>? ReadGposKernScriptFeatureIndices(byte[] bytes, int gposOffset, int tableEnd)
    {
        int scriptList = gposOffset + U16(bytes, gposOffset + 4);
        if (scriptList + 2 > tableEnd)
        {
            return null;
        }

        ushort scriptCount = U16(bytes, scriptList);
        var accepted = new HashSet<int>();
        for (int i = 0; i < scriptCount && scriptList + 2 + i * 6 + 6 <= tableEnd; i++)
        {
            int record = scriptList + 2 + i * 6;
            string tag = Encoding.ASCII.GetString(bytes, record, 4).Trim();
            if (!KernScriptTags.Contains(tag))
            {
                continue;
            }

            int script = scriptList + U16(bytes, record + 4);
            if (script + 4 > tableEnd)
            {
                continue;
            }

            int defaultLangSys = U16(bytes, script);
            if (defaultLangSys != 0)
            {
                CollectScriptFeatureIndices(bytes, script + defaultLangSys, tableEnd, accepted);
            }

            ushort langSysCount = U16(bytes, script + 2);
            for (int j = 0; j < langSysCount && script + 4 + j * 6 + 6 <= tableEnd; j++)
            {
                int langRecord = script + 4 + j * 6;
                CollectScriptFeatureIndices(bytes, script + U16(bytes, langRecord + 4), tableEnd, accepted);
            }
        }

        return accepted.Count > 0 ? accepted : null;
    }

    private static void CollectScriptFeatureIndices(byte[] bytes, int langSys, int tableEnd, HashSet<int> accepted)
    {
        if (langSys + 6 > tableEnd)
        {
            return;
        }

        ushort featureCount = U16(bytes, langSys + 4);
        for (int i = 0; i < featureCount && langSys + 6 + i * 2 + 2 <= tableEnd; i++)
        {
            accepted.Add(U16(bytes, langSys + 6 + i * 2));
        }
    }

    private static void ReadGposPairAdjustmentLookupSubtable(
        byte[] bytes,
        int subtable,
        ushort lookupType,
        int tableEnd,
        Dictionary<uint, short> pairs)
    {
        if (lookupType == 2)
        {
            ReadGposPairAdjustmentSubtable(bytes, subtable, tableEnd, pairs);
            return;
        }

        if (lookupType != 9 || subtable + 8 > tableEnd || U16(bytes, subtable) != 1)
        {
            return;
        }

        ushort extensionLookupType = U16(bytes, subtable + 2);
        uint extensionOffset = U32(bytes, subtable + 4);
        if (extensionLookupType != 2 || extensionOffset > int.MaxValue)
        {
            return;
        }

        int extensionSubtable = subtable + (int)extensionOffset;
        if (extensionSubtable < subtable || extensionSubtable >= tableEnd)
        {
            return;
        }

        ReadGposPairAdjustmentSubtable(bytes, extensionSubtable, tableEnd, pairs);
    }

    private static void ReadGposPairAdjustmentSubtable(byte[] bytes, int subtable, int tableEnd, Dictionary<uint, short> pairs)
    {
        if (subtable + 10 > tableEnd)
        {
            return;
        }

        ushort positionFormat = U16(bytes, subtable);
        int coverage = subtable + U16(bytes, subtable + 2);
        ushort valueFormat1 = U16(bytes, subtable + 4);
        ushort valueFormat2 = U16(bytes, subtable + 6);
        int valueRecordSize1 = ValueRecordSize(valueFormat1);
        int valueRecordSize2 = ValueRecordSize(valueFormat2);
        if (!TryReadCoverage(bytes, coverage, tableEnd, out ushort[] coverageGlyphs))
        {
            return;
        }

        if (positionFormat == 1)
        {
            ushort pairSetCount = U16(bytes, subtable + 8);
            for (int i = 0; i < pairSetCount && i < coverageGlyphs.Length && subtable + 10 + i * 2 + 2 <= tableEnd; i++)
            {
                int pairSet = subtable + U16(bytes, subtable + 10 + i * 2);
                if (pairSet + 2 > tableEnd)
                {
                    continue;
                }

                ushort pairValueCount = U16(bytes, pairSet);
                int pairValue = pairSet + 2;
                for (int j = 0; j < pairValueCount && pairValue + 2 + valueRecordSize1 + valueRecordSize2 <= tableEnd; j++)
                {
                    ushort rightGlyph = U16(bytes, pairValue);
                    short xAdvance = ReadXAdvance(bytes, pairValue + 2, valueFormat1);
                    if (xAdvance != 0)
                    {
                        pairs[((uint)coverageGlyphs[i] << 16) | rightGlyph] = xAdvance;
                    }

                    pairValue += 2 + valueRecordSize1 + valueRecordSize2;
                }
            }
        }
        else if (positionFormat == 2 && subtable + 16 <= tableEnd)
        {
            ushort class1Count = U16(bytes, subtable + 12);
            ushort class2Count = U16(bytes, subtable + 14);
            int classDef1 = subtable + U16(bytes, subtable + 8);
            int classDef2 = subtable + U16(bytes, subtable + 10);
            int classRecord = subtable + 16;
            int classRecordSize = valueRecordSize1 + valueRecordSize2;
            if (classRecordSize == 0)
            {
                return;
            }

            Dictionary<ushort, ushort[]> rightGlyphsByClass = ReadClassGlyphs(bytes, classDef2, tableEnd);
            foreach (ushort leftGlyph in coverageGlyphs)
            {
                ushort leftClass = ReadGlyphClass(bytes, classDef1, tableEnd, leftGlyph);
                if (leftClass >= class1Count)
                {
                    continue;
                }

                foreach (KeyValuePair<ushort, ushort[]> rightClassGlyphs in rightGlyphsByClass)
                {
                    if (rightClassGlyphs.Key == 0 || rightClassGlyphs.Key >= class2Count)
                    {
                        continue;
                    }

                    int record = classRecord + ((leftClass * class2Count + rightClassGlyphs.Key) * classRecordSize);
                    if (record + classRecordSize > tableEnd)
                    {
                        return;
                    }

                    short xAdvance = ReadXAdvance(bytes, record, valueFormat1);
                    if (xAdvance != 0)
                    {
                        foreach (ushort rightGlyph in rightClassGlyphs.Value)
                        {
                            pairs[((uint)leftGlyph << 16) | rightGlyph] = xAdvance;
                        }
                    }
                }
            }
        }
    }

    private static int ValueRecordSize(ushort valueFormat)
    {
        int size = 0;
        for (int bit = 0; bit < 8; bit++)
        {
            if ((valueFormat & (1 << bit)) != 0)
            {
                size += 2;
            }
        }

        return size;
    }

    private static short ReadXAdvance(byte[] bytes, int offset, ushort valueFormat)
    {
        for (int bit = 0; bit < 4; bit++)
        {
            if ((valueFormat & (1 << bit)) == 0)
            {
                continue;
            }

            if (bit == 2)
            {
                return I16(bytes, offset);
            }

            offset += 2;
        }

        return 0;
    }

    private static bool TryReadCoverage(byte[] bytes, int offset, int tableEnd, out ushort[] glyphs)
    {
        glyphs = [];
        if (offset + 4 > tableEnd)
        {
            return false;
        }

        ushort format = U16(bytes, offset);
        ushort count = U16(bytes, offset + 2);
        if (format == 1)
        {
            if (offset + 4 + count * 2 > tableEnd)
            {
                return false;
            }

            glyphs = new ushort[count];
            for (int i = 0; i < glyphs.Length; i++)
            {
                glyphs[i] = U16(bytes, offset + 4 + i * 2);
            }

            return true;
        }

        if (format != 2 || offset + 4 + count * 6 > tableEnd)
        {
            return false;
        }

        var list = new List<ushort>();
        for (int i = 0; i < count; i++)
        {
            int range = offset + 4 + i * 6;
            ushort start = U16(bytes, range);
            ushort end = U16(bytes, range + 2);
            for (int glyph = start; glyph <= end; glyph++)
            {
                list.Add((ushort)glyph);
            }
        }

        glyphs = list.ToArray();
        return true;
    }

    private static ushort ReadGlyphClass(byte[] bytes, int offset, int tableEnd, ushort glyphId)
    {
        if (offset + 4 > tableEnd)
        {
            return 0;
        }

        ushort format = U16(bytes, offset);
        if (format == 1)
        {
            ushort start = U16(bytes, offset + 2);
            ushort count = U16(bytes, offset + 4);
            int classOffset = offset + 6;
            if (glyphId < start || glyphId >= start + count || classOffset + count * 2 > tableEnd)
            {
                return 0;
            }

            return U16(bytes, classOffset + (glyphId - start) * 2);
        }

        if (format != 2)
        {
            return 0;
        }

        ushort rangeCount = U16(bytes, offset + 2);
        if (offset + 4 + rangeCount * 6 > tableEnd)
        {
            return 0;
        }

        for (int i = 0; i < rangeCount; i++)
        {
            int range = offset + 4 + i * 6;
            ushort start = U16(bytes, range);
            ushort end = U16(bytes, range + 2);
            if (glyphId >= start && glyphId <= end)
            {
                return U16(bytes, range + 4);
            }
        }

        return 0;
    }

    private static Dictionary<ushort, ushort[]> ReadClassGlyphs(byte[] bytes, int offset, int tableEnd)
    {
        var classes = new Dictionary<ushort, List<ushort>>();
        if (offset + 4 > tableEnd)
        {
            return [];
        }

        ushort format = U16(bytes, offset);
        if (format == 1)
        {
            ushort start = U16(bytes, offset + 2);
            ushort count = U16(bytes, offset + 4);
            int classOffset = offset + 6;
            if (classOffset + count * 2 > tableEnd)
            {
                return [];
            }

            for (int i = 0; i < count; i++)
            {
                ushort classValue = U16(bytes, classOffset + i * 2);
                if (classValue == 0)
                {
                    continue;
                }

                AddClassGlyph(classes, classValue, (ushort)(start + i));
            }
        }
        else if (format == 2)
        {
            ushort rangeCount = U16(bytes, offset + 2);
            if (offset + 4 + rangeCount * 6 > tableEnd)
            {
                return [];
            }

            for (int i = 0; i < rangeCount; i++)
            {
                int range = offset + 4 + i * 6;
                ushort start = U16(bytes, range);
                ushort end = U16(bytes, range + 2);
                ushort classValue = U16(bytes, range + 4);
                if (classValue == 0)
                {
                    continue;
                }

                for (int glyph = start; glyph <= end; glyph++)
                {
                    AddClassGlyph(classes, classValue, (ushort)glyph);
                }
            }
        }

        return classes.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray());
    }

    private static void AddClassGlyph(Dictionary<ushort, List<ushort>> classes, ushort classValue, ushort glyphId)
    {
        if (!classes.TryGetValue(classValue, out List<ushort>? glyphs))
        {
            glyphs = [];
            classes[classValue] = glyphs;
        }

        glyphs.Add(glyphId);
    }
}
