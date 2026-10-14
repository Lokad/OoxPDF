using System.Buffers.Binary;

namespace Lokad.OoxPdf.Fonts;

// An immutable view into the owning font's GSUB bytes. Select math/dflt/ssty
// once at font load; glyph reads do not rescan scripts/features or publish caches.
internal readonly struct OpenTypeMathScriptStyles
{
    private readonly ReadOnlyMemory<byte> bytes;
    private readonly int featureOffset, lookupOffset;
    public bool IsValid { get; }
    private OpenTypeMathScriptStyles(ReadOnlyMemory<byte> bytes, int feature, int lookup, bool valid)
    { this.bytes = bytes; featureOffset = feature; lookupOffset = lookup; IsValid = valid; }
    public static OpenTypeMathScriptStyles None => new(default, 0, 0, true);

    public static OpenTypeMathScriptStyles Read(ReadOnlyMemory<byte> bytes, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        ReadOnlySpan<byte> data = bytes.Span;
        if (data.Length < 10 || U(data, 0) != 1 || U(data, 2) > 1 || U(data, 2) == 1 && data.Length < 14) { return default; }
        int scriptOffset = U(data, 4), featuresOffset = U(data, 6), lookupsOffset = U(data, 8);
        if (!Window(data, scriptOffset, 2, out var scripts) || !Window(data, featuresOffset, 2, out var features) ||
            !Window(data, lookupsOffset, 2, out var lookups) || scripts.Length - 2 < U(scripts, 0) * 6 ||
            features.Length - 2 < U(features, 0) * 6 || lookups.Length - 2 < U(lookups, 0) * 2) { return default; }
        for (int index = 0; index < U(scripts, 0); index++)
        {
            token.ThrowIfCancellationRequested();
            int at = 2 + index * 6;
            if (BinaryPrimitives.ReadUInt32BigEndian(scripts.Slice(at, 4)) != 0x6D617468) { continue; } // math
            if (!Window(scripts, U(scripts, at + 4), 4, out var script) || script.Length - 4 < U(script, 2) * 6) { return default; }
            int languageOffset = U(script, 0);
            if (languageOffset == 0) { return None; }
            if (!Window(script, languageOffset, 6, out var language) || U(language, 0) != 0 ||
                language.Length - 6 < U(language, 4) * 2) { return default; }
            int selected = -1;
            for (int slot = -1; slot < U(language, 4); slot++)
            {
                token.ThrowIfCancellationRequested();
                int featureIndex = U(language, slot < 0 ? 2 : 6 + slot * 2);
                if (slot < 0 && featureIndex == 0xFFFF) { continue; }
                if (featureIndex >= U(features, 0)) { return default; }
                int record = 2 + featureIndex * 6;
                if (BinaryPrimitives.ReadUInt32BigEndian(features.Slice(record, 4)) != 0x73737479) { continue; } // ssty
                int relative = U(features, record + 4);
                if (!Window(features, relative, 4, out var feature) || U(feature, 0) != 0 ||
                    feature.Length - 4 < U(feature, 2) * 2) { return default; }
                if (selected >= 0 && selected != featuresOffset + relative) { return default; }
                selected = featuresOffset + relative;
            }
            return selected < 0 ? None : new(bytes, selected, lookupsOffset, true);
        }
        return None;
    }

    public bool TryGetGlyph(ushort glyph, int level, CancellationToken token, out ushort result)
    {
        result = glyph;
        token.ThrowIfCancellationRequested();
        if (!IsValid || level is < 1 or > 2) { return false; }
        if (featureOffset == 0) { return true; }
        ReadOnlySpan<byte> data = bytes.Span, feature = data.Slice(featureOffset), lookups = data.Slice(lookupOffset);
        for (int index = 0; index < U(feature, 2); index++)
        {
            token.ThrowIfCancellationRequested();
            int lookupIndex = U(feature, 4 + index * 2);
            if (lookupIndex >= U(lookups, 0) || !Window(lookups, U(lookups, 2 + lookupIndex * 2), 6, out var lookup) ||
                U(lookup, 2) != 0 || lookup.Length - 6 < U(lookup, 4) * 2) { return false; }
            int type = U(lookup, 0);
            for (int slot = 0; slot < U(lookup, 4); slot++)
            {
                token.ThrowIfCancellationRequested();
                if (!Window(lookup, U(lookup, 6 + slot * 2), 6, out var subtable)) { return false; }
                int innerType = type;
                if (type == 7)
                {
                    if (subtable.Length < 8 || U(subtable, 0) != 1) { return false; }
                    innerType = U(subtable, 2);
                    uint extension = BinaryPrimitives.ReadUInt32BigEndian(subtable.Slice(4, 4));
                    if (innerType == 7 || extension < 8 || extension > (uint)subtable.Length ||
                        subtable.Length - (int)extension < 6) { return false; }
                    subtable = subtable.Slice((int)extension);
                }
                if (!TrySubstitute(subtable, innerType, result, level, out ushort changed, out bool covered)) { return false; }
                if (covered) { result = changed; break; }
            }
        }
        return true;
    }

    private static bool TrySubstitute(ReadOnlySpan<byte> data, int type, ushort glyph, int level,
        out ushort result, out bool covered)
    {
        result = glyph; covered = false;
        int format = U(data, 0), count = U(data, 4);
        int recordsEnd = type == 1 && format == 1 ? 6 : 6 + count * 2;
        if (type is not (1 or 3) || type == 1 && format is not (1 or 2) || type == 3 && format != 1 ||
            data.Length < recordsEnd || U(data, 2) < recordsEnd ||
            !Window(data, U(data, 2), 4, out var coverage) ||
            !OpenTypeFont.TryGetCoverageIndex(coverage, glyph, out int index)) { return false; }
        if (index < 0) { return true; }
        covered = true;
        if (type == 1 && format == 1)
        {
            int delta = BinaryPrimitives.ReadInt16BigEndian(data.Slice(4, 2));
            result = (ushort)((glyph + delta) & 0xFFFF);
            return true;
        }
        if (index >= count) { return false; }
        if (type == 1) { result = U(data, 6 + index * 2); return true; }
        if (U(data, 6 + index * 2) < recordsEnd || !Window(data, U(data, 6 + index * 2), 2, out var alternates) ||
            U(alternates, 0) == 0 || alternates.Length - 2 < U(alternates, 0) * 2) { return false; }
        if (level <= U(alternates, 0)) { result = U(alternates, 2 + (level - 1) * 2); }
        return true;
    }

    private static ushort U(ReadOnlySpan<byte> data, int at) => BinaryPrimitives.ReadUInt16BigEndian(data.Slice(at, 2));
    private static bool Window(ReadOnlySpan<byte> data, int offset, int minimum, out ReadOnlySpan<byte> window)
    {
        window = default;
        if (offset < 2 || offset > data.Length || data.Length - offset < minimum) { return false; }
        window = data.Slice(offset);
        return true;
    }
}
