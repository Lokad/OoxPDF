using System.Buffers.Binary;
using System.Security.Cryptography;

namespace Lokad.OoxPdf.Pdf;

internal sealed partial class PdfEmbeddedFont
{
    // Only additional source identities need this lookup. Ordinary resources
    // keep their existing glyph/CID identity and allocate no alias dictionary.
    private readonly SortedDictionary<(ushort Glyph, int Unicode), ushort>? sourceAliases;
    private readonly int firstAliasCid;
    public int CidToGidMapLength => sourceAliases is { Count: > 0 } ? 2 * (firstAliasCid + sourceAliases.Count) : 0;

    public bool TryGetEncodedCid(ushort originalGlyph, int unicode, out ushort cid)
    {
        if (sourceAliases is not null && sourceAliases.TryGetValue((originalGlyph, unicode), out cid)) return true;
        return TryGetCid(originalGlyph, out cid);
    }

    private static void AddSourceMapping(SortedDictionary<ushort, int> preferred,
        ref SortedSet<(ushort Glyph, int Unicode)>? aliases, ushort glyph, int unicode)
    {
        if (!preferred.TryGetValue(glyph, out int existing)) preferred.Add(glyph, unicode);
        else if (unicode != existing)
        {
            // Preserve the independently validated shared-space normalization.
            if (unicode == 0xA0 && existing == 0x20) return;
            if (unicode == 0x20 && existing == 0xA0) preferred[glyph] = unicode;
            else
            {
                aliases ??= [];
                // Stable preference is essential: page resources with the same
                // source set must assign identical CIDs regardless of run order.
                bool replace = unicode == 0x20 || existing != 0x20 && unicode < existing;
                aliases.Add((glyph, replace ? existing : unicode));
                if (replace) preferred[glyph] = unicode;
            }
            if (preferred[glyph] == 0x20) aliases?.Remove((glyph, 0xA0));
        }
        if ((long)preferred.Count + (aliases?.Count ?? 0) > ushort.MaxValue)
            throw new OoxPdfLimitExceededException("PDF font character identities exceed the two-byte CID namespace.");
    }

    private SortedDictionary<(ushort Glyph, int Unicode), ushort>? CreateAliasCids(
        SortedSet<(ushort Glyph, int Unicode)>? aliases, CancellationToken token)
    {
        if (aliases is not { Count: > 0 }) return null;
        int next = firstAliasCid;
        if ((long)next + aliases.Count > ushort.MaxValue + 1L)
            throw new OoxPdfLimitExceededException("PDF font character identities exceed the two-byte CID namespace.");
        var result = new SortedDictionary<(ushort Glyph, int Unicode), ushort>();
        foreach (var mapping in aliases)
        {
            token.ThrowIfCancellationRequested();
            result.Add(mapping, checked((ushort)next++));
        }
        return result;
    }

    private bool HasIdenticalAliases(PdfEmbeddedFont other) =>
        (sourceAliases?.Count ?? 0) == (other.sourceAliases?.Count ?? 0) &&
        (sourceAliases is null || sourceAliases.All(pair => other.sourceAliases is not null &&
            other.sourceAliases.TryGetValue(pair.Key, out ushort cid) && pair.Value == cid));

    private static string ComputeAliasHash(SortedSet<(ushort Glyph, int Unicode)> aliases, CancellationToken token)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        Span<byte> record = stackalloc byte[6];
        foreach (var mapping in aliases)
        {
            token.ThrowIfCancellationRequested();
            BinaryPrimitives.WriteUInt16LittleEndian(record, mapping.Glyph);
            BinaryPrimitives.WriteInt32LittleEndian(record[2..], mapping.Unicode);
            hash.AppendData(record);
        }
        return Convert.ToHexString(hash.GetHashAndReset());
    }

    private IEnumerable<(ushort Cid, ushort OriginalGlyph)> EnumerateEncodedGlyphs()
    {
        foreach (ushort glyph in UnicodeByOriginalGlyph.Keys)
            if (TryGetCid(glyph, out ushort cid)) yield return (cid, glyph);
        if (sourceAliases is not null)
            foreach (var mapping in sourceAliases) yield return (mapping.Value, mapping.Key.Glyph);
    }

    internal async ValueTask WriteCidToGidMapAsync(PdfObjectWriter writer, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (sourceAliases is null) return;
        int firstAlias = firstAliasCid, total = firstAlias + sourceAliases.Count, used = 0;
        Memory<byte> block = writer.Scratch[..4096];
        using IEnumerator<KeyValuePair<(ushort Glyph, int Unicode), ushort>> aliases = sourceAliases.GetEnumerator();
        for (int cid = 0; cid < total; cid++)
        {
            if ((cid & 255) == 0) token.ThrowIfCancellationRequested();
            ushort glyph = checked((ushort)cid);
            if (cid >= firstAlias)
            {
                if (!aliases.MoveNext() || !TryGetCid(aliases.Current.Key.Glyph, out glyph))
                    throw new InvalidDataException("PDF source character has no embedded outline.");
            }
            BinaryPrimitives.WriteUInt16BigEndian(block.Span[used..], glyph);
            used += 2;
            if (used == block.Length)
            {
                await writer.WriteRawBytesAsync(block).ConfigureAwait(false);
                used = 0;
            }
        }
        if (used != 0) await writer.WriteRawBytesAsync(block[..used]).ConfigureAwait(false);
    }
}
