using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfEmbeddedFont
{
    private readonly IReadOnlyDictionary<ushort, ushort> cidByOriginalGlyph;
    private readonly byte[]? subsetFontBytes;

    private PdfEmbeddedFont(
        OpenTypeFont font,
        string baseFontName,
        string codepointSetHash,
        IReadOnlyDictionary<ushort, int> unicodeByOriginalGlyph,
        OpenTypeFontSubset? subset)
    {
        Font = font;
        BaseFontName = baseFontName;
        CodepointSetHash = codepointSetHash;
        UnicodeByOriginalGlyph = unicodeByOriginalGlyph;
        if (subset is not null)
        {
            cidByOriginalGlyph = subset.CidByOriginalGlyph;
            subsetFontBytes = subset.Bytes;
        }
        else
        {
            cidByOriginalGlyph = new Dictionary<ushort, ushort>();
        }

        UnicodeByCid = BuildUnicodeByCid();
        // ResourceKey is read on every writer lookup (grouping, object
        // numbering, content emission), often in per-rune paths. Cache the key once
        // instead of concatenating a fresh string per access.
        ResourceKey = BaseFontName + "-U" + CodepointSetHash;
    }

    public OpenTypeFont Font { get; }

    public string BaseFontName { get; }

    public IReadOnlyDictionary<ushort, int> UnicodeByOriginalGlyph { get; }

    public IReadOnlyDictionary<ushort, int> UnicodeByCid { get; }

    public ReadOnlyMemory<byte> FontProgramBytes => subsetFontBytes is null ? Font.Bytes : subsetFontBytes;

    public bool UsesSubsetFontProgram => subsetFontBytes is not null;

    // Subset CIDs belong to the selected glyph mapping. Canonical character
    // mappings retain their established identity; alternate glyphs also include
    // glyph-to-Unicode ownership so distinct remaps cannot alias resources.
    public string CodepointSetHash { get; }

    public string ResourceKey { get; }

    public static PdfEmbeddedFont Create(OpenTypeFont font, IEnumerable<int> codePoints, CancellationToken cancellationToken)
    {
        var unicodeByOriginalGlyph = new SortedDictionary<ushort, int>();
        foreach (int codePoint in codePoints)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = font.MapCodePoint(codePoint);
            if (glyph != 0 && (!unicodeByOriginalGlyph.TryGetValue(glyph, out int existing) || (existing != 0x20 && codePoint == 0x20)))
            {
                unicodeByOriginalGlyph[glyph] = codePoint;
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        return CreateFromOriginalGlyphs(font, unicodeByOriginalGlyph, cancellationToken);
    }

    public static PdfEmbeddedFont Merge(IEnumerable<PdfEmbeddedFont> fonts, CancellationToken cancellationToken)
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

        // repeated slides/DOCX runs commonly reference the very same subset
        // instance. Rebuilding its dictionary, re-hashing the font program, and
        // re-subsetting on every merge is pure duplicate work: return it directly.
        bool allIdentical = true;
        for (int i = 1; i < items.Length; i++)
        {
            if (!ReferenceEquals(items[i], items[0]))
            {
                allIdentical = false;
                break;
            }
        }

        if (allIdentical)
        {
            return items[0];
        }

        // Distinct instances can still carry provably identical remaps: same font
        // program plus equal glyph mappings subset to identical bytes. Merging those
        // must not re-subset either. Anything else takes the union path below; equal
        // ResourceKeys alone are not sufficient to conclude the remaps agree.
        if (HasIdenticalMappings(items))
        {
            return items[0];
        }

        var unicodeByOriginalGlyph = new SortedDictionary<ushort, int>();
        foreach (PdfEmbeddedFont font in items)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach ((ushort glyph, int codePoint) in font.UnicodeByOriginalGlyph)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!unicodeByOriginalGlyph.TryGetValue(glyph, out int existingMerge) || (existingMerge != 0x20 && codePoint == 0x20))
                {
                    unicodeByOriginalGlyph[glyph] = codePoint;
                }
            }
        }

        return CreateFromOriginalGlyphs(items[0].Font, unicodeByOriginalGlyph, cancellationToken);
    }

    internal static PdfEmbeddedFont CreateForGlyphMappings(OpenTypeFont font,
        IEnumerable<(ushort Glyph, int Unicode)> mappings, CancellationToken token)
    {
        var owned = new SortedDictionary<ushort, int>();
        foreach ((ushort glyph, int unicode) in mappings)
        {
            token.ThrowIfCancellationRequested();
            if (glyph == 0 || glyph >= font.GlyphCount || !Rune.IsValid(unicode))
                throw new ArgumentException("Explicit font mappings require a valid source glyph and Unicode scalar.", nameof(mappings));
            if (!owned.TryGetValue(glyph, out int existing) || existing != 0x20 && unicode == 0x20) { owned[glyph] = unicode; }
        }
        token.ThrowIfCancellationRequested();
        return CreateFromOriginalGlyphs(font, owned, token);
    }

    private static PdfEmbeddedFont CreateFromOriginalGlyphs(
        OpenTypeFont font,
        IReadOnlyDictionary<ushort, int> unicodeByOriginalGlyph,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!OpenTypeFontSubsetter.TryCreate(font, unicodeByOriginalGlyph, cancellationToken, out OpenTypeFontSubset? subset, out string? subsetFailure))
        {
            // Whole-TrueType fallback is only valid for TrueType outlines: embedding
            // another format (e.g. CFF) as FontFile2/CIDFontType2 would silently
            // produce a mismatched font dictionary.
            //
            // P04 embed-call-site audit: every Create/Merge caller either gates on
            // HasTrueTypeOutlines with a FONT_UNSUPPORTED_OUTLINES diagnostic
            // (DOCX substitution, document-fallback, and per-character guards;
            // PPTX substitution and emission skip) or re-embeds from an existing
            // PdfEmbeddedFont.Font, which is TrueType by construction here (subset
            // success and the whole-font fallback above both imply it). This throw
            // is therefore unreachable through diagnosed conversion paths and stays
            // as a loud backstop. Pinned by
            // FontFormatTests.CffKindFontRefusesCidFontType2Embedding.
            if (!font.HasTrueTypeOutlines)
            {
                throw new InvalidDataException("Cannot embed font '" + font.FamilyName + "' without TrueType outlines as CIDFontType2: " + subsetFailure);
            }

            subset = null;
        }
        var sortedGlyphs = new SortedDictionary<ushort, int>();
        foreach ((ushort glyph, int codePoint) in unicodeByOriginalGlyph)
        {
            sortedGlyphs[glyph] = codePoint;
        }

        // the codepoint-set hash was computed twice per subset (once for the
        // base-font tag, once in the constructor). Compute it once here and share it.
        string setHash = ComputeResourceMappingHash(font, sortedGlyphs);
        string baseFontName = CreateBaseFontName(font, setHash, subset is not null);
        var embedded = new PdfEmbeddedFont(font, baseFontName, setHash, sortedGlyphs, subset);
        return embedded;
    }

    private static bool HasIdenticalMappings(PdfEmbeddedFont[] items)
    {
        IReadOnlyDictionary<ushort, int> first = items[0].UnicodeByOriginalGlyph;
        for (int i = 1; i < items.Length; i++)
        {
            IReadOnlyDictionary<ushort, int> other = items[i].UnicodeByOriginalGlyph;
            if (!ReferenceEquals(items[i].Font, items[0].Font) || other.Count != first.Count)
            {
                return false;
            }

            foreach ((ushort glyph, int codePoint) in first)
            {
                if (!other.TryGetValue(glyph, out int otherCodePoint) || otherCodePoint != codePoint)
                {
                    return false;
                }
            }
        }

        return true;
    }

    internal static string CreateBaseFontName(OpenTypeFont font, string codepointSetHash, bool isSubset)
    {
        string sanitizedFamily = SanitizeName(font.FamilyName);
        if (!isSubset)
        {
            return sanitizedFamily;
        }

        string fontHash = Convert.ToHexString(SHA256.HashData(font.Bytes.Span)).Substring(0, 8);
        string tag = CreateSubsetTag(fontHash, codepointSetHash);
        return tag + "+" + sanitizedFamily;
    }

    internal static string ComputeCodepointSetHash(IReadOnlyDictionary<ushort, int> unicodeByOriginalGlyph)
    {
        int[] codePoints = unicodeByOriginalGlyph.Values.Distinct().OrderBy(codePoint => codePoint).ToArray();
        byte[] bytes = new byte[codePoints.Length * 4];
        for (int i = 0; i < codePoints.Length; i++)
        {
            bytes[i * 4] = (byte)codePoints[i];
            bytes[i * 4 + 1] = (byte)(codePoints[i] >> 8);
            bytes[i * 4 + 2] = (byte)(codePoints[i] >> 16);
            bytes[i * 4 + 3] = (byte)(codePoints[i] >> 24);
        }

        return Convert.ToHexString(SHA256.HashData(bytes)).Substring(0, 12);
    }

    private static string ComputeResourceMappingHash(OpenTypeFont font, SortedDictionary<ushort, int> mappings)
    {
        string canonical = ComputeCodepointSetHash(mappings);
        if (mappings.All(pair => font.MapCodePoint(pair.Value) == pair.Key)) { return canonical; }
        using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        Span<byte> record = stackalloc byte[6];
        foreach ((ushort glyph, int unicode) in mappings)
        {
            System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(record, glyph);
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(record.Slice(2), unicode);
            hash.AppendData(record);
        }
        return canonical + "G" + Convert.ToHexString(hash.GetHashAndReset()).Substring(0, 12);
    }

    internal static string CreateSubsetTag(string fontHash, string codepointSetHash)
    {
        byte[] tagHash = SHA256.HashData(Encoding.ASCII.GetBytes(fontHash + "|" + codepointSetHash));
        var builder = new StringBuilder(6);
        for (int i = 0; i < 6; i++)
        {
            builder.Append((char)('A' + (tagHash[i] % 26)));
        }

        return builder.ToString();
    }

    private static readonly string ToUnicodePrefix = string.Join(Environment.NewLine,
        "/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
        "/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> def",
        "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def", "1 begincodespacerange",
        "<0000> <FFFF>", "endcodespacerange", "");
    private static readonly string ToUnicodeSuffix = string.Join(Environment.NewLine,
        "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end", "");

    public int MeasureToUnicodeCMap(Memory<byte> scratch, CancellationToken cancellationToken) =>
        EncodeToUnicodeCMapAsync(null, scratch, cancellationToken).GetAwaiter().GetResult();

    public async ValueTask WriteToUnicodeCMapAsync(PdfObjectWriter writer, CancellationToken cancellationToken) =>
        _ = await EncodeToUnicodeCMapAsync(writer, writer.Scratch, cancellationToken).ConfigureAwait(false);

    private async ValueTask<int> EncodeToUnicodeCMapAsync(PdfObjectWriter? writer, Memory<byte> scratch, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (writer is not null) await writer.WriteAsciiAsync(ToUnicodePrefix).ConfigureAwait(false);
        int length = ToUnicodePrefix.Length + ToUnicodeSuffix.Length;
        // One bounded encoder measures and emits the same immutable map. Each
        // block contains at most 100 mappings; even full int hex values fit 4KB.
        // Keep the established platform newlines and LF-only block declaration.
        Memory<byte> block = scratch[..4096];
        string newline = Environment.NewLine;
        int remaining = UnicodeByCid.Count, inBlock = 0, used = 0, blockCount = 0;
        foreach ((ushort cid, int codePoint) in UnicodeByCid)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (inBlock == 0)
            {
                blockCount = Math.Min(100, remaining);
                _ = blockCount.TryFormat(block.Span, out used, provider: CultureInfo.InvariantCulture);
                " beginbfchar\n"u8.CopyTo(block.Span[used..]);
                used += " beginbfchar\n"u8.Length;
            }
            block.Span[used++] = (byte)'<';
            _ = cid.TryFormat(block.Span[used..], out int written, "X4", CultureInfo.InvariantCulture);
            used += written;
            "> <"u8.CopyTo(block.Span[used..]); used += 3;
            if (codePoint <= 0xFFFF)
            {
                _ = codePoint.TryFormat(block.Span[used..], out written, "X4", CultureInfo.InvariantCulture);
                used += written;
            }
            else
            {
                int scalar = codePoint - 0x10000;
                int high = 0xD800 + (scalar >> 10), low = 0xDC00 + (scalar & 0x3FF);
                _ = high.TryFormat(block.Span[used..], out written, "X4", CultureInfo.InvariantCulture);
                used += written;
                _ = low.TryFormat(block.Span[used..], out written, "X4", CultureInfo.InvariantCulture);
                used += written;
            }
            block.Span[used++] = (byte)'>';
            foreach (char ch in newline) block.Span[used++] = (byte)ch;
            if (++inBlock == blockCount)
            {
                "endbfchar"u8.CopyTo(block.Span[used..]); used += "endbfchar"u8.Length;
                foreach (char ch in newline) block.Span[used++] = (byte)ch;
                length = checked(length + used);
                if (writer is not null) await writer.WriteRawBytesAsync(block[..used]).ConfigureAwait(false);
                remaining -= blockCount;
                inBlock = 0;
            }
        }
        if (writer is not null) await writer.WriteAsciiAsync(ToUnicodeSuffix).ConfigureAwait(false);
        return length;
    }

    public async ValueTask WriteWidthArrayAsync(PdfObjectWriter writer, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Memory<byte> block = writer.Scratch[..512];
        int used = 1;
        block.Span[0] = (byte)'[';
        bool first = true;
        // Source glyphs and the subset's complete glyph closure are sorted.
        // Subset CIDs therefore preserve source order; whole-font CIDs are identity.
        foreach (ushort originalGlyph in UnicodeByOriginalGlyph.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryGetCid(originalGlyph, out ushort cid))
            {
                continue;
            }

            if (block.Length - used < 32)
            {
                await writer.WriteRawBytesAsync(block[..used]).ConfigureAwait(false);
                used = 0;
            }
            if (!first)
            {
                block.Span[used++] = (byte)' ';
            }
            first = false;
            int width = (int)Math.Round(Font.GetAdvanceWidth(originalGlyph) * 1000d / Font.UnitsPerEm);
            _ = cid.TryFormat(block.Span[used..], out int written, provider: CultureInfo.InvariantCulture);
            used += written;
            " ["u8.CopyTo(block.Span[used..]); used += 2;
            _ = width.TryFormat(block.Span[used..], out written, provider: CultureInfo.InvariantCulture);
            used += written;
            block.Span[used++] = (byte)']';
        }
        block.Span[used++] = (byte)']';
        await writer.WriteRawBytesAsync(block[..used]).ConfigureAwait(false);
    }

    public string EncodeGlyphHex(string text)
    {
        var builder = new StringBuilder(text.Length * 4);
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = Font.MapCodePoint(rune.Value);
            if (glyph != 0 && TryGetCid(glyph, out ushort cid))
            {
                builder.Append(cid.ToString("X4", CultureInfo.InvariantCulture));
            }
        }

        return builder.ToString();
    }

    public bool TryGetEncodedCid(ushort originalGlyph, out ushort cid)
    {
        return TryGetCid(originalGlyph, out cid);
    }

    public string? EncodeGlyphPositioningArray(string text, double characterSpacingPoints, double fontSize, bool forcePositioningArray, bool kerningEnabled, double? nominalAdvanceFontSize = null)
    {
        var glyphs = new List<EncodedGlyph>();
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = Font.MapCodePoint(rune.Value);
            if (glyph != 0 && TryGetCid(glyph, out ushort cid))
            {
                glyphs.Add(new EncodedGlyph(glyph, cid));
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
                if (nominalAdvanceFontSize is { } nominalSize && fontSize > 0d)
                {
                    double advanceEm = Font.GetAdvanceWidth(glyphs[i - 1].OriginalGlyph) / (double)Font.UnitsPerEm;
                    // Keep the physical glyph size; adjust only its advance before
                    // the following glyph. The final natural glyph width stays intact.
                    adjustment += Math.Round(advanceEm * 1000d) - advanceEm * nominalSize * 1000d / fontSize;
                }
                short kerning = kerningEnabled ? Font.GetKerning(glyphs[i - 1].OriginalGlyph, glyphs[i].OriginalGlyph) : (short)0;
                if (kerning != 0)
                {
                    double pairScale = nominalAdvanceFontSize is { } pairSize && fontSize > 0d ? pairSize / fontSize : 1d;
                    adjustment += -kerning * 1000d / Font.UnitsPerEm * pairScale;
                }

                if (Math.Abs(adjustment) > 0.001d)
                {
                    AppendGlyphChunk(builder, glyphChunk);
                    builder.Append(' ').Append(adjustment.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ');
                    hasPositioning = true;
                }
            }

            glyphChunk.Append(glyphs[i].Cid.ToString("X4", CultureInfo.InvariantCulture));
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

    private IReadOnlyDictionary<ushort, int> BuildUnicodeByCid()
    {
        var unicodeByCid = new SortedDictionary<ushort, int>();
        foreach ((ushort originalGlyph, int codePoint) in UnicodeByOriginalGlyph)
        {
            if (TryGetCid(originalGlyph, out ushort cid))
            {
                // When subsetting merges identically-outlined glyphs (e.g. space and NBSP) into one CID,
                // report plain space: Office also extracts spaces there (NBSP probe 2026-09-06: 15 source NBSPs,
                // Word reference decodes 245 spaces and 0 NBSP while we decoded 315 NBSPs).
                if (!unicodeByCid.TryGetValue(cid, out int existing) || (existing != 0x20 && codePoint == 0x20))
                {
                    unicodeByCid[cid] = codePoint;
                }
            }
        }

        return unicodeByCid;
    }

    private bool TryGetCid(ushort originalGlyph, out ushort cid)
    {
        if (cidByOriginalGlyph.TryGetValue(originalGlyph, out cid))
        {
            return true;
        }

        if (UsesSubsetFontProgram)
        {
            cid = 0;
            return false;
        }

        cid = originalGlyph;
        return true;
    }

    public double MeasureNominalAdvanceAdjustmentPoints(string text, double fontSize, double nominalFontSize)
    {
        double adjustment = 0d;
        ushort previous = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = Font.MapCodePoint(rune.Value);
            if (glyph == 0 || !TryGetCid(glyph, out _)) continue;
            if (previous != 0)
            {
                double advanceEm = Font.GetAdvanceWidth(previous) / (double)Font.UnitsPerEm;
                adjustment += advanceEm * nominalFontSize - Math.Round(advanceEm * 1000d) * fontSize / 1000d;
            }
            previous = glyph;
        }
        return adjustment;
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
            builder.Append(char.IsLetterOrDigit(c) ? c : c == '+' ? '+' : '-');
        }

        return builder.Length == 0 ? "Font" : builder.ToString();
    }

    private readonly record struct EncodedGlyph(ushort OriginalGlyph, ushort Cid);
}
