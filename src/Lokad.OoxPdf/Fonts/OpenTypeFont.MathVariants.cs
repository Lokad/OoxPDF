namespace Lokad.OoxPdf.Fonts;

internal readonly record struct OpenTypeMathGlyphVariant(ushort Glyph, ushort Advance);
internal readonly record struct OpenTypeMathGlyphPart(ushort Glyph, ushort StartConnectorLength,
    ushort EndConnectorLength, ushort FullAdvance, bool IsExtender);
internal sealed record OpenTypeMathGlyphAssembly(short ItalicCorrection, ushort MinimumConnectorOverlap,
    IReadOnlyList<OpenTypeMathGlyphPart> Parts);

internal sealed partial class OpenTypeFont
{
    // Select a ready-made vertical variant from the owning immutable font.
    // Assembly decoding is separate; selection never publishes an inadequate
    // ready-made variant. Device corrections remain outside design-unit metrics.
    public bool TryGetMathVerticalVariant(ushort glyph, double minimumHeight, CancellationToken token,
        out OpenTypeMathGlyphVariant variant) => TrySelectMathVerticalVariant(glyph, minimumHeight, token, nearest: false, out variant);

    public bool TryGetNearestMathVerticalVariant(ushort glyph, double height, CancellationToken token,
        out OpenTypeMathGlyphVariant variant) => TrySelectMathVerticalVariant(glyph, height, token, nearest: true, out variant);

    private bool TrySelectMathVerticalVariant(ushort glyph, double minimumHeight, CancellationToken token,
        bool nearest, out OpenTypeMathGlyphVariant variant)
    {
        variant = default;
        token.ThrowIfCancellationRequested();
        if (glyph == 0 || glyph >= GlyphCount || !double.IsFinite(minimumHeight) || minimumHeight < 0d ||
            minimumHeight > ushort.MaxValue || !TryGetMathVerticalConstruction(glyph, out _, out var construction, out _)) { return false; }
        int count = U16(construction, 2);
        bool found = false, adequate = false;
        OpenTypeMathGlyphVariant selected = default;
        for (int i = 0; i < count; i++)
        {
            token.ThrowIfCancellationRequested();
            ushort candidate = U16(construction, 4 + i * 4), advance = U16(construction, 6 + i * 4);
            if (candidate == 0 || candidate >= GlyphCount || advance == 0) { return false; }
            adequate |= advance >= minimumHeight;
            bool take = nearest
                ? !found || Math.Abs(advance - minimumHeight) < Math.Abs(selected.Advance - minimumHeight) ||
                    Math.Abs(advance - minimumHeight) == Math.Abs(selected.Advance - minimumHeight) && advance > selected.Advance
                : advance >= minimumHeight && (!found || advance < selected.Advance);
            if (take)
            { selected = new(candidate, advance); found = true; }
        }
        // A nearer smaller shape is useful for delimiter fitting, but a target
        // beyond every ready-made variant still needs an assembly.
        if (found && adequate) { variant = selected; return true; }
        return false;
    }

    public bool TryGetMathVerticalAssembly(ushort glyph, CancellationToken token, out OpenTypeMathGlyphAssembly? assembly)
    {
        assembly = null;
        token.ThrowIfCancellationRequested();
        if (!TryGetMathVerticalConstruction(glyph, out var variants, out var construction, out int constructionStart)) return false;
        int offset = U16(construction, 0), recordsEnd = 4 + U16(construction, 2) * 4;
        if (offset == 0 || offset < recordsEnd || construction.Length - offset < 6) return false;
        ReadOnlySpan<byte> source = construction.Slice(offset);
        int count = U16(source, 4);
        if (count == 0 || source.Length - 6 < count * 10) return false;
        for (int index = 0; index < U16(construction, 2); index++)
        {
            token.ThrowIfCancellationRequested();
            ushort variantGlyph = U16(construction, 4 + index * 4), advance = U16(construction, 6 + index * 4);
            if (variantGlyph == 0 || variantGlyph >= GlyphCount || advance == 0) return false;
        }
        // Validate the complete source before publishing its immutable view.
        for (int index = 0; index < count; index++)
        {
            token.ThrowIfCancellationRequested();
            int at = 6 + index * 10;
            ushort partGlyph = U16(source, at), advance = U16(source, at + 6), flags = U16(source, at + 8);
            if (partGlyph == 0 || partGlyph >= GlyphCount || advance == 0 || (flags & ~1) != 0) return false;
        }
        // Like MathConstants, retain the signed design-unit value without
        // following optional device-table offsets. The private font already owns
        // these records; decoding needs no second retained copy or cache.
        assembly = new(System.Buffers.Binary.BinaryPrimitives.ReadInt16BigEndian(source),
            U16(variants, 0), new MathGlyphParts(bytes.AsMemory(checked(constructionStart + offset + 6), count * 10)));
        return true;
    }

    private bool TryGetMathVerticalConstruction(ushort glyph, out ReadOnlySpan<byte> variants,
        out ReadOnlySpan<byte> construction, out int constructionStart)
    {
        variants = construction = default;
        constructionStart = 0;
        if (glyph == 0 || glyph >= GlyphCount || !tables.TryGetValue("MATH", out TableRecord math)) return false;
        ReadOnlySpan<byte> table = bytes.AsSpan(math.Offset, math.Length);
        if (table.Length < 10 || U16(table, 0) != 1 || U16(table, 2) != 0) { return false; }
        int variantsOffset = U16(table, 8);
        if (variantsOffset < 10 || table.Length - variantsOffset < 10) { return false; }
        variants = table.Slice(variantsOffset);
        int coverageOffset = U16(variants, 2), verticalCount = U16(variants, 6), horizontalCount = U16(variants, 8);
        int offsetsEnd = 10 + (verticalCount + horizontalCount) * 2;
        if (verticalCount == 0 || variants.Length < offsetsEnd || coverageOffset < offsetsEnd ||
            variants.Length - coverageOffset < 4 ||
            !TryGetCoverageIndex(variants.Slice(coverageOffset), glyph, out int index) || index < 0 || index >= verticalCount)
            { return false; }
        int constructionOffset = U16(variants, 10 + index * 2);
        if (constructionOffset < offsetsEnd || variants.Length - constructionOffset < 4) { return false; }
        construction = variants.Slice(constructionOffset);
        constructionStart = checked(math.Offset + variantsOffset + constructionOffset);
        int count = U16(construction, 2);
        return construction.Length - 4 >= count * 4;
    }

    // The view retains the private immutable SFNT, already owned by the font.
    private sealed class MathGlyphParts(ReadOnlyMemory<byte> source) : IReadOnlyList<OpenTypeMathGlyphPart>
    {
        public int Count => source.Length / 10;
        public OpenTypeMathGlyphPart this[int index]
        {
            get
            {
                if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
                var part = source.Span.Slice(index * 10, 10);
                return new(U16(part, 0), U16(part, 2), U16(part, 4), U16(part, 6), (U16(part, 8) & 1) != 0);
            }
        }
        public IEnumerator<OpenTypeMathGlyphPart> GetEnumerator()
        {
            for (int index = 0; index < Count; index++) yield return this[index];
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
