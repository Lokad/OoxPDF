namespace Lokad.OoxPdf.Fonts;

internal readonly record struct OpenTypeMathGlyphVariant(ushort Glyph, ushort Advance);

internal sealed partial class OpenTypeFont
{
    // Select a ready-made vertical variant from the owning immutable font.
    // Assemblies and device corrections remain separate, unsupported features.
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
            minimumHeight > ushort.MaxValue || !tables.TryGetValue("MATH", out TableRecord math)) { return false; }
        ReadOnlySpan<byte> table = bytes.AsSpan(math.Offset, math.Length);
        if (table.Length < 10 || U16(table, 0) != 1 || U16(table, 2) != 0) { return false; }
        int variantsOffset = U16(table, 8);
        if (variantsOffset < 10 || table.Length - variantsOffset < 10) { return false; }
        ReadOnlySpan<byte> variants = table.Slice(variantsOffset);
        int coverageOffset = U16(variants, 2), verticalCount = U16(variants, 6), horizontalCount = U16(variants, 8);
        int offsetsEnd = 10 + (verticalCount + horizontalCount) * 2;
        if (verticalCount == 0 || variants.Length < offsetsEnd || coverageOffset < offsetsEnd ||
            variants.Length - coverageOffset < 4 ||
            !TryGetCoverageIndex(variants.Slice(coverageOffset), glyph, out int index) || index < 0 || index >= verticalCount)
            { return false; }
        int constructionOffset = U16(variants, 10 + index * 2);
        if (constructionOffset < offsetsEnd || variants.Length - constructionOffset < 4) { return false; }
        ReadOnlySpan<byte> construction = variants.Slice(constructionOffset);
        int count = U16(construction, 2);
        if (construction.Length - 4 < count * 4) { return false; }
        if (count > 0) { OoxConversionBudget.Current?.ChargeFontWork(1); }
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
        // beyond every ready-made variant still needs an unsupported assembly.
        if (found && adequate) { variant = selected; return true; }
        return false;
    }
}
