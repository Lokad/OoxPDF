using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Fonts;

internal sealed class PresentationFontResolver
{
    private readonly IFontResolver primary;
    private readonly WindowsFontResolver windowsCatalog;
    private readonly IFontCatalog fontCatalog;
    // T01: program identity is an ordinal (StableId, FaceIndex) tuple. StableIds are
    // opaque resolver-scoped identities and must not be case-normalized (matching the
    // DOCX font-plan keys and the IFontProgramSource ownership contract); family-name
    // case-insensitivity lives one layer up in FontRequestKeyComparer.
    private readonly Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> openTypeFonts = new();
    private readonly Dictionary<string, SubsetEntry> subsets = new(StringComparer.Ordinal);

    public PresentationFontResolver(IFontResolver? primary)
    {
        this.primary = primary ?? new WindowsFontResolver();
        windowsCatalog = this.primary as WindowsFontResolver ?? new WindowsFontResolver();
        fontCatalog = this.primary as IFontCatalog ?? windowsCatalog;
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        return primary.Resolve(request);
    }

    public FontFaceResolution ResolvePresentationTextFace(FontRequest request)
    {
        return primary is WindowsFontResolver windows
            ? windows.ResolvePresentationTextFace(request)
            : primary.Resolve(request);
    }

    public (FontFaceResolution Resolution, OpenTypeFont Font)? ResolvePresentationOpenTypeFont(FontRequest request, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        FontFaceResolution resolution = ResolvePresentationTextFace(request);
        OpenTypeFont? font = GetOrLoadOpenTypeFont(resolution, cancellationToken);
        return font is null ? null : (resolution, font);
    }

    public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return fontCatalog.GetDiscoveredFonts();
    }

    internal OpenTypeFont? GetOrLoadOpenTypeFont(FontFaceResolution resolution, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var key = (resolution.Source.StableId, resolution.FontFaceIndex);
        if (openTypeFonts.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        // PLAN Q01: conversion-wide cumulative charge before font parsing allocates.
        OoxConversionBudget.Current?.ChargeFontWork(1);
        cached = FontProgramLoader.Load(resolution, cancellationToken);
        openTypeFonts[key] = cached;
        return cached;
    }

    // Conversion-scoped identical subsets: repeated slides embed the same glyph
    // sets, so subsetting runs once per (program, codepoint set) instead of once
    // per slide. The subsetter is a pure function of those inputs, so shared
    // instances keep identical bytes, CID maps, and resource keys (G04).
    internal PdfEmbeddedFont GetOrCreateSubset(FontFaceResolution resolution, OpenTypeFont font, IReadOnlyList<int> codePoints, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // T01: the 64-bit set hash is a lookup shortcut, not an identity proof. Stored
        // sets are compared on every hit so a hash collision can never corrupt output;
        // a mismatch falls back to a fresh uncached subset instead of poisoning the cache.
        string key = resolution.Source.StableId + "\u001f" + resolution.FontFaceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\u001f" + HashCodePointSet(codePoints);
        int[] sorted = SortDeduplicate(codePoints);
        if (subsets.TryGetValue(key, out SubsetEntry? cached) && cached.CodePoints.AsSpan().SequenceEqual(sorted))
        {
            return cached.Subset;
        }

        // PLAN Q01: conversion-wide cumulative charge before subsetting allocates.
        OoxConversionBudget.Current?.ChargeFontWork(1);
        PdfEmbeddedFont created = PdfEmbeddedFont.Create(font, codePoints, cancellationToken);
        if (cached is null)
        {
            subsets[key] = new SubsetEntry(created, sorted);
        }

        return created;
    }

    private sealed record SubsetEntry(PdfEmbeddedFont Subset, int[] CodePoints);

    private static int[] SortDeduplicate(IReadOnlyList<int> codePoints)
    {
        int[] sorted = codePoints.ToArray();
        Array.Sort(sorted);
        int count = 0;
        int previous = 0;
        bool first = true;
        foreach (int codePoint in sorted)
        {
            if (!first && codePoint == previous)
            {
                continue;
            }

            first = false;
            previous = codePoint;
            sorted[count++] = codePoint;
        }

        Array.Resize(ref sorted, count);
        return sorted;
    }

    private static string HashCodePointSet(IReadOnlyList<int> codePoints)
    {
        int[] sorted = SortDeduplicate(codePoints);
        ulong hash = 14695981039346656037ul;
        foreach (int codePoint in sorted)
        {
            hash ^= (uint)codePoint;
            hash *= 1099511628211ul;
        }

        return hash.ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
    }
}
