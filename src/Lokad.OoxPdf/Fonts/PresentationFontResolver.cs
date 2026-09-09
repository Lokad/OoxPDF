using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Fonts;

internal sealed class PresentationFontResolver
{
    private readonly IFontResolver primary;
    private readonly WindowsFontResolver windowsCatalog;
    private readonly IFontCatalog fontCatalog;
    private readonly Dictionary<string, OpenTypeFont?> openTypeFonts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, PdfEmbeddedFont> subsets = new(StringComparer.Ordinal);

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
        string key = resolution.Source.StableId + "\u001f" + resolution.FontFaceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (openTypeFonts.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

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
        string key = resolution.Source.StableId + "\u001f" + resolution.FontFaceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\u001f" + HashCodePointSet(codePoints);
        if (subsets.TryGetValue(key, out PdfEmbeddedFont? cached))
        {
            return cached;
        }

        cached = PdfEmbeddedFont.Create(font, codePoints, cancellationToken);
        subsets[key] = cached;
        return cached;
    }

    private static string HashCodePointSet(IReadOnlyList<int> codePoints)
    {
        int[] sorted = codePoints.ToArray();
        Array.Sort(sorted);
        ulong hash = 14695981039346656037ul;
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
            hash ^= (uint)codePoint;
            hash *= 1099511628211ul;
        }

        return hash.ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
    }
}
