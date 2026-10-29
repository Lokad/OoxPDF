using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

// Empty primary and discovery sources make standard-font fallback independent
// of installed host fonts.
internal sealed class MissingTestFontResolver : IFontResolver, IFontCatalog
{
    private static readonly MemoryFontProgramSource Source = new("test:no-font", ReadOnlyMemory<byte>.Empty);

    public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName,
        "Missing", new FontStyleKey(request.Bold, request.Italic), Source, IsFallback: true);

    public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts() => [];
}
