using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal sealed class MapFontResolver : IFontResolver
{
    private readonly HashSet<string> availableFamilies;
    private readonly string fallbackFamily;

    public MapFontResolver(IEnumerable<string> availableFamilies, string fallbackFamily)
    {
        this.availableFamilies = new HashSet<string>(availableFamilies, StringComparer.OrdinalIgnoreCase);
        this.fallbackFamily = fallbackFamily;
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        return availableFamilies.Contains(request.FamilyName)
            ? CreateResolution(request.FamilyName, request.FamilyName, request, isFallback: false)
            : CreateResolution(request.FamilyName, fallbackFamily, request, isFallback: true);
    }

    private static FontFaceResolution CreateResolution(string requestedFamily, string resolvedFamily, FontRequest request, bool isFallback)
    {
        return new FontFaceResolution(
            requestedFamily,
            resolvedFamily,
            new FontStyleKey(request.Bold, request.Italic),
            new MemoryFontProgramSource("test:" + resolvedFamily, ReadOnlyMemory<byte>.Empty),
            isFallback);
    }
}
