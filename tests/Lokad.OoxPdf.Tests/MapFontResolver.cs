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

    public FontResolution Resolve(FontRequest request)
    {
        return availableFamilies.Contains(request.FamilyName)
            ? new FontResolution(request.FamilyName, request.FamilyName + ".ttf", IsFallback: false, request.Bold, request.Italic)
            : new FontResolution(fallbackFamily, fallbackFamily + ".ttf", IsFallback: true, request.Bold, request.Italic);
    }
}
