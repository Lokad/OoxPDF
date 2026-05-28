using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal sealed class CountingFontResolver : IFontResolver
{
    private readonly WindowsFontResolver fallback = new();

    public int ResolveCalls { get; private set; }

    public FontResolution Resolve(FontRequest request)
    {
        ResolveCalls++;
        return fallback.Resolve(request);
    }
}
