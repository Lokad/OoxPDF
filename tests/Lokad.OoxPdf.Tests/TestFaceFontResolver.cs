using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

// RV01: deterministic embeddable test face serving the synthetic TrueType for every
// requested family, so font-budget and admission tests assert identical behavior with
// or without installed fonts (no environmental skips).
internal sealed class TestFaceFontResolver : IFontResolver
{
    private static readonly byte[] FaceBytes = TestFontBuilder.CreateTestFont();

    public FontFaceResolution Resolve(FontRequest request)
    {
        return new FontFaceResolution(
            request.FamilyName,
            "TestFace",
            new FontStyleKey(request.Bold, request.Italic),
            new MemoryFontProgramSource("test:face", FaceBytes),
            IsFallback: false);
    }
}
