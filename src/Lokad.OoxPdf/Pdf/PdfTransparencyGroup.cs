namespace Lokad.OoxPdf.Pdf;

// SVG vector isolation. Names are local to each Form, and identity belongs to
// this conversion rather than a global cache or a content digest.
internal sealed record PdfTransparencyGroupResource(string ResourceName, PdfTransparencyGroup Group);

internal sealed class PdfTransparencyGroup(
    PdfRectangle bounds,
    string content,
    IReadOnlyList<PdfExtGStateResource> extGStates,
    IReadOnlyList<PdfShadingResource> shadings,
    IReadOnlyList<PdfTransparencyGroupResource> groups,
    IReadOnlyList<PdfShadingPatternResource>? shadingPatterns = null)
{
    internal const int MaxDepth = 32;

    public PdfRectangle Bounds { get; } = bounds;
    public string Content { get; } = content;
    public IReadOnlyList<PdfExtGStateResource> ExtGStates { get; } = extGStates.ToArray();
    public IReadOnlyList<PdfShadingResource> Shadings { get; } = shadings.ToArray();
    public IReadOnlyList<PdfTransparencyGroupResource> Groups { get; } = groups.ToArray();
    public IReadOnlyList<PdfShadingPatternResource> ShadingPatterns { get; } = shadingPatterns?.ToArray() ?? [];
}
