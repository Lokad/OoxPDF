namespace Lokad.OoxPdf.Pdf;

// SVG vector isolation. Names are local to each Form, and identity belongs to
// this conversion rather than a global cache or a content digest.
internal sealed record PdfTransparencyGroupResource(string ResourceName, PdfTransparencyGroup Group);

internal sealed class PdfTransparencyGroup
{
    internal const int MaxDepth = 32;

    public PdfTransparencyGroup(
        PdfRectangle bounds,
        string content,
        IReadOnlyList<PdfExtGStateResource> extGStates,
        IReadOnlyList<PdfShadingResource> shadings,
        IReadOnlyList<PdfTransparencyGroupResource> groups,
        IReadOnlyList<PdfShadingPatternResource>? shadingPatterns = null)
    {
        Bounds = bounds;
        Content = content;
        ExtGStates = extGStates.ToArray();
        Shadings = shadings.ToArray();
        Groups = groups.ToArray();
        ShadingPatterns = shadingPatterns?.ToArray() ?? [];
    }

    // Staging transfers the new child bindings and shares the existing snapshots.
    // None of the shared resources retains the original Form or its payload.
    internal PdfTransparencyGroup(PdfTransparencyGroup source, PdfTransparencyGroupResource[] stagedChildren)
    {
        Bounds = source.Bounds;
        Content = string.Empty;
        ExtGStates = source.ExtGStates;
        Shadings = source.Shadings;
        Groups = stagedChildren;
        ShadingPatterns = source.ShadingPatterns;
    }

    public PdfRectangle Bounds { get; }
    public string Content { get; }
    public IReadOnlyList<PdfExtGStateResource> ExtGStates { get; }
    public IReadOnlyList<PdfShadingResource> Shadings { get; }
    public IReadOnlyList<PdfTransparencyGroupResource> Groups { get; }
    public IReadOnlyList<PdfShadingPatternResource> ShadingPatterns { get; }
}
