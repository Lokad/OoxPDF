namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxDocument(string PresentationPartName, IReadOnlyList<PptxSlide> Slides, double SlideWidthPoints, double SlideHeightPoints)
{
    public long FirstSlideNumber { get; init; } = 1;
}

internal sealed record PptxSlide(string PartName, int Index)
{
    // Index addresses emitted pages. SourceIndex includes hidden/skipped slides;
    // manually constructed documents retain their existing index-based default.
    public int? SourceIndex { get; init; }
}
