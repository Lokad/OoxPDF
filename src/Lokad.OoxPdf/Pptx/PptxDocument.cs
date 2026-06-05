namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxDocument(string PresentationPartName, IReadOnlyList<PptxSlide> Slides, double SlideWidthPoints, double SlideHeightPoints);

internal sealed record PptxSlide(string PartName, int Index)
{
}
