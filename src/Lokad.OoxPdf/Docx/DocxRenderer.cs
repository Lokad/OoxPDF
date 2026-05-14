using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxRenderer
{
    public IReadOnlyList<PdfPage> RenderBlankPages(DocxDocument document)
    {
        return [new PdfPage(document.PageWidthPoints, document.PageHeightPoints)];
    }
}
