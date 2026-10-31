using Lokad.OoxPdf.Pptx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static DocxParagraph? ParagraphPaintSource(DocxLayoutItem item) => item switch
    {
        DocxTextLineLayout line => line.SourceParagraph,
        DocxEquationLayout equation => equation.SourceParagraph,
        DocxInlineImageLayout image => image.SourceParagraph,
        _ => null
    };

    private static RgbColor? ResolveParagraphBackground(DocxLayoutItem item, RgbColor? background)
    {
        return item.ParagraphShadingFrame is not null && ParagraphPaintSource(item)?.Shading is { } shading
            ? ResolveShadingBackground(shading.FillHex, shading.Value, shading.ColorHex, background) : background;
    }

    private static void RenderParagraphBackgrounds(IEnumerable<DocxLayoutItem> items,
        PdfGraphicsBuilder graphics, DocxMarkupContext markupContext, CancellationToken token)
    {
        // Paint behind the complete container's text before any glyphs. Exact or
        // overlapping line boxes cannot erase a preceding line's visible ink.
        foreach (DocxLayoutItem item in items)
        {
            token.ThrowIfCancellationRequested();
            if (item.ParagraphShadingFrame is not { } frame || ParagraphPaintSource(item)?.Shading is not { } shading)
                continue;
            var (x, baseline) = item switch
            {
                DocxTextLineLayout line => (line.X, line.BaselineY),
                DocxEquationLayout equation => (equation.X, equation.BaselineY),
                DocxInlineImageLayout image => (image.X, image.Y),
                _ => (0d, 0d)
            };
            double yOffset = item is DocxInlineImageLayout { StaticStoryUsesPhysicalY: true }
                ? 0d : ResolveTextEmissionBaselineOffset(markupContext);
            RenderShadingFill(shading.FillHex, shading.Value, shading.ColorHex, graphics,
                x + frame.OffsetX + ResolveTextEmissionXOffset(markupContext),
                baseline + frame.BaselineInset - frame.Height - yOffset,
                frame.Width, frame.Height);
        }
    }
}
