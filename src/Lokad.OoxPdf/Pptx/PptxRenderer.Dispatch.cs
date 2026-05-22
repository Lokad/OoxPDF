using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool CanRenderSlideInOrder(XDocument slideXml)
    {
        return !slideXml
            .Descendants(PresentationNamespace + "graphicFrame")
            .Any(frame => !IsTableGraphicFrame(frame));
    }

    private static void RenderOrderedShapeTextContainer(
        XElement container,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        List<PdfImageResource> images,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == PresentationNamespace + "sp")
            {
                if (renderPlaceholders || !IsPlaceholder(child))
                {
                    RenderShape(
                        child,
                        context.SlideRelationships,
                        context.Package,
                        context.Document,
                        graphics,
                        context.DiagnosticSink,
                        context.SlideNumber,
                        context.Theme,
                        transform,
                        images,
                        context.ImageCache,
                        ref imageIndex);
                    DrawTextSpansWithFonts(ReadTextSpansForShape(child, context, renderPlaceholders), graphics, fonts);
                }

                continue;
            }

            if (child.Name == PresentationNamespace + "cxnSp")
            {
                RenderShape(
                    child,
                    context.SlideRelationships,
                    context.Package,
                    context.Document,
                    graphics,
                    context.DiagnosticSink,
                    context.SlideNumber,
                    context.Theme,
                    transform,
                    images,
                    context.ImageCache,
                    ref imageIndex);
                continue;
            }

            if (child.Name == PresentationNamespace + "pic")
            {
                RenderPicture(
                    child,
                    context.SlideRelationships,
                    context.Package,
                    context.Document,
                    context.Theme,
                    graphics,
                    context.DiagnosticSink,
                    context.SlideNumber,
                    transform,
                    images,
                    context.ImageCache,
                    ref imageIndex);
                continue;
            }

            if (child.Name == PresentationNamespace + "graphicFrame")
            {
                IReadOnlyList<TextRun> tableTextRuns = RenderTableFrame(context, child, graphics);
                DrawTextRunsWithFonts(tableTextRuns, graphics, fonts);
                continue;
            }

            if (child.Name == PresentationNamespace + "grpSp")
            {
                RenderOrderedShapeTextContainer(
                    child,
                    context,
                    graphics,
                    fonts,
                    images,
                    ref imageIndex,
                    transform.Combine(ReadGroupTransform(child)),
                    renderPlaceholders);
            }
        }
    }
}
