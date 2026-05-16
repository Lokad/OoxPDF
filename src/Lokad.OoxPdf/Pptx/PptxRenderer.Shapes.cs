using Lokad.OoxPdf.Pdf;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderBackground(PptxRenderContext context, XDocument slideXml, PdfGraphicsBuilder graphics)
    {
        XElement? background = slideXml.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        if (TryReadSolidColor(background, context.Theme, out RgbColor color))
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(0, 0, context.Document.SlideWidthPoints, context.Document.SlideHeightPoints);
        }
    }

    private static void RenderShapes(PptxRenderContext context, XDocument slideXml, PdfGraphicsBuilder graphics, bool renderPlaceholders)
    {
        foreach (XElement shapeTree in slideXml.Descendants(PresentationNamespace + "spTree"))
        {
            RenderShapeContainer(context, shapeTree, graphics, GroupTransform.Identity, renderPlaceholders);
        }
    }

    private static void RenderShapeContainer(
        PptxRenderContext context,
        XElement container,
        PdfGraphicsBuilder graphics,
        GroupTransform transform,
        bool renderPlaceholders)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == PresentationNamespace + "sp")
            {
                if (renderPlaceholders || !IsPlaceholder(child))
                {
                    RenderShape(child, context.Document, graphics, context.Theme, transform);
                }

                continue;
            }

            if (child.Name == PresentationNamespace + "cxnSp")
            {
                RenderShape(child, context.Document, graphics, context.Theme, transform);
                continue;
            }

            if (child.Name == PresentationNamespace + "grpSp")
            {
                GroupTransform childTransform = transform.Combine(ReadGroupTransform(child));
                RenderShapeContainer(context, child, graphics, childTransform, renderPlaceholders);
            }
        }
    }

    private static void RenderShape(
        XElement shape,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        PptxTheme theme,
        GroupTransform groupTransform)
    {
        int imageIndex = 1;
        RenderShape(
            shape,
            relationships: null,
            package: null,
            document,
            graphics,
            diagnosticSink: null,
            slideIndex: 0,
            theme,
            groupTransform,
            images: null,
            imageCache: null,
            ref imageIndex);
    }
}
