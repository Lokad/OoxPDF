using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool CanRenderSlideInOrder(PptxSceneSlide sceneSlide)
    {
        return !sceneSlide.SlideNodes.Any(ContainsUnknownGraphicFrame);
    }

    private static bool ContainsUnknownGraphicFrame(PptxSceneNode node)
    {
        return node.Kind == PptxSceneNodeKind.UnknownGraphicFrame ||
            node.Children.Any(ContainsUnknownGraphicFrame);
    }

    private static void RenderOrderedSceneNodes(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        List<PdfImageResource> images,
        List<PdfFontResource> chartFonts,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders)
    {
        foreach (PptxSceneNode node in nodes)
        {
            XElement source = node.Source;
            switch (node.Kind)
            {
                case PptxSceneNodeKind.Shape:
                    if (renderPlaceholders || !node.IsPlaceholder)
                    {
                        RenderShape(
                            node,
                            relationships,
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
                        DrawTextSpansWithFonts(ReadTextSpansForShape(source, context, renderPlaceholders), graphics, fonts);
                    }

                    break;
                case PptxSceneNodeKind.Connector:
                    RenderShape(
                        node,
                        relationships,
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
                    break;
                case PptxSceneNodeKind.Picture:
                    RenderPicture(
                        node,
                        context,
                        graphics,
                        transform,
                        images,
                        relationships,
                        ref imageIndex);
                    break;
                case PptxSceneNodeKind.Table:
                    IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = RenderTableFrame(context, source, graphics);
                    DrawTextSpansWithFonts(tableTextSpans, graphics, fonts);
                    break;
                case PptxSceneNodeKind.Chart:
                    RenderChartFrame(context, graphics, chartFonts, source, transform, relationships);
                    break;
                case PptxSceneNodeKind.Group:
                    RenderOrderedSceneNodes(
                        node.Children,
                        context,
                        graphics,
                        fonts,
                        images,
                        chartFonts,
                        relationships,
                        ref imageIndex,
                        transform.Combine(ReadGroupTransform(source)),
                        renderPlaceholders);
                    break;
            }
        }
    }
}
