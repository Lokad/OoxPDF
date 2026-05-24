using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderOrderedSceneNodes(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        List<PdfImageResource> images,
        List<PdfFontResource> chartFonts,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        string? sourcePartName,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders)
    {
        foreach (PptxSceneNode node in nodes)
        {
            switch (node.Kind)
            {
                case PptxSceneNodeKind.Shape:
                    if (renderPlaceholders || !node.IsPlaceholder)
                    {
                        RenderShapeNode(
                            node,
                            context,
                            graphics,
                            fonts,
                            images,
                            relationships,
                            transform,
                            ref imageIndex,
                            renderPlaceholders);
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
                    IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = RenderTableFrame(context, node, graphics, transform);
                    DrawTextSpansWithFonts(tableTextSpans, graphics, fonts);
                    break;
                case PptxSceneNodeKind.Chart:
                    RenderChartFrame(context, graphics, chartFonts, node, transform, relationships);
                    break;
                case PptxSceneNodeKind.UnknownGraphicFrame:
                    RenderUnsupportedGraphicFrame(node, context, sourcePartName);
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
                        sourcePartName,
                        ref imageIndex,
                        transform.Combine(ToGroupTransform(node.GroupTransform)),
                        renderPlaceholders);
                    break;
            }
        }
    }

    private static void RenderShapeNode(
        PptxSceneNode node,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        List<PdfImageResource> images,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        GroupTransform transform,
        ref int imageIndex,
        bool renderPlaceholders)
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
        RenderTextNode(node, context, graphics, fonts, renderPlaceholders);
    }

    private static void RenderTextNode(
        PptxSceneNode node,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        bool renderPlaceholders)
    {
        if (node.TextBody is null)
        {
            return;
        }

        DrawTextSpansWithFonts(ReadTextSpansForSceneNode(node, context, renderPlaceholders), graphics, fonts);
    }

    private static void RenderUnsupportedGraphicFrame(PptxSceneNode node, PptxRenderContext context, string? sourcePartName)
    {
        string effectivePartName = sourcePartName ?? context.Slide.PartName;
        if (context.DiagnosticSink is null || effectivePartName == context.Slide.PartName || node.IsSmartArtGraphicFrame)
        {
            return;
        }

        context.DiagnosticSink(new OoxPdfDiagnostic(
            "PPTX_UNSUPPORTED_GRAPHIC_FRAME",
            OoxPdfSeverity.Warning,
            "Unsupported PPTX graphic frame was detected and ignored.",
            effectivePartName,
            SlideIndex: context.SlideNumber,
            Feature: "graphic frame",
            Fallback: "Ignored"));
    }
}
