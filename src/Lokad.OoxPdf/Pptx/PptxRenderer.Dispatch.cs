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
                        if (HasDrawableShape(node))
                        {
                            BeginSlideNodeClip(context, graphics);
                            RenderShape(
                                node,
                                context.Document,
                                graphics,
                                context.DiagnosticSink,
                                context.SlideNumber,
                                context.Theme,
                                transform,
                                images,
                                context.ImageCache,
                                ref imageIndex);
                            EndSlideNodeClip(graphics);
                        }

                        RenderTextNode(
                            node,
                            context,
                            graphics,
                            fonts,
                            renderPlaceholders);
                    }

                    break;
                case PptxSceneNodeKind.Connector:
                    BeginSlideNodeClip(context, graphics);
                    RenderShape(
                        node,
                        context.Document,
                        graphics,
                        context.DiagnosticSink,
                        context.SlideNumber,
                        context.Theme,
                        transform,
                        images,
                        context.ImageCache,
                        ref imageIndex);
                    EndSlideNodeClip(graphics);
                    break;
                case PptxSceneNodeKind.Picture:
                    RenderPicture(
                        node,
                        context,
                        graphics,
                        transform,
                        images,
                        ref imageIndex);
                    break;
                case PptxSceneNodeKind.Table:
                    BeginSlideNodeClip(context, graphics);
                    IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = RenderTableFrame(context, node, graphics, transform);
                    DrawTextSpansWithFonts(tableTextSpans, graphics, fonts);
                    EndSlideNodeClip(graphics);
                    break;
                case PptxSceneNodeKind.Chart:
                    BeginSlideNodeClip(context, graphics);
                    RenderChartFrame(context, graphics, chartFonts, node, transform, relationships);
                    EndSlideNodeClip(graphics);
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

    private static bool HasDrawableShape(PptxSceneNode node)
    {
        return node.Shape is
        {
            Fill.HasFill: true
        } or
        {
            GradientFill.HasGradient: true
        } or
        {
            PatternFill.HasPattern: true
        } or
        {
            PictureFill.HasPicture: true
        } or
        {
            Glow.HasGlow: true
        } or
        {
            OuterShadow.HasShadow: true
        } or
        {
            Line.HasLine: true
        };
    }

    private static void BeginSlideNodeClip(PptxRenderContext context, PdfGraphicsBuilder graphics)
    {
        graphics.SaveState();
        ClipSlideBoundsEvenOdd(context.Document, graphics);
    }

    private static void EndSlideNodeClip(PdfGraphicsBuilder graphics)
    {
        graphics.RestoreState();
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
