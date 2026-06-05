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
        string? sourcePartName,
        PptxColorMap sourceColorMap,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders,
        CancellationToken cancellationToken = default)
    {
        foreach (PptxSceneNode node in nodes)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int stateDepth = graphics.StateDepth;
            try
            {
                RenderOrderedSceneNode(
                    node,
                    context,
                    graphics,
                    fonts,
                    images,
                    chartFonts,
                    sourcePartName,
                    sourceColorMap,
                    ref imageIndex,
                    transform,
                    renderPlaceholders,
                    cancellationToken);
            }
            catch (Exception ex) when (IsRecoverableNodeRenderException(ex))
            {
                graphics.RestoreToStateDepth(stateDepth);
                EmitNodeRenderFailureDiagnostic(node, context, sourcePartName);
            }
        }
    }

    private static void RenderOrderedSceneNode(
        PptxSceneNode node,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        List<PdfImageResource> images,
        List<PdfFontResource> chartFonts,
        string? sourcePartName,
        PptxColorMap sourceColorMap,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
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
                        sourceColorMap,
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
                IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = RenderTableFrame(context, node, graphics, transform, sourceColorMap);
                DrawTextSpansWithFonts(tableTextSpans, graphics, fonts);
                EndSlideNodeClip(graphics);
                break;
            case PptxSceneNodeKind.Chart:
                BeginSlideNodeClip(context, graphics);
                cancellationToken.ThrowIfCancellationRequested();
                RenderChartFrame(context, graphics, chartFonts, node, transform);
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
                    sourcePartName,
                    sourceColorMap,
                    ref imageIndex,
                    transform.Combine(ToGroupTransform(node.GroupTransform)),
                    renderPlaceholders,
                    cancellationToken);
                break;
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

    private static void RenderTextNode(
        PptxSceneNode node,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        PptxColorMap sourceColorMap,
        bool renderPlaceholders)
    {
        if (node.TextBody is null)
        {
            return;
        }

        DrawTextSpansWithFonts(ReadTextSpansForSceneNode(node, context, sourceColorMap, renderPlaceholders), graphics, fonts);
    }

    private static void RenderUnsupportedGraphicFrame(PptxSceneNode node, PptxRenderContext context, string? sourcePartName)
    {
        string effectivePartName = sourcePartName ?? context.SlidePartName;
        if (context.DiagnosticSink is null || effectivePartName == context.SlidePartName || node.IsSmartArtGraphicFrame)
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

    private static bool IsRecoverableNodeRenderException(Exception exception)
    {
        return exception is FormatException or InvalidDataException or NotSupportedException or ArgumentException;
    }

    private static void EmitNodeRenderFailureDiagnostic(PptxSceneNode node, PptxRenderContext context, string? sourcePartName)
    {
        context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic(
            "PPTX_NODE_RENDER_FAILED",
            OoxPdfSeverity.Warning,
            "PPTX node rendering failed and the node was ignored while rendering continued.",
            sourcePartName ?? context.SlidePartName,
            SlideIndex: context.SlideNumber,
            Feature: node.Kind.ToString(),
            Fallback: "Ignored"));
    }
}
