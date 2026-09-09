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
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        string? sourcePartName,
        PptxColorMap sourceColorMap,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders,
        CancellationToken cancellationToken)
    {
        foreach (PptxSceneNode node in nodes)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PdfGraphicsBuilder.ContentMark contentMark = graphics.MarkContent();
            int annotationCount = linkAnnotations.Count;
            try
            {
                RenderOrderedSceneNode(
                    node,
                    context,
                    graphics,
                    fonts,
                    images,
                    chartFonts,
                    linkAnnotations,
                    reportedHyperlinkIds,
                    sourcePartName,
                    sourceColorMap,
                    ref imageIndex,
                    transform,
                    renderPlaceholders,
                    cancellationToken);
            }
            catch (Exception ex) when (IsRecoverableNodeRenderException(ex))
            {
                graphics.TruncateContent(contentMark);
                linkAnnotations.RemoveRange(annotationCount, linkAnnotations.Count - annotationCount);
                EmitNodeRenderFailureDiagnostic(node);
            }
        }

        void EmitNodeRenderFailureDiagnostic(PptxSceneNode node)
        {
            context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "PPTX_NODE_RENDER_FAILED",
                OoxPdfSeverity.Warning,
                "PPTX node rendering failed and the node was ignored while rendering continued.",
                sourcePartName ?? context.SlidePartName,
                PageIndex: null,
                SlideIndex: context.SlideNumber,
                Feature: node.Kind.ToString(),
                Fallback: "Ignored"));
        }

        bool IsRecoverableNodeRenderException(Exception exception)
        {
            return exception is FormatException or InvalidDataException or NotSupportedException or ArgumentException;
        }
    }

    private static void RenderOrderedSceneNode(
        PptxSceneNode node,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        List<PdfImageResource> images,
        List<PdfFontResource> chartFonts,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
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
                    if (HasDrawableShape())
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

                    // Node-level links go first so run-level links sit above them.
                    AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                    RenderTextNode(
                        node,
                        context,
                        graphics,
                        fonts,
                        sourceColorMap,
                        renderPlaceholders,
                        linkAnnotations,
                        reportedHyperlinkIds,
                        sourcePartName);
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
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                break;
            case PptxSceneNodeKind.Picture:
                RenderPicture(
                    node,
                    context,
                    graphics,
                    transform,
                    images,
                    ref imageIndex);
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                break;
            case PptxSceneNodeKind.Table:
                BeginSlideNodeClip(context, graphics);
                // Frame-level links go first so cell-run links sit above them.
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = RenderTableFrame(context, node, graphics, transform, sourceColorMap);
                var tableHyperlinkScope = new PptxTextHyperlinkScope(context, sourcePartName, linkAnnotations, reportedHyperlinkIds);
                DrawTextSpansWithFonts(tableTextSpans, graphics, fonts, tableHyperlinkScope);
                EndSlideNodeClip(graphics);
                break;
            case PptxSceneNodeKind.Chart:
                BeginSlideNodeClip(context, graphics);
                cancellationToken.ThrowIfCancellationRequested();
                // Node-level links go first so chart run links sit above them.
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                RenderChartFrame(context, graphics, chartFonts, node, transform, linkAnnotations, reportedHyperlinkIds);
                EndSlideNodeClip(graphics);
                break;
            case PptxSceneNodeKind.UnknownGraphicFrame:
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                RenderUnsupportedGraphicFrame(node, context, sourcePartName);
                break;
            case PptxSceneNodeKind.Group:
                AddNodeHyperlinkAnnotation(node, context, sourcePartName, transform, linkAnnotations, reportedHyperlinkIds);
                RenderOrderedSceneNodes(
                    node.Children,
                    context,
                    graphics,
                    fonts,
                    images,
                    chartFonts,
                    linkAnnotations,
                    reportedHyperlinkIds,
                    sourcePartName,
                    sourceColorMap,
                    ref imageIndex,
                    transform.Combine(ToGroupTransform(node.GroupTransform)),
                    renderPlaceholders,
                    cancellationToken);
                break;
        }

        bool HasDrawableShape()
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
        bool renderPlaceholders,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        string? sourcePartName)
    {
        if (node.TextBody is null)
        {
            return;
        }

        var hyperlinkScope = new PptxTextHyperlinkScope(context, sourcePartName, linkAnnotations, reportedHyperlinkIds);
        DrawTextSpansWithFonts(ReadTextSpansForSceneNode(node, context, sourceColorMap, renderPlaceholders), graphics, fonts, hyperlinkScope);
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
            PageIndex: null,
            SlideIndex: context.SlideNumber,
            Feature: "graphic frame",
            Fallback: "Ignored"));
    }

}
