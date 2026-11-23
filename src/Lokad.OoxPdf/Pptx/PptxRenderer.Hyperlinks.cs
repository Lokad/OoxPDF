using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Node- and frame-level hyperlink clicks. Run-level clicks are collected from
    // positioned text spans (PptxRenderer.TextHyperlinks); chart text uses its
    // own part relationships and the same span-box geometry.
    private static void AddNodeHyperlinkAnnotation(
        PptxSceneNode node,
        PptxRenderContext context,
        string? sourcePartName,
        GroupTransform transform,
        List<PdfLinkAnnotation> annotations,
        HashSet<string> reportedIds)
    {
        if (!node.HyperlinkClick.IsDefined ||
            node.Kind is PptxSceneNodeKind.Unknown)
        {
            return;
        }

        ShapeBounds? sourceBounds = node.Bounds is { } bounds ? ToShapeBounds(bounds)
            : node.Kind == PptxSceneNodeKind.Shape
                ? ReadShapeBounds(node.Source, FindInheritedPlaceholderShapes(node.Source, context.InheritedXml), out _)
                : null;
        if (sourceBounds is null) return;
        ShapeBounds mapped = transform.Apply(sourceBounds.Value);
        double x = OoxUnits.EmuToPoints(mapped.X);
        double yTop = OoxUnits.EmuToPoints(mapped.Y);
        double width = OoxUnits.EmuToPoints(mapped.Width);
        double height = OoxUnits.EmuToPoints(mapped.Height);
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        double y = context.Document.SlideHeightPoints - yTop - height;
        // Chart/table frame painters retain an unrotated frame fallback. Keep
        // navigation on their actual paint; only shape-based owners use this matrix.
        if (node.Kind is PptxSceneNodeKind.Shape or PptxSceneNodeKind.Connector or PptxSceneNodeKind.Picture or PptxSceneNodeKind.Group &&
            (Math.Abs(mapped.RotationDegrees) > 0.001d || mapped.FlipHorizontal || mapped.FlipVertical))
        {
            PdfRectangle frame = TransformedRectangleBounds(new PdfRectangle(x, y, width, height),
                ShapeTransformMatrix(x, y, width, height, mapped));
            (x, y, width, height) = (frame.X, frame.Y, frame.Width, frame.Height);
        }

        if (!TryIntersectWithSlideBounds(x, y, width, height, context.Document, out x, out y, out width, out height))
        {
            return;
        }

        string partName = sourcePartName ?? context.SceneSlide.PartName;
        string? relationshipId = node.HyperlinkClick.RelationshipId;
        if (string.IsNullOrEmpty(relationshipId))
        {
            if (!string.IsNullOrEmpty(node.HyperlinkClick.Action))
            {
                EmitHyperlinkDiagnostic(context, partName, reportedIds, node.HyperlinkClick.Action,
                    "PPTX_UNSUPPORTED_HYPERLINK_ACTION", $"Hyperlink action '{node.HyperlinkClick.Action}' is not supported and was ignored.");
            }

            return;
        }

        PptxHyperlinkResolution resolution = ResolveHyperlink(context, sourcePartName, relationshipId);
        if (!resolution.Resolved)
        {
            EmitHyperlinkDiagnostic(context, partName, reportedIds, relationshipId, resolution.FailureId, resolution.FailureMessage);

            return;
        }

        annotations.Add(resolution.Uri is not null
            ? PdfLinkAnnotation.ToUri(x, y, width, height, resolution.Uri)
            : PdfLinkAnnotation.ToDestination(x, y, width, height, resolution.Destination!.Value));
    }

    // Chart-title run clicks resolve against the chart part's own relationships so
    // title links navigate independently of the whole-frame link (S08). Run boxes
    // reuse the shared span-box math; chart layout space is already PDF-flipped.
    private static async ValueTask AddChartTextRunHyperlinkAnnotationAsync(
        ChartTextRunLink link,
        IReadOnlyDictionary<string, OoxRelationship> chartRelationships,
        PptxRenderContext context,
        string? chartPartName,
        List<PdfLinkAnnotation> annotations,
        HashSet<string> reportedIds,
        IReadOnlyDictionary<FontRequest, RenderedFont> renderedFonts)
    {
        TextRun run = link.Run;
        if (run.Text.Length == 0 || run.Width <= 0d)
        {
            return;
        }

        if (string.IsNullOrEmpty(link.RelationshipId))
        {
            if (!string.IsNullOrEmpty(link.Action))
            {
                EmitHyperlinkDiagnostic(context, chartPartName ?? context.SceneSlide.PartName, reportedIds, link.Action,
                    "PPTX_UNSUPPORTED_HYPERLINK_ACTION", $"Hyperlink action '{link.Action}' is not supported and was ignored.");
            }

            return;
        }

        PptxHyperlinkResolution resolution = ResolveHyperlink(context, chartRelationships, link.RelationshipId);
        if (!resolution.Resolved)
        {
            EmitHyperlinkDiagnostic(context, chartPartName ?? context.SceneSlide.PartName, reportedIds, link.RelationshipId,
                resolution.FailureId, resolution.FailureMessage);

            return;
        }

        if (renderedFonts.ContainsKey(FontRequestForRun(run)))
        {
            AddPaintedRun(run);
        }
        else if (renderedFonts.Count != 0)
        {
            // A completely substituted run has no resource under its requested
            // family. Navigation consumes the same split and prepared faces as paint.
            foreach (TextRun part in await SplitRunsByResolvedTypefaceAsync([run], context.FontResolver, null,
                context.CancellationToken).ConfigureAwait(false)) AddPaintedRun(part);
        }

        void AddPaintedRun(TextRun painted)
        {
            if (!renderedFonts.TryGetValue(FontRequestForRun(painted), out RenderedFont rendered) ||
                (rendered.Font is null && rendered.FallbackFace is null) ||
                !TryGetSpanLinkRect(painted, rendered.Font, context.Document, out double x, out double y, out double width, out double height)) return;

            annotations.Add(resolution.Uri is not null
                ? PdfLinkAnnotation.ToUri(x, y, width, height, resolution.Uri)
                : PdfLinkAnnotation.ToDestination(x, y, width, height, resolution.Destination!.Value));
        }
    }

    // One annotation per linked chart-text run for axis titles and data labels (S08).
    // Null relationships (scene-less fallback rendering) skip emission silently.
    private static async ValueTask AddChartTextRunHyperlinkAnnotationsAsync(
        List<ChartTextRunLink>? links,
        PptxSceneChart? chart,
        PptxRenderContext context,
        RenderedFonts renderedFonts,
        List<PdfLinkAnnotation> annotations,
        HashSet<string> reportedIds)
    {
        if (links is null || chart?.Relationships is not { } chartRelationships)
        {
            return;
        }

        foreach (ChartTextRunLink link in links)
        {
            await AddChartTextRunHyperlinkAnnotationAsync(link, chartRelationships, context, chart.TargetPartName,
                annotations, reportedIds, renderedFonts.Fonts).ConfigureAwait(false);
        }
    }

    private readonly record struct PptxHyperlinkResolution(
        bool Resolved,
        string? Uri,
        PdfLinkDestination? Destination,
        string FailureId,
        string FailureMessage);

    private static PptxHyperlinkResolution ResolveHyperlink(
        PptxRenderContext context,
        string? sourcePartName,
        string relationshipId)
    {
        return ResolveHyperlink(context, ResolvePartRelationships(context.SceneSlide, sourcePartName), relationshipId);
    }

    // Chart-text run clicks resolve against the chart part's own relationships (S08).
    private static PptxHyperlinkResolution ResolveHyperlink(
        PptxRenderContext context,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        string relationshipId)
    {

        if (!relationships.TryGetValue(relationshipId, out OoxRelationship? relationship))
        {
            return new PptxHyperlinkResolution(false, null, null, "PPTX_UNSUPPORTED_HYPERLINK", $"Hyperlink relationship '{relationshipId}' could not be resolved and was ignored.");
        }

        if (relationship.IsExternal)
        {
            if (string.IsNullOrEmpty(relationship.Target))
            {
                return new PptxHyperlinkResolution(false, null, null, "PPTX_UNSUPPORTED_HYPERLINK", $"Hyperlink relationship '{relationshipId}' has no target and was ignored.");
            }

            return new PptxHyperlinkResolution(true, relationship.Target, null, string.Empty, string.Empty);
        }

        if (relationship.Type.Equals(PptxReader.SlideRelationshipType, StringComparison.Ordinal) &&
            relationship.ResolvedTarget is not null)
        {
            int pageIndex = -1;
            for (int i = 0; i < context.Document.Slides.Count; i++)
            {
                if (string.Equals(context.Document.Slides[i].PartName, relationship.ResolvedTarget, StringComparison.OrdinalIgnoreCase))
                {
                    pageIndex = i;
                    break;
                }
            }

            if (pageIndex < 0)
            {
                return new PptxHyperlinkResolution(false, null, null, "PPTX_UNSUPPORTED_HYPERLINK", $"Hyperlink relationship '{relationshipId}' points at an unknown slide part and was ignored.");
            }

            return new PptxHyperlinkResolution(true, null, new PdfLinkDestination(pageIndex, null, null, null), string.Empty, string.Empty);
        }

        return new PptxHyperlinkResolution(false, null, null, "PPTX_UNSUPPORTED_HYPERLINK", $"Hyperlink relationship '{relationshipId}' of type '{relationship.Type}' is not supported and was ignored.");
    }

    private static IReadOnlyDictionary<string, OoxRelationship> ResolvePartRelationships(PptxSceneSlide slide, string? sourcePartName)
    {
        if (string.Equals(sourcePartName, slide.MasterPartName, StringComparison.Ordinal))
        {
            return slide.MasterRelationships;
        }

        if (string.Equals(sourcePartName, slide.LayoutPartName, StringComparison.Ordinal))
        {
            return slide.LayoutRelationships;
        }

        return slide.SlideRelationships;
    }

    private static void EmitHyperlinkDiagnostic(PptxRenderContext context, string partName,
        HashSet<string> reportedIds, string localIdentity, string id, string message)
    {
        // Relationship IDs and actions belong to their source part, even when
        // node, run and chart painters share one reporting set for the slide.
        if (!reportedIds.Add(partName + "\0" + id + "\0" + localIdentity)) return;
        context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic(
            id,
            OoxPdfSeverity.Warning,
            message,
            partName,
            PageIndex: null,
            SlideIndex: context.SlideNumber,
            Feature: "hyperlink",
            Fallback: "Ignored"));
    }
}
