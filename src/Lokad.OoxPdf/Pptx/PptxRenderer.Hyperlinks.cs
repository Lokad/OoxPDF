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

    // Each authored title, axis or label link covers the resolved glyph
    // fragments from that source run. Relationships belong to the chart.
    private static void AddChartTextRunHyperlinkAnnotations(
        List<ChartTextRunLink>? links,
        PptxSceneChart? chart,
        PptxRenderContext context,
        (IReadOnlyList<TextRun> Runs, IReadOnlyDictionary<FontRequest, RenderedFont> Fonts) paint,
        List<PdfLinkAnnotation> annotations,
        HashSet<string> reportedIds)
    {
        if (links is null || chart?.Relationships is not { } relationships)
        {
            return;
        }

        string partName = chart.TargetPartName ?? context.SceneSlide.PartName;
        foreach (ChartTextRunLink link in links)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (link.RunCount == 0)
            {
                continue;
            }
            if (string.IsNullOrEmpty(link.RelationshipId))
            {
                if (!string.IsNullOrEmpty(link.Action))
                {
                    EmitHyperlinkDiagnostic(context, partName, reportedIds, link.Action,
                        "PPTX_UNSUPPORTED_HYPERLINK_ACTION", $"Hyperlink action '{link.Action}' is not supported and was ignored.");
                }
                continue;
            }

            PptxHyperlinkResolution resolution = ResolveHyperlink(context, relationships, link.RelationshipId);
            if (!resolution.Resolved)
            {
                EmitHyperlinkDiagnostic(context, partName, reportedIds, link.RelationshipId,
                    resolution.FailureId, resolution.FailureMessage);
                continue;
            }

            double left = double.PositiveInfinity, bottom = double.PositiveInfinity;
            double right = double.NegativeInfinity, top = double.NegativeInfinity;
            for (int i = link.RunIndex; i < link.RunIndex + link.RunCount; i++)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                TextRun run = paint.Runs[i];
                if (run.Text.Length == 0 || run.Width <= 0d ||
                    !paint.Fonts.TryGetValue(FontRequestForRun(run), out RenderedFont rendered) ||
                    (rendered.Font is null && rendered.FallbackFace is null) ||
                    !TryGetSpanLinkRect(run, rendered.Font, context.Document, out double x, out double y, out double width, out double height))
                {
                    continue;
                }

                left = Math.Min(left, x);
                bottom = Math.Min(bottom, y);
                right = Math.Max(right, x + width);
                top = Math.Max(top, y + height);
            }

            if (right > left && top > bottom)
            {
                annotations.Add(resolution.Uri is not null
                    ? PdfLinkAnnotation.ToUri(left, bottom, right - left, top - bottom, resolution.Uri)
                    : PdfLinkAnnotation.ToDestination(left, bottom, right - left, top - bottom, resolution.Destination!.Value));
            }
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
