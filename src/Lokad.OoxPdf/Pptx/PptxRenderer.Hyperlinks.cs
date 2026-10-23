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
            node.Kind is PptxSceneNodeKind.Unknown ||
            node.Bounds is not { } bounds)
        {
            return;
        }

        ShapeBounds mapped = transform.Apply(ToShapeBounds(bounds));
        double x = OoxUnits.EmuToPoints(mapped.X);
        double yTop = OoxUnits.EmuToPoints(mapped.Y);
        double width = OoxUnits.EmuToPoints(mapped.Width);
        double height = OoxUnits.EmuToPoints(mapped.Height);
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        double y = context.Document.SlideHeightPoints - yTop - height;
        string partName = sourcePartName ?? context.SlidePartName;
        string? relationshipId = node.HyperlinkClick.RelationshipId;
        if (string.IsNullOrEmpty(relationshipId))
        {
            if (!string.IsNullOrEmpty(node.HyperlinkClick.Action) && reportedIds.Add("action:" + node.HyperlinkClick.Action))
            {
                EmitHyperlinkDiagnostic(context, partName, "PPTX_UNSUPPORTED_HYPERLINK_ACTION", $"Hyperlink action '{node.HyperlinkClick.Action}' is not supported and was ignored.");
            }

            return;
        }

        PptxHyperlinkResolution resolution = ResolveHyperlink(context, sourcePartName, relationshipId);
        if (!resolution.Resolved)
        {
            if (reportedIds.Add(resolution.FailureId + "\u0000" + relationshipId))
            {
                EmitHyperlinkDiagnostic(context, partName, resolution.FailureId, resolution.FailureMessage);
            }

            return;
        }

        annotations.Add(resolution.Uri is not null
            ? PdfLinkAnnotation.ToUri(x, y, width, height, resolution.Uri)
            : PdfLinkAnnotation.ToDestination(x, y, width, height, resolution.Destination!.Value));
    }

    // Chart-title run clicks resolve against the chart part's own relationships so
    // title links navigate independently of the whole-frame link (S08). Run boxes
    // reuse the shared span-box math; chart layout space is already PDF-flipped.
    private static void AddChartTextRunHyperlinkAnnotation(
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
            if (!string.IsNullOrEmpty(link.Action) && reportedIds.Add("action:" + link.Action))
            {
                EmitHyperlinkDiagnostic(context, chartPartName ?? context.SlidePartName, "PPTX_UNSUPPORTED_HYPERLINK_ACTION", $"Hyperlink action '{link.Action}' is not supported and was ignored.");
            }

            return;
        }

        PptxHyperlinkResolution resolution = ResolveHyperlink(context, chartRelationships, link.RelationshipId);
        if (!resolution.Resolved)
        {
            if (reportedIds.Add(resolution.FailureId + "\u0000" + link.RelationshipId))
            {
                EmitHyperlinkDiagnostic(context, chartPartName ?? context.SlidePartName, resolution.FailureId, resolution.FailureMessage);
            }

            return;
        }

        // RV01: fallback spans stay link-free like other unresolved spans.
        if (!renderedFonts.TryGetValue(FontRequestForRun(run), out RenderedFont rendered) ||
            rendered.Font is not { } linkEmbedded ||
            !TryGetSpanLinkRect(run, linkEmbedded, out double x, out double y, out double width, out double height))
        {
            return;
        }

        annotations.Add(resolution.Uri is not null
            ? PdfLinkAnnotation.ToUri(x, y, width, height, resolution.Uri)
            : PdfLinkAnnotation.ToDestination(x, y, width, height, resolution.Destination!.Value));
    }

    // One annotation per linked chart-text run for axis titles and data labels (S08).
    // Null relationships (scene-less fallback rendering) skip emission silently.
    private static void AddChartTextRunHyperlinkAnnotations(
        List<ChartTextRunLink>? links,
        IReadOnlyDictionary<string, OoxRelationship>? chartRelationships,
        PptxRenderContext context,
        RenderedFonts renderedFonts,
        List<PdfLinkAnnotation> annotations,
        HashSet<string> reportedIds)
    {
        if (links is null || chartRelationships is null)
        {
            return;
        }

        foreach (ChartTextRunLink link in links)
        {
            AddChartTextRunHyperlinkAnnotation(link, chartRelationships, context, null, annotations, reportedIds, renderedFonts.Fonts);
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
        return ResolveHyperlink(context, ResolvePartRelationships(context, sourcePartName), relationshipId);
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

    private static IReadOnlyDictionary<string, OoxRelationship> ResolvePartRelationships(PptxRenderContext context, string? sourcePartName)
    {
        if (string.Equals(sourcePartName, context.MasterPartName, StringComparison.Ordinal))
        {
            return context.MasterRelationships;
        }

        if (string.Equals(sourcePartName, context.LayoutPartName, StringComparison.Ordinal))
        {
            return context.LayoutRelationships;
        }

        return context.SlideRelationships;
    }

    private static void EmitHyperlinkDiagnostic(PptxRenderContext context, string partName, string id, string message)
    {
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
