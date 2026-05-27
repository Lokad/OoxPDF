using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static PptxSceneSnapshot InspectScene(PptxDocument document, OoxPackage package)
    {
        PptxScene scene = new PptxSceneBuilder().Build(document, package);
        return new PptxSceneSnapshot(scene.Slides.Select(ToSnapshot).ToArray());
    }

    private static PptxSceneSlideSnapshot ToSnapshot(PptxSceneSlide slide)
    {
        return new PptxSceneSlideSnapshot(
            slide.Index,
            slide.PartName,
            slide.MasterPartName,
            slide.LayoutPartName,
            slide.MasterXml is not null,
            slide.LayoutXml is not null,
            slide.SlideXml.Root is not null,
            slide.MasterRelationships.Count,
            slide.LayoutRelationships.Count,
            slide.SlideRelationships.Count,
            slide.MasterBackground.HasFill,
            slide.LayoutBackground.HasFill,
            slide.SlideBackground.HasFill,
            slide.MasterNodes.Select(ToSnapshot).ToArray(),
            slide.LayoutNodes.Select(ToSnapshot).ToArray(),
            slide.SlideNodes.Select(ToSnapshot).ToArray());
    }

    private static PptxSceneNodeSnapshot ToSnapshot(PptxSceneNode node)
    {
        int textParagraphCount = node.TextBody?.Paragraphs.Count ?? 0;
        int textRunCount = node.TextBody?.Paragraphs.Sum(paragraph => paragraph.Runs.Count) ?? 0;
        int tableRowCount = node.Table?.Rows.Count ?? 0;
        int tableCellCount = node.Table?.Rows.Sum(row => row.Cells.Count) ?? 0;
        PptxSceneBounds? bounds = node.Bounds;
        return new PptxSceneNodeSnapshot(
            node.Kind.ToString(),
            node.IsPlaceholder,
            node.IsSmartArtGraphicFrame,
            node.HyperlinkClick.IsDefined,
            node.HyperlinkClick.RelationshipId,
            node.HyperlinkClick.Action,
            bounds is not null,
            bounds?.RotationDegrees ?? 0d,
            bounds?.FlipHorizontal ?? false,
            bounds?.FlipVertical ?? false,
            node.Shape is not null,
            node.Shape?.Preset ?? string.Empty,
            node.TextBody is not null,
            textParagraphCount,
            textRunCount,
            node.Picture is not null,
            node.Picture?.Resource is not null,
            node.Picture?.Resource?.ContentType ?? string.Empty,
            node.Shape?.PictureFill.Resource is not null,
            node.Shape?.PictureFill.Resource?.ContentType ?? string.Empty,
            node.Picture?.Recolor.Kind.ToString() ?? string.Empty,
            node.Table is not null,
            tableRowCount,
            tableCellCount,
            node.Chart is not null,
            node.Chart?.Plots.Count ?? 0,
            node.Chart?.Axes.Count ?? 0,
            node.Chart?.ExternalData.IsDefined ?? false,
            node.Chart?.ExternalData.RelationshipId ?? string.Empty,
            node.Chart?.ExternalData.TargetPartName ?? string.Empty,
            node.Chart?.ExternalData.AutoUpdate,
            node.Chart?.ExternalData.Resource is not null,
            node.Chart?.ExternalData.Resource?.ContentType ?? string.Empty,
            node.Chart?.Options.PlotVisibleOnly,
            node.Kind == PptxSceneNodeKind.Group,
            node.Children.Select(ToSnapshot).ToArray());
    }
}
