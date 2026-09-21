using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void RenderWordCompatibleMarkupLaneBackground(
        DocxLayoutPage page,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            return;
        }

        double bottom = page.MarginBottom + WordCompatibleAllMarkupLaneBackgroundBottomInsetPoints;
        double top = page.Height - page.MarginTop - WordCompatibleAllMarkupLaneBackgroundTopInsetPoints;
        double height = Math.Max(0d, top - bottom);
        if (height <= 0d)
        {
            return;
        }

        double width = WordCompatibleAllMarkupLaneBackgroundWidthPoints;
        double x = ShouldUseLeftMarkupLane(page)
            ? WordCompatibleAllMarkupLaneBackgroundRightBleedPoints
            : page.Width - width - WordCompatibleAllMarkupLaneBackgroundRightBleedPoints;
        if (!ShouldUseLeftMarkupLane(page) &&
            Math.Abs(markupContext.WordCompatiblePrintScale - 1d) >= 0.000000001d)
        {
            // Office (W5-X1): the gray lane is 259.4pt design wide ending at the page edge.
            width = WordCompatibleAllMarkupLaneWidthPoints * markupContext.WordCompatiblePrintScale;
            x = page.Width - width - WordCompatibleAllMarkupLaneBackgroundRightBleedPoints;
        }
        graphics.SetFillRgb(242, 242, 242);
        graphics.FillRectangle(x, bottom, width, height);
    }

    private static bool ShouldUseLeftMarkupLane(DocxLayoutPage page)
    {
        double leftAvailable = Math.Max(0d, page.MarginLeft - 8d);
        double rightAvailable = Math.Max(0d, page.MarginRight - 8d);
        return leftAvailable > rightAvailable && leftAvailable >= MinimumMarkupBalloonBodyWidthPoints;
    }

    private static void RenderWordCompatibleRevisionBar(
        DocxLayout layout,
        DocxLayoutPage page,
        int pageIndex,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext,
        CancellationToken cancellationToken)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(markupContext) ||
            !markupContext.DrawsChangeBars)
        {
            return;
        }

        double? bottom = null;
        double? top = null;
        void IncludeRevisionBounds(double y, double height)
        {
            bottom = bottom is null ? y : Math.Min(bottom.Value, y);
            top = top is null ? y + height : Math.Max(top.Value, y + height);
        }

        foreach (DocxTextLineLayout line in EnumerateRenderedPageTextLines(layout, page, pageIndex, markupContext, page.Height))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!HasTextLineRevision(line))
            {
                continue;
            }

            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            double height = Math.Max(6d, line.LineHeight ?? scaledFontSize * 1.2d);
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double y = baselineY - height * 0.25d;
            IncludeRevisionBounds(y, height);
        }

        foreach (DocxTableRowLayout row in EnumerateMarkupBalloonTableRows(page, EnumeratePageFloatingDrawings(layout, pageIndex).ToArray()))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!HasTableRowRevision(row))
            {
                continue;
            }

            IncludeRevisionBounds(row.Y, Math.Max(6d, row.Height));
        }

        if (bottom is null || top is null)
        {
            return;
        }

        double barBottom = Math.Max(0d, bottom.Value - WordCompatibleAllMarkupRevisionBarBottomOutsetPoints);
        double barTop = Math.Min(page.Height, top.Value - WordCompatibleAllMarkupRevisionBarTopInsetPoints);
        double barHeight = Math.Max(0d, barTop - barBottom);
        if (barHeight <= 0d)
        {
            return;
        }

        graphics.SetFillRgb(0, 0, 0);
        graphics.FillRectangle(
            WordCompatibleAllMarkupRevisionBarXPoints,
            barBottom,
            WordCompatibleAllMarkupRevisionBarWidthPoints,
            barHeight);

        bool HasTableRowRevision(DocxTableRowLayout row)
        {
            return row.RevisionCount != 0 ||
                row.Table.Revisions?.Count > 0;
        }

        bool HasTextLineRevision(DocxTextLineLayout line)
        {
            return CollectTextLineRevisions(line).Count != 0;
        }
    }

    private static IReadOnlyList<DocxRevisionInfo> CollectTextLineRevisions(DocxTextLineLayout line)
    {
        var revisions = new List<DocxRevisionInfo>();
        if (line.SourceParagraph is not null)
        {
            revisions.AddRange(line.SourceParagraph.Revisions);
        }

        AddTextRunRevisions(line.StyleRun, revisions);
        foreach (DocxTextSegmentLayout segment in line.Segments)
        {
            AddTextRunRevisions(segment.StyleRun, revisions);
        }

        return revisions;
    }

    private static void AddTextRunRevisions(DocxTextRun run, List<DocxRevisionInfo> revisions)
    {
        if (run.Revision is not null)
        {
            revisions.Add(run.Revision);
        }

        revisions.AddRange(run.Revisions);
    }

    private static void RenderWordCompatibleCommentRangeMarkers(
        DocxTextLineLayout line,
        DocxParagraph paragraph,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        foreach (DocxInlineReference reference in paragraph.InlineReferences.Where(reference => reference.Kind == DocxRelatedStoryKind.Comment))
        {
            DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
                string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
            if (range is not null &&
                TryResolveWordCompatibleCommentRangeBounds(line, range, ResolveTextEmissionXOffset(markupContext), out double startX, out double endX))
            {
                double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
                RenderWordCompatibleCommentRangeMarker(startX, endX, baselineY, graphics);
            }
            else if ((range is null || !HasCommentRangeBounds(range)) &&
                TryResolveWordCompatibleCommentReferenceMarkerBounds(line, reference, ResolveTextEmissionXOffset(markupContext), out startX, out endX))
            {
                double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
                RenderWordCompatibleCommentRangeMarker(startX, endX, baselineY, graphics);
            }
        }

        bool HasCommentRangeBounds(DocxCommentRange range)
        {
            return range.StartSourceRunIndex is not null ||
                range.StartTextOffset is not null ||
                range.EndSourceRunIndex is not null ||
                range.EndTextOffset is not null;
        }
    }

    private static bool TryResolveWordCompatibleCommentReferenceMarkerBounds(
        DocxTextLineLayout line,
        DocxInlineReference reference,
        double wordCompatibleXOffset,
        out double startX,
        out double endX)
    {
        if (!TryResolveSourceOffsetAnchorX(line, reference.SourceRunIndex, reference.TextOffsetInRun, out double anchorX) &&
            !TryResolvePreviousSourceRunEndAnchorX(line, reference.SourceRunIndex, minimumSourceRunIndex: null, out anchorX))
        {
            startX = 0d;
            endX = 0d;
            return false;
        }

        double centerX = anchorX + wordCompatibleXOffset;
        startX = centerX - WordCompatibleAllMarkupCommentReferenceMarkerWidthPoints / 2d;
        endX = centerX + WordCompatibleAllMarkupCommentReferenceMarkerWidthPoints / 2d;
        return endX > startX;
    }

    private static bool TryResolveWordCompatibleCommentRangeBounds(
        DocxTextLineLayout line,
        DocxCommentRange range,
        double wordCompatibleXOffset,
        out double startX,
        out double endX)
    {
        startX = 0d;
        endX = 0d;
        DocxTextSegmentLayout[] rangeSegments = line.Segments
            .Where(segment => SegmentOverlapsCommentRange(segment, range))
            .ToArray();
        if (rangeSegments.Length == 0)
        {
            return false;
        }

        if (!TryResolveSourceOffsetAnchorX(line, range.StartSourceRunIndex, range.StartTextOffset, out double layoutStartX))
        {
            layoutStartX = rangeSegments.Min(segment => segment.X);
        }

        if (!TryResolveSourceOffsetAnchorX(line, range.EndSourceRunIndex, range.EndTextOffset, out double layoutEndX) &&
            !TryResolvePreviousSourceRunEndAnchorX(line, range.EndSourceRunIndex, range.StartSourceRunIndex, out layoutEndX))
        {
            layoutEndX = rangeSegments.Max(segment => segment.X + Math.Max(0d, segment.Width));
        }

        if (layoutEndX <= layoutStartX)
        {
            return false;
        }

        startX = layoutStartX + wordCompatibleXOffset -
            WordCompatibleAllMarkupCommentRangeStartInsetPoints;
        endX = layoutEndX + wordCompatibleXOffset -
            WordCompatibleAllMarkupCommentRangeEndInsetPoints;
        return endX > startX;
    }

    private static bool SegmentOverlapsCommentRange(
        DocxTextSegmentLayout segment,
        DocxCommentRange range)
    {
        if (segment.SourceTextRunIndex < 0 ||
            range.StartSourceRunIndex is not { } startRunIndex)
        {
            return false;
        }

        int endRunIndex = range.EndSourceRunIndex ?? range.ReferenceSourceRunIndex ?? startRunIndex;
        int segmentRunIndex = segment.SourceTextRunIndex;
        if (segmentRunIndex < startRunIndex || segmentRunIndex > endRunIndex)
        {
            return false;
        }

        int segmentStartOffset = Math.Max(0, segment.SourceTextOffsetInRun);
        int segmentEndOffset = segmentStartOffset + Math.Max(0, segment.Text.Length);
        if (segmentRunIndex == startRunIndex &&
            segmentEndOffset <= Math.Max(0, range.StartTextOffset ?? 0))
        {
            return false;
        }

        if (segmentRunIndex == endRunIndex &&
            segmentStartOffset >= Math.Max(0, range.EndTextOffset ?? int.MaxValue))
        {
            return false;
        }

        return segmentEndOffset > segmentStartOffset || segment.Width > 0d;
    }

    private static void RenderWordCompatibleCommentRangeMarker(
        double startX,
        double endX,
        double baselineY,
        PdfGraphicsBuilder graphics)
    {
        DocxMarkupBalloonRgb color = WordCompatibleAllMarkupReviewStrokeRgb;
        double bottomTickY = baselineY + WordCompatibleAllMarkupCommentRangeBottomTickYOffsetPoints;
        double verticalBottomY = baselineY + WordCompatibleAllMarkupCommentRangeVerticalBottomYOffsetPoints;
        double verticalTopY = baselineY + WordCompatibleAllMarkupCommentRangeVerticalTopYOffsetPoints;
        double topTickY = baselineY + WordCompatibleAllMarkupCommentRangeTopTickYOffsetPoints;
        graphics.SetFillRgb(WordCompatibleAllMarkupReviewFillRgb.Red, WordCompatibleAllMarkupReviewFillRgb.Green, WordCompatibleAllMarkupReviewFillRgb.Blue);
        graphics.FillRectangle(
            Math.Max(0d, startX - WordCompatibleAllMarkupCommentRangeFillXInsetPoints),
            baselineY + WordCompatibleAllMarkupCommentRangeFillBaselineYOffsetPoints,
            Math.Max(0d, endX - startX),
            WordCompatibleAllMarkupCommentRangeFillHeightPoints);
        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        graphics.SetLineWidth(WordCompatibleAllMarkupCommentRangeStrokeWidthPoints);
        graphics.StrokeLine(
            startX + WordCompatibleAllMarkupCommentRangeTickLengthPoints,
            bottomTickY,
            startX + WordCompatibleAllMarkupCommentRangeFarTickLengthPoints,
            bottomTickY);
        graphics.StrokeLine(startX, verticalBottomY, startX, verticalTopY);
        graphics.StrokeLine(
            startX,
            topTickY,
            startX + WordCompatibleAllMarkupCommentRangeTickLengthPoints,
            topTickY);
        graphics.StrokeLine(
            endX - WordCompatibleAllMarkupCommentRangeFarTickLengthPoints,
            bottomTickY,
            endX - WordCompatibleAllMarkupCommentRangeTickLengthPoints,
            bottomTickY);
        graphics.StrokeLine(endX, verticalBottomY, endX, verticalTopY);
        graphics.StrokeLine(
            endX - WordCompatibleAllMarkupCommentRangeFarTickLengthPoints,
            topTickY,
            endX,
            topTickY);
    }

    private static bool ShouldDrawCommentMarkerLabel(DocxMarkupContext markupContext)
    {
        return markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin;
    }

    private static string ResolveCommentMarkerLabel(DocxParagraph paragraph)
    {
        string? id = paragraph
            .InlineReferences
            .Where(reference => reference.Kind == DocxRelatedStoryKind.Comment)
            .Select(reference => reference.Id)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        return id is null ? "?" : id;
    }

    private static DocxRevisionMarkupPalette ResolveRevisionMarkupPalette(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        DocxMarkupBalloonRgb stroke = ResolveRevisionAuthorColor(revisions);
        return new DocxRevisionMarkupPalette(
            FillRgb: MixRgb(stroke, new DocxMarkupBalloonRgb(255, 255, 255), 0.88d),
            StrokeRgb: stroke,
            TitleRgb: MixRgb(stroke, new DocxMarkupBalloonRgb(0, 0, 0), 0.35d));
    }

    private static DocxMarkupBalloonRgb ResolveRevisionAuthorColor(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        int bucket = ResolveRevisionAuthorBucket(ResolveRevisionAuthorBucketKey(revisions));
        return RevisionAuthorColorPalette[bucket];
    }

    internal static (byte Red, byte Green, byte Blue) ResolveRevisionAuthorColorSnapshot(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(revisions);
        return (color.Red, color.Green, color.Blue);
    }

    private static string ResolveRevisionAuthorBucketKey(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        return revisions
            .Select(revision => NormalizeRevisionAuthorBucketKey(revision.Author))
            .Where(author => author.Length != 0)
            .GroupBy(author => author, StringComparer.Ordinal)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.Ordinal)
            .Select(group => group.Key)
            .FirstOrDefault() ?? string.Empty;
    }

    private static int ResolveRevisionAuthorBucket(string? author)
    {
        string normalized = NormalizeRevisionAuthorBucketKey(author);
        unchecked
        {
            uint hash = 2166136261u;
            foreach (char character in normalized)
            {
                hash ^= character;
                hash *= 16777619u;
            }

            return (int)(hash % (uint)RevisionAuthorColorPalette.Length);
        }
    }

    private static string NormalizeRevisionAuthorBucketKey(string? author)
    {
        string? value = FirstNonEmpty(author);
        if (value is null)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value.Trim().Normalize(NormalizationForm.FormC))
        {
            if (char.IsWhiteSpace(character))
            {
                pendingSpace = builder.Length != 0;
                continue;
            }

            if (pendingSpace)
            {
                builder.Append(' ');
                pendingSpace = false;
            }

            builder.Append(char.ToUpperInvariant(character));
        }

        return builder.ToString();
    }

    private static DocxMarkupBalloonRgb MixRgb(DocxMarkupBalloonRgb source, DocxMarkupBalloonRgb target, double targetWeight)
    {
        double sourceWeight = 1d - targetWeight;
        return new DocxMarkupBalloonRgb(
            (byte)Math.Round(source.Red * sourceWeight + target.Red * targetWeight),
            (byte)Math.Round(source.Green * sourceWeight + target.Green * targetWeight),
            (byte)Math.Round(source.Blue * sourceWeight + target.Blue * targetWeight));
    }
}
