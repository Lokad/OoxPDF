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
        if (!UsesWordCompatibleAllMarkupTextProfile(markupContext) ||
            markupContext.WordCompatibleHasPrintedBalloons == false)
        {
            return;
        }

        // Review chrome follows the printed paper, independently of body margins.
        var frame = new DocxPrintedPageFrame(page.Height, markupContext.WordCompatiblePrintScale);
        double bottom = frame.MapY(0d);
        double top = frame.MapY(page.Height);
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
            string.Equals(ResolveMarkupBalloonArea(page, markupContext).Side, "Right", StringComparison.Ordinal))
        {
            // Office A/B (mirrored/dense/landscape/author probes plus the unresolved
            // reference, Word-COM rendered across four print scales): the gray lane hugs
            // balloon bodies at a 22.56pt design inset (pads 16.73/17.11/17.52/18.12 at
            // scales 0.7423/0.7575/0.7762/0.8028, all within 0.02 of the fit), so the edge
            // follows placed bodies. Unit-scale Word controls retain that same
            // physical inset when the existing page margin fits the balloons.
            double balloonBodyX = ResolveMarkupBalloonArea(page, markupContext).X;
            x = Math.Max(0d, balloonBodyX - WordCompatibleAllMarkupBalloonLaneBackgroundBalloonInsetPoints * markupContext.WordCompatiblePrintScale);
            width = Math.Max(0d, page.Width - WordCompatibleAllMarkupLaneBackgroundRightBleedPoints - x);
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
        FloatingDrawingPageIndex.PageIndexPair drawingPages,
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
        void IncludeRevisionBounds(double lower, double upper)
        {
            bottom = bottom is null ? lower : Math.Min(bottom.Value, lower);
            top = top is null ? upper : Math.Max(top.Value, upper);
        }

        void PaintRevisionBounds()
        {
            if (bottom is null || top is null)
            {
                return;
            }

            double barBottom = Math.Max(0d, bottom.Value);
            double barTop = Math.Min(page.Height, top.Value);
            double barHeight = Math.Max(0d, barTop - barBottom);
            if (barHeight > 0d)
            {
                // Word centers the marker in a narrow left margin, with its
                // distance from the body capped at half an inch. Both its X and
                // thickness belong to the printed page frame, not text styling.
                double margin = Math.Max(0d, page.MarginLeft);
                double scale = markupContext.WordCompatiblePrintScale;
                double barX = (margin - Math.Min(margin / 2d,
                    WordCompatibleAllMarkupRevisionBarMaximumTextDistancePoints)) * scale;
                graphics.SetFillRgb(0, 0, 0);
                graphics.FillRectangle(
                    barX,
                    barBottom,
                    WordCompatibleAllMarkupRevisionBarWidthPoints * scale,
                    barHeight);
            }
            bottom = null;
            top = null;
        }

        bool separateBodyBands = page.ColumnFrames.Count <= 1 &&
            !drawingPages.PageAll(pageIndex).Any(drawing => !IsStaticStoryFloatingDrawing(drawing)) &&
            page.Items.All(item => item switch
            {
                DocxTextLineLayout { SourceParagraph: not null, SourceBlockIndex: >= 0 } line =>
                    line.Story is null || line.Story?.Kind == DocxStoryKind.Body,
                DocxInlineImageLayout { SourceBlockIndex: >= 0, SourceLineIndex: >= 0 } image =>
                    image.Story is null || image.Story?.Kind == DocxStoryKind.Body,
                _ => false
            });
        if (separateBodyBands)
        {
            int? lastChangedBlock = null;
            foreach (DocxTextLineLayout line in page.Items.OfType<DocxTextLineLayout>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!HasTextLineRevision(line))
                {
                    PaintRevisionBounds();
                    lastChangedBlock = null;
                    continue;
                }

                // Word joins adjacent changed paragraphs, but leaves an unchanged
                // paragraph clear. A skipped source block can be an empty spacer
                // with no emitted text line and still separates the changed bands.
                int sourceBlock = line.SourceBlockIndex!.Value;
                if (lastChangedBlock is { } previous && sourceBlock > previous + 1)
                {
                    PaintRevisionBounds();
                }
                var bounds = ResolveRevisionLineBounds(line, markupContext);
                IncludeRevisionBounds(bounds.Bottom, bounds.Top);
                lastChangedBlock = sourceBlock;
            }
            PaintRevisionBounds();
        }

        var storyBounds = new Dictionary<DocxStoryId, (double Bottom, double Top)>();
        foreach (var source in EnumerateRenderedPageLineSources(page, drawingPages.PageAll(pageIndex), markupContext, page.Height))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var line = source.Line;
            var owner = RevisionBandOwner(source.Owner);
            if (separateBodyBands && owner.Kind == DocxStoryKind.Body || !HasTextLineRevision(line)) continue;

            var bounds = ResolveRevisionLineBounds(line, markupContext);
            IncludeStoryBounds(owner, bounds.Bottom, bounds.Top);
        }

        foreach (DocxTableRowLayout row in EnumerateMarkupBalloonTableRows(page, drawingPages.PageAll(pageIndex)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!HasTableRowRevision(row))
            {
                continue;
            }

            IncludeStoryBounds(RevisionBandOwner(row.Story ?? DocxStoryId.Body()),
                row.Y - WordCompatibleAllMarkupRevisionBarBottomOutsetPoints,
                row.Y + Math.Max(6d, row.Height) - WordCompatibleAllMarkupRevisionBarTopInsetPoints);
        }
        foreach (var bounds in storyBounds.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bottom = bounds.Bottom; top = bounds.Top;
            PaintRevisionBounds();
        }

        void IncludeStoryBounds(DocxStoryId owner, double lower, double upper)
        {
            storyBounds[owner] = storyBounds.TryGetValue(owner, out var bounds)
                ? (Math.Min(bounds.Bottom, lower), Math.Max(bounds.Top, upper)) : (lower, upper);
        }

        static DocxStoryId RevisionBandOwner(DocxStoryId owner) =>
            owner.Kind is DocxStoryKind.Header or DocxStoryKind.Footer or DocxStoryKind.Footnote or DocxStoryKind.Endnote
                ? owner : DocxStoryId.Body();

        bool HasTableRowRevision(DocxTableRowLayout row)
        {
            bool mainStory = row.Story?.Kind is null or DocxStoryKind.Body;
            bool rows = row.Revisions is { Count: > 0 } revisions
                ? revisions.Any(revision => CanPrintRevisionOwner(revision, mainStory)) : row.RevisionCount != 0;
            return rows || row.Table.Revisions?.Any(revision => CanPrintRevisionOwner(revision, mainStory)) == true;
        }

        bool HasTextLineRevision(DocxTextLineLayout line)
        {
            return CollectTextLineRevisions(line, markupContext, cancellationToken).Count != 0;
        }
    }

    private static (double Bottom, double Top) ResolveRevisionLineBounds(
        DocxTextLineLayout line, DocxMarkupContext context)
    {
        double baselineOffset = ResolveTextEmissionBaselineOffset(context);
        double after = line.ParagraphAfterSpacing ?? 0d;
        // A paragraph-mark formatting range owns the complete measured row and
        // its resolved trailing gap. Other revision owners keep their fallback.
        if (line.SourceParagraph?.ParagraphMarkRun?.Revisions.Any(revision =>
                revision.Kind == DocxRevisionKind.RunPropertiesChange) == true &&
            double.IsFinite(after) && after >= 0d &&
            DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double lower, out double upper))
            return (lower - baselineOffset - after, upper - baselineOffset);

        // Related-story LineHeight may still be a design-space slot. Prefer the
        // mapped physical frame, preserving the existing unmeasured bar insets.
        double height = Math.Max(6d, line.InlineFrame?.Height ?? line.LineHeight ??
            line.FontSize * ResolveTextEmissionFontScale(context) * 1.2d);
        double y = line.BaselineY - baselineOffset - height * 0.25d;
        return (y - WordCompatibleAllMarkupRevisionBarBottomOutsetPoints,
            y + height - WordCompatibleAllMarkupRevisionBarTopInsetPoints);
    }

    private static IReadOnlyList<DocxRevisionInfo> CollectTextLineRevisions(DocxTextLineLayout line, DocxMarkupContext markupContext,
        CancellationToken token = default)
    {
        var revisions = new List<DocxRevisionInfo>();
        if (line.SourceParagraph is { } paragraph)
        {
            bool nonRendering = false;
            if (UsesWordCompatibleAllMarkupTextProfile(markupContext) && paragraph.Revisions.Count != 0)
            {
                var classifications = markupContext.NonRenderingRevisionParagraphs;
                if (classifications is null || !classifications.TryGetValue(paragraph, out nonRendering))
                {
                    nonRendering = DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph, token: token);
                    classifications?.Add(paragraph, nonRendering);
                }
            }
            if (!nonRendering) revisions.AddRange(paragraph.Revisions);
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
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        CancellationToken cancellationToken)
    {
        DocxOccupiedLineBox? fillLineBox = TryResolveBodyCommentLineBoxes(line, paragraph, fontResources,
            markupContext, cancellationToken, out var measuredBox, out var textLineBox) ? measuredBox : null;
        foreach (DocxInlineReference reference in paragraph.InlineReferences.Where(reference => reference.Kind == DocxRelatedStoryKind.Comment))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!markupContext.IsCommentMarkupVisible(paragraph, reference)) continue;
            DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
                string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
            if (range is not null &&
                TryResolveWordCompatibleCommentRangeBounds(line, range, ResolveTextEmissionXOffset(markupContext),
                    out double startX, out double endX, out var sourceBounds))
            {
                double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
                int authorSlot = ResolveCommentAuthorSlot(markupContext.CommentAuthorPaletteSlots, reference.Id);
                var (above, below) = DocxInlineLineGeometry.VerticalExtents(line);
                PdfRectangle? fill = fillLineBox is { } box && double.IsFinite(sourceBounds.Start) &&
                    double.IsFinite(sourceBounds.End) && sourceBounds.End > sourceBounds.Start
                    ? new PdfRectangle(Math.Max(0d, sourceBounds.Start), baselineY + box.BaselineInset - box.Height,
                        sourceBounds.End - sourceBounds.Start, box.Height) : null;
                PdfRectangle? glyphs = fill is not null && textLineBox is { } textBox
                    ? new PdfRectangle(sourceBounds.Start, baselineY + textBox.BaselineInset - textBox.Height,
                        sourceBounds.End - sourceBounds.Start, textBox.Height) : null;
                RenderWordCompatibleCommentRangeMarker(startX, endX, baselineY, graphics, CommentAuthorFillRgb(authorSlot), CommentAuthorStrokeRgb(authorSlot), above, below,
                    usesExactFrame: DocxInlineLineGeometry.TryGetExactBodyLineBox(line, out _), fillRectangle: fill, glyphRectangle: glyphs);
            }
            else if ((range is null || !HasCommentRangeBounds(range)) &&
                TryResolveWordCompatibleCommentReferenceMarkerBounds(line, reference, ResolveTextEmissionXOffset(markupContext), out startX, out endX))
            {
                double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
                int authorSlot = ResolveCommentAuthorSlot(markupContext.CommentAuthorPaletteSlots, reference.Id);
                RenderWordCompatibleCommentRangeMarker(startX, endX, baselineY, graphics, CommentAuthorFillRgb(authorSlot), CommentAuthorStrokeRgb(authorSlot));
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

    // Selection consumes an authored exact slot or the natural text strut.
    // Brackets consume the whole painted line's font box, including unselected
    // runs, independently of authored minimum/exact spacing. Other containers,
    // inline objects, note markers and unmeasured lines retain their fallback.
    private static bool TryResolveBodyCommentLineBoxes(DocxTextLineLayout line, DocxParagraph paragraph,
        DocxFontResources fontResources, DocxMarkupContext markupContext, CancellationToken cancellationToken,
        out DocxOccupiedLineBox box, out DocxOccupiedLineBox? textBox)
    {
        box = default;
        textBox = null;
        if (line.Story?.Kind != DocxStoryKind.Body || line.InlineFrame is not null ||
            line.BodyLineBoxHeightPoints is not > 0d || line.BodyLineBoxBaselineInsetPoints is null ||
            paragraph.Images.Count != 0 || paragraph.InlineTextBoxes.Count != 0 || paragraph.Equations.Count != 0 ||
            paragraph.ListLabel is not null || paragraph.FieldReferences.Count != 0 ||
            paragraph.InlineReferences.Any(reference => reference.Kind != DocxRelatedStoryKind.Comment)) return false;

        bool exact = DocxInlineLineGeometry.TryGetExactBodyLineBox(line, out box);
        double height = line.SingleLineHeight ?? 0d;
        if (!exact && (line.LineHeightSource is not (DocxLineHeightSource.BodySingleLineAuto or DocxLineHeightSource.AtLeastLineSpacing) ||
            !double.IsFinite(height) || height <= 0d)) return false;
        if (fontResources.TextMeasurer is not IDocxStaticTextMetricsProvider metrics) return exact;

        double above = 0d, below = 0d;
        bool hasText = false, hasAscent = true;
        foreach (var segment in line.Segments)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (segment.Role != DocxTextSegmentRole.Text) return exact;
            if (string.IsNullOrWhiteSpace(segment.Text)) continue;
            double fontSize = segment.FontSize ?? segment.StyleRun.EffectiveProperties.FontSize;
            double scale = ResolveTextEmissionFontScale(markupContext);
            double descent = metrics.MeasureWindowsDescender(segment.StyleRun, fontSize) * scale;
            if (!double.IsFinite(descent) || descent < 0d) return exact;
            double ascent = metrics.MeasureWindowsAscender(segment.StyleRun, fontSize) * scale;
            hasAscent &= double.IsFinite(ascent) && ascent > 0d;
            above = Math.Max(above, ascent);
            below = Math.Max(below, descent);
            hasText = true;
        }
        if (hasText && hasAscent && double.IsFinite(above + below))
            textBox = new(above + below, above);
        if (exact) return true;
        if (!hasText || below > height) return false;
        box = new(height, height - below);
        return true;
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
        out double endX,
        out (double Start, double End) sourceBounds)
    {
        startX = 0d;
        endX = 0d;
        sourceBounds = default;
        if (TryResolveInlineCommentRangeBounds(line, range, out double inlineStart, out double inlineEnd))
        {
            // Ordered source extents already delimit the selected content. The
            // legacy prose reconstruction insets do not belong to this geometry.
            startX = inlineStart + wordCompatibleXOffset + WordCompatibleAllMarkupCommentRangeFillXInsetPoints;
            endX = inlineEnd + wordCompatibleXOffset + WordCompatibleAllMarkupCommentRangeFillXInsetPoints;
            sourceBounds = (inlineStart + wordCompatibleXOffset, inlineEnd + wordCompatibleXOffset);
            return endX > startX;
        }
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
        sourceBounds = (layoutStartX + wordCompatibleXOffset, layoutEndX + wordCompatibleXOffset);
        return endX > startX;
    }

    private static bool TryResolveInlineCommentRangeBounds(DocxTextLineLayout line, DocxCommentRange range,
        out double startX, out double endX)
    {
        startX = double.PositiveInfinity;
        endX = double.NegativeInfinity;
        if (range.StartInlineItemIndex is not int start || range.EndInlineItemIndex is not int end || end <= start)
            return false;
        bool startsWithProse = false, endsWithProse = false;
        foreach (var extent in line.InlineSourceExtents)
            if (extent.ItemIndex >= start && extent.ItemIndex < end)
            {
                startX = Math.Min(startX, line.X + extent.StartOffset);
                endX = Math.Max(endX, line.X + extent.EndOffset);
                startsWithProse |= extent.ItemIndex == start && !extent.IsMath;
                endsWithProse |= extent.ItemIndex == end - 1 && !extent.IsMath;
            }
        // Inter-object whitespace belongs to adjacent prose selections, while
        // a selected formula owns its visible box rather than layout padding.
        foreach (var extent in line.InlineSourceExtents)
        {
            if (startsWithProse && extent.ItemIndex == start - 1 && extent.IsMath)
                startX = Math.Min(startX, line.X + extent.EndOffset);
            if (endsWithProse && extent.ItemIndex == end && extent.IsMath)
                endX = Math.Max(endX, line.X + extent.StartOffset);
        }
        return double.IsFinite(startX) && double.IsFinite(endX) && endX > startX;
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
        PdfGraphicsBuilder graphics,
        DocxMarkupBalloonRgb fillRgb,
        DocxMarkupBalloonRgb strokeRgb,
        double objectAbove = 0d, double objectBelow = 0d, bool usesExactFrame = false,
        PdfRectangle? fillRectangle = null, PdfRectangle? glyphRectangle = null)
    {
        // Exact source boxes own both expansion and compression. Content-driven
        // inline frames retain the legacy minimum and grow for their visible ink.
        double bottomExtension = objectBelow + WordCompatibleAllMarkupCommentRangeFillBaselineYOffsetPoints;
        double topExtension = objectAbove - WordCompatibleAllMarkupCommentRangeFillHeightPoints - WordCompatibleAllMarkupCommentRangeFillBaselineYOffsetPoints;
        if (!usesExactFrame)
        {
            bottomExtension = Math.Max(0d, bottomExtension);
            topExtension = Math.Max(0d, topExtension);
        }
        // Authored spacing changes selection background, while its brackets
        // remain attached to glyphs. Content-driven inline ink grows both.
        double markerBottomExtension = usesExactFrame ? 0d : bottomExtension;
        double markerTopExtension = usesExactFrame ? 0d : topExtension;
        double bottomTickY = baselineY + WordCompatibleAllMarkupCommentRangeBottomTickYOffsetPoints - markerBottomExtension;
        double verticalBottomY = baselineY + WordCompatibleAllMarkupCommentRangeVerticalBottomYOffsetPoints - markerBottomExtension;
        double verticalTopY = baselineY + WordCompatibleAllMarkupCommentRangeVerticalTopYOffsetPoints + markerTopExtension;
        double topTickY = baselineY + WordCompatibleAllMarkupCommentRangeTopTickYOffsetPoints + markerTopExtension;
        graphics.SetFillRgb(fillRgb.Red, fillRgb.Green, fillRgb.Blue);
        PdfRectangle fill = fillRectangle ?? new(
            Math.Max(0d, startX - WordCompatibleAllMarkupCommentRangeFillXInsetPoints),
            baselineY + WordCompatibleAllMarkupCommentRangeFillBaselineYOffsetPoints - bottomExtension,
            Math.Max(0d, endX - startX),
            WordCompatibleAllMarkupCommentRangeFillHeightPoints + topExtension + bottomExtension);
        graphics.FillRectangle(fill.X, fill.Y, fill.Width, fill.Height);
        graphics.SetStrokeRgb(strokeRgb.Red, strokeRgb.Green, strokeRgb.Blue);
        graphics.SetFillRgb(strokeRgb.Red, strokeRgb.Green, strokeRgb.Blue);
        if (glyphRectangle is { } glyphs)
        {
            // Word's prose brackets keep a physical stroke across print scales.
            // Their natural font box and source endpoints do not inherit the
            // selection's authored spacing or the legacy object-marker offsets.
            const double stroke = 0.14d;
            double left = glyphs.X, right = glyphs.X + glyphs.Width;
            double bottom = glyphs.Y, top = glyphs.Y + glyphs.Height;
            graphics.SetLineWidth(stroke);
            graphics.StrokeLine(left + stroke, bottom - stroke, left, bottom);
            graphics.StrokeLine(left, bottom, left, top);
            graphics.StrokeLine(left, top, left + 2d * stroke, top + 2d * stroke);
            graphics.StrokeLine(right - stroke, bottom - stroke, right, bottom);
            graphics.StrokeLine(right, bottom, right, top);
            graphics.StrokeLine(right, top, right - 2d * stroke, top + 2d * stroke);
            return;
        }
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

    internal static string NormalizeRevisionAuthorBucketKey(string? author) => NormalizeMarkupAuthor(author, foldCase: true);

    private static string NormalizeReviewAuthorIdentity(string? author) => NormalizeMarkupAuthor(author, foldCase: false);

    private static string NormalizeMarkupAuthor(string? author, bool foldCase)
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

            builder.Append(foldCase ? char.ToUpperInvariant(character) : character);
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
