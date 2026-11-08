namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Both sides of the gap join only between shaded paragraphs with matching
    // background edges. Text alignment, first-line indent and colors may differ.
    private static bool CanJoinParagraphShading(DocxParagraph previous, DocxParagraph next) =>
        previous.Shading is not null && next.Shading is not null &&
        ParagraphShadingLeftInset(previous, 1d) == ParagraphShadingLeftInset(next, 1d) &&
        GetParagraphRightInset(previous, 1d) == GetParagraphRightInset(next, 1d);

    private static bool HasFollowingParagraphShading(IReadOnlyList<DocxBodyElement> elements, int index) =>
        index + 1 < elements.Count && elements[index] is DocxParagraphElement current &&
        elements[index + 1] is DocxParagraphElement next && CanJoinParagraphShading(current.Paragraph, next.Paragraph);

    private static double ParagraphShadingLeftInset(DocxParagraph paragraph, double scale) =>
        Math.Min(GetParagraphTextStartOffset(paragraph, scale), paragraph.ListLabel is null
            ? GetParagraphFirstLineIndentOffset(paragraph, scale)
            : GetParagraphLabelStartOffset(paragraph, scale));

    private static double ResolveParagraphShadingBeforeSpacing(DocxParagraph? previous, DocxParagraph paragraph,
        DocxParagraphSpacingProfile spacing)
    {
        if (previous is null || !CanJoinParagraphShading(previous, paragraph)) return 0d;
        return Math.Max(0d, spacing.AppliedBeforeSpacing - spacing.PendingAfterSpacing);
    }

    private static DocxTextLineLayout JoinPrintedParagraphShading(List<DocxLayoutItem> placed,
        DocxTextLineLayout line, DocxParagraphSpacingProfile spacing)
    {
        if (line.ParagraphShadingFrame is not { } paint || line.SourceParagraph is not { } paragraph ||
            line.SourceBlockIndex is not int block || placed.Count == 0 ||
            placed[^1] is not DocxTextLineLayout { ParagraphShadingFrame: { } previousPaint,
                SourceParagraph: { } previousParagraph, SourceBlockIndex: { } previousBlock,
                BodyLineBoxBaselineInsetPoints: not null } previous ||
            previous.LineHeightSource is not (DocxLineHeightSource.BodySingleLineAuto or DocxLineHeightSource.AtLeastLineSpacing) ||
            !DocxPlainBodyFrame.IsPlainParagraph(previousParagraph)) return line;
        bool sameParagraph = previousBlock == block;
        if (!sameParagraph && (previousBlock + 1 != block || !CanJoinParagraphShading(previousParagraph, paragraph))) return line;
        if (Math.Abs(previous.X + previousPaint.OffsetX - line.X - paint.OffsetX) > .000001d ||
            Math.Abs(previousPaint.Width - paint.Width) > .000001d) return line;

        // The first natural box owns its top. Connected lines fill the actual
        // gap above the next box, including minimum leading. Paragraph spacing
        // already resolved contextual suppression; an authored after-value
        // cannot paint space that placement did not consume. Different colors
        // keep their respective after/before portions of the placed gap.
        double previousAfter = sameParagraph ? 0d : spacing.PendingAfterSpacing;
        double before = sameParagraph ? 0d : ResolveParagraphShadingBeforeSpacing(previousParagraph, paragraph, spacing);
        previousPaint = previousPaint with { Height = Math.Max(0d, previousPaint.Height - previousAfter) };
        paint = paint with { BaselineInset = paint.BaselineInset - before, Height = Math.Max(0d, paint.Height - before) };
        double previousBottom = previous.BaselineY + previousPaint.BaselineInset - previousPaint.Height;
        double top = line.BaselineY + paint.BaselineInset;
        double gap = Math.Max(0d, previousBottom - top);
        double after = sameParagraph ? 0d : Math.Min(gap, Math.Min(previousAfter, spacing.AppliedBeforeSpacing));
        previousPaint = previousPaint with { Height = previousPaint.Height + after };
        paint = paint with { BaselineInset = paint.BaselineInset + gap - after, Height = paint.Height + gap - after };
        placed[^1] = previous with { ParagraphShadingFrame = previousPaint };
        return line with { ParagraphShadingFrame = paint };
    }

    private static DocxParagraphPaintFrame? CreateParagraphShadingFrame(
        DocxParagraph paragraph, bool continuesShading, double frameX, double frameWidth, double itemX,
        double baselineInset, double lineHeight, double scale, double afterSpacing, double beforeSpacing)
    {
        if (paragraph.Shading is null) return null;
        // Independent Word controls: the same bleed at 8/12/18pt, in body and
        // static stories, and within cells with different margins. Containers
        // retain their existing clips; hanging first lines widen the paragraph.
        double bleed = 1.44d * scale;
        double left = ParagraphShadingLeftInset(paragraph, scale);
        double right = GetParagraphRightInset(paragraph, scale);
        return new(frameX + left - bleed - itemX, baselineInset + beforeSpacing,
            Math.Max(0d, frameWidth - left - right + 2d * bleed),
            Math.Max(0d, lineHeight + beforeSpacing + (continuesShading ? afterSpacing : 0d)));
    }
}
