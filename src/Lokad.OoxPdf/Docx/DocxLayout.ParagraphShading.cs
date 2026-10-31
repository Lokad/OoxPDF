namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Word paints after-spacing only between adjacent shaded paragraphs.
    // Tables, explicit breaks and the end of a container stop the background.
    private static bool HasFollowingParagraphShading(IReadOnlyList<DocxBodyElement> elements, int index) =>
        index + 1 < elements.Count && elements[index + 1] is DocxParagraphElement { Paragraph.Shading: not null };

    private static double ParagraphShadingLeftInset(DocxParagraph paragraph, double scale) =>
        Math.Min(GetParagraphTextStartOffset(paragraph, scale), paragraph.ListLabel is null
            ? GetParagraphFirstLineIndentOffset(paragraph, scale)
            : GetParagraphLabelStartOffset(paragraph, scale));

    private static double ResolveParagraphShadingBeforeSpacing(DocxParagraph? previous, DocxParagraph paragraph,
        DocxParagraphSpacingProfile spacing)
    {
        // Word joins the excess gap only when the background edges agree.
        // Text alignment and first-line indentation may still differ.
        if (previous?.Shading is null ||
            ParagraphShadingLeftInset(previous, 1d) != ParagraphShadingLeftInset(paragraph, 1d) ||
            GetParagraphRightInset(previous, 1d) != GetParagraphRightInset(paragraph, 1d)) return 0d;
        return Math.Max(0d, spacing.AppliedBeforeSpacing - spacing.PendingAfterSpacing);
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
