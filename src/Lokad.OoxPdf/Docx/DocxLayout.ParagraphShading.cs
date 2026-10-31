namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Word paints after-spacing only between adjacent shaded paragraphs.
    // Tables, explicit breaks and the end of a container stop the background.
    private static bool HasFollowingParagraphShading(IReadOnlyList<DocxBodyElement> elements, int index) =>
        index + 1 < elements.Count && elements[index + 1] is DocxParagraphElement { Paragraph.Shading: not null };

    private static DocxParagraphPaintFrame? CreateParagraphShadingFrame(
        DocxParagraph paragraph, bool continuesShading, double frameX, double frameWidth, double itemX,
        double baselineInset, double lineHeight, double scale, double afterSpacing)
    {
        if (paragraph.Shading is null) return null;
        // Independent Word controls: the same bleed at 8/12/18pt, in body and
        // static stories, and within cells with different margins. Containers
        // retain their existing clips; hanging first lines widen the paragraph.
        double bleed = 1.44d * scale;
        double continuation = GetParagraphTextStartOffset(paragraph, scale);
        double left = Math.Min(continuation, paragraph.ListLabel is null
            ? GetParagraphFirstLineIndentOffset(paragraph, scale)
            : GetParagraphLabelStartOffset(paragraph, scale));
        double right = GetParagraphRightInset(paragraph, scale);
        return new(frameX + left - bleed - itemX, baselineInset,
            Math.Max(0d, frameWidth - left - right + 2d * bleed),
            Math.Max(0d, lineHeight + (continuesShading ? afterSpacing : 0d)));
    }
}
