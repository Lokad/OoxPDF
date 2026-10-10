namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool UsesSimpleSectionPagination(DocxDocument document)
    {
        if (document.FinalSectionBreak is null || document.Settings.MirrorMargins == true ||
            document.FloatingDrawings.Count != 0 || document.HeaderParagraphs.Count != 0 || document.FooterParagraphs.Count != 0 ||
            document.HeaderParagraphsByType.Count != 0 || document.FooterParagraphsByType.Count != 0 ||
            document.HeaderBodyElementsByType.Count != 0 || document.FooterBodyElementsByType.Count != 0 ||
            document.HeaderFloatingDrawingsByType.Count != 0 || document.FooterFloatingDrawingsByType.Count != 0 ||
            document.BodyElements.Any(e => e.Revisions.Count != 0 ||
                e is not (DocxParagraphElement or DocxSectionBreakElement or DocxPageBreakElement or DocxManualBreakElement)) ||
            DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).Any(p => !IsSimpleSectionParagraph(p)))
        {
            return false;
        }

        return document.RelatedStories.All(s =>
            (s.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) &&
            s.FloatingDrawings.Count == 0 && s.Tables.Count == 0 &&
            s.Paragraphs.All(IsSimpleSectionParagraph));
    }

    private static bool IsSimpleSectionParagraph(DocxParagraph paragraph) =>
        paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0 &&
        paragraph.FieldReferences.Count == 0 && paragraph.CommentRanges.Count == 0 &&
        paragraph.Revisions.Count == 0 && paragraph.RevisionRanges.Count == 0 &&
        !paragraph.HasDeletedParagraphMark && paragraph.DeletedText.Length == 0 && paragraph.ListLabel is null;

    private static bool HasStaticSectionContent(DocxPageSettings settings) =>
        settings.HeaderParagraphsByType.Count != 0 || settings.FooterParagraphsByType.Count != 0 ||
        settings.HeaderBodyElementsByType.Count != 0 || settings.FooterBodyElementsByType.Count != 0 ||
        settings.HeaderFloatingDrawingsByType.Count != 0 || settings.FooterFloatingDrawingsByType.Count != 0;
}
