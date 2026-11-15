namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool UsesSimpleSectionFootnoteAdmission(DocxDocument document, CancellationToken cancellationToken)
    {
        if (document.FinalSectionBreak is null || document.Settings.MirrorMargins == true ||
            document.FloatingDrawings.Count != 0 || !DocxPlainBodyFrame.HasNoSelectableStaticStories(document, cancellationToken) ||
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
        paragraph.Revisions.Count == 0 && !DocxInlineContent.HasContentRevisionRanges(paragraph) &&
        !paragraph.HasDeletedParagraphMark && paragraph.DeletedText.Length == 0 && paragraph.ListLabel is null;

    private static bool HasStaticSectionContent(DocxPageSettings settings) =>
        settings.HeaderParagraphsByType.Keys.Concat(settings.FooterParagraphsByType.Keys)
            .Concat(settings.HeaderBodyElementsByType.Keys).Concat(settings.FooterBodyElementsByType.Keys)
            .Concat(settings.HeaderFloatingDrawingsByType.Keys).Concat(settings.FooterFloatingDrawingsByType.Keys)
            .Any(variant => !IsStaticStoryVariantKnownDisabled(settings, variant));
}
