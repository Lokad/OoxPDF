namespace Lokad.OoxPdf.Docx;

internal static class DocxPlainBodyFrame
{
    internal static bool IsEligible(DocxDocument document) =>
        HasPlainBody(document) &&
        document.HeaderParagraphs.Count == 0 && document.FooterParagraphs.Count == 0 &&
        document.HeaderParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count == 0 &&
        document.HeaderBodyElementsByType.Count == 0 && document.FooterBodyElementsByType.Count == 0 &&
        document.HeaderFloatingDrawingsByType.Count == 0 && document.FooterFloatingDrawingsByType.Count == 0;

    internal static bool HasPlainBody(DocxDocument document) =>
        document.FinalSectionBreak is { Revisions.Count: 0 } &&
        document.Settings.MirrorMargins != true && document.FloatingDrawings.Count == 0 &&
        document.BodyElements.Count != 0 && document.BodyElements.All(e => e.Revisions.Count == 0 &&
            e is DocxParagraphElement or DocxPageBreakElement) &&
        document.Paragraphs.All(IsPlainParagraph);

    internal static bool HasAutomaticLineFlow(DocxParagraph paragraph) =>
        paragraph.LineSpacingPoints is null && paragraph.Spacing.LineRuleValue is (null or "auto");

    internal static bool IsPlainParagraph(DocxParagraph p) =>
        p.Alignment == DocxTextAlignment.Left && p.Images.Count == 0 && p.InlineTextBoxes.Count == 0 &&
        p.FieldReferences.Count == 0 && p.CommentRanges.Count == 0 && p.Revisions.Count == 0 &&
        p.RevisionRanges.Count == 0 && !p.HasDeletedParagraphMark && p.DeletedText.Length == 0 &&
        p.ListLabel is null && p.Hyperlinks.Count == 0 && p.BookmarkAnchors.Count == 0 &&
        p.Runs.All(r => r.Revision is null && r.Revisions.Count == 0 &&
            r.Text.All(ch => ch is >= ' ' and <= '~' or '\r' or '\n'));
}
