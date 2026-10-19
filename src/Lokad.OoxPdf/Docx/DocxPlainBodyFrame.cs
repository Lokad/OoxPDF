namespace Lokad.OoxPdf.Docx;

internal static class DocxPlainBodyFrame
{
    internal static bool IsOrdinaryBody(DocxDocument document) =>
        document.RelatedStories.Count == 0 && document.Paragraphs.All(p => p.InlineReferences.Count == 0);

    internal static bool HasSingleLineParagraphs(DocxLayout layout) =>
        layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>()
            .Where(l => l.SourceParagraph is not null && l.Segments.Any(s => s.Role == DocxTextSegmentRole.Text))
            .GroupBy(l => l.SourceParagraph!, ReferenceEqualityComparer.Instance).All(g => g.Count() == 1);

    internal static bool HasSingleLineLayout(DocxDocument document, DocxLayout layout) =>
        HasSingleLineParagraphs(layout) &&
        !layout.Pages.Any(p => p.ColumnFrames.Count > 1 || p.StaticTextLines.Count != 0 || p.StaticInlineImages.Count != 0) &&
        !document.Paragraphs.Any(p => p.KeepRules.KeepNext == true || p.KeepRules.KeepLines == true);

    internal static bool IsEligible(DocxDocument document) =>
        HasPlainBody(document) && HasNoStaticStories(document);

    internal static bool HasNoStaticStories(DocxDocument document) =>
        document.HeaderParagraphs.Count == 0 && document.FooterParagraphs.Count == 0 &&
        document.HeaderParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count == 0 &&
        document.HeaderBodyElementsByType.Count == 0 && document.FooterBodyElementsByType.Count == 0 &&
        document.HeaderFloatingDrawingsByType.Count == 0 && document.FooterFloatingDrawingsByType.Count == 0;

    internal static bool HasKnownBodyFrame(DocxDocument document) =>
        document.FinalSectionBreak is { Revisions.Count: 0 } &&
        document.Settings.MirrorMargins != true && document.FloatingDrawings.Count == 0;

    internal static bool HasPlainBody(DocxDocument document) =>
        HasKnownBodyFrame(document) && document.BodyElements.Count != 0 && document.BodyElements.All(e =>
            e is DocxParagraphElement paragraph ? DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph.Paragraph, e.Revisions) :
            e is DocxPageBreakElement && e.Revisions.Count == 0) &&
        document.Paragraphs.All(IsPlainParagraph);

    internal static bool HasAutomaticLineFlow(DocxParagraph paragraph) =>
        paragraph.LineSpacingPoints is null && paragraph.Spacing.LineRuleValue is (null or "auto");

    internal static bool HasContentDrivenLineFlow(DocxParagraph paragraph) =>
        HasAutomaticLineFlow(paragraph) ||
        paragraph.LineSpacingPoints is { } minimum && double.IsFinite(minimum) && minimum >= 0d &&
        string.Equals(paragraph.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase);

    internal static bool HasFixedLineFlow(DocxParagraph paragraph) =>
        paragraph.LineSpacingPoints is { } height && double.IsFinite(height) && height > 0d &&
        string.Equals(paragraph.Spacing.LineRuleValue, "exact", StringComparison.OrdinalIgnoreCase);

    internal static bool IsPlainParagraph(DocxParagraph p) =>
        p.Alignment == DocxTextAlignment.Left && HasPlainTextContent(p);

    // Bookmarks are non-rendering source metadata. Navigation validates their
    // anchor ownership separately; they do not change the plain flow frame.
    internal static bool HasPlainTextContent(DocxParagraph p) =>
        p.Images.Count == 0 && p.InlineTextBoxes.Count == 0 && p.Equations.Count == 0 &&
        p.FieldReferences.Count == 0 && p.CommentRanges.Count == 0 && DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(p, p.Revisions) &&
        p.RevisionRanges.Count == 0 && !p.HasDeletedParagraphMark && p.DeletedText.Length == 0 &&
        p.ListLabel is null && p.Hyperlinks.Count == 0 &&
        p.Runs.All(r => r.Revision is null && r.Revisions.Count == 0 &&
            DocxScriptClassifier.IsSimpleLineText(r.Text));
}
