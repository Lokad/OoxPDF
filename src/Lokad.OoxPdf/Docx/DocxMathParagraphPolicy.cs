namespace Lokad.OoxPdf.Docx;

internal sealed record DocxPreparedMathParagraph(DocxEquation Equation, DocxMathBox Box, double BaselineInset, double LineHeight);

internal static class DocxMathParagraphPolicy
{
    // Initial semantic capability: a standalone, automatic display row. Mixed
    // prose, revision/reference decorations and explicit line rules keep their
    // existing fallback until those owning paths receive separate qualification.
    public static DocxEquation? GetDisplayRow(DocxParagraph paragraph)
    {
        if (!HasSupportedOwner(paragraph) || !paragraph.Equations[0].Display ||
            paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text)) ||
            paragraph.Spacing.LineValue is not null || paragraph.Spacing.LineRuleValue is not null)
        { return null; }
        DocxEquation equation = paragraph.Equations[0];
        return equation.AlignmentValue is null or "center" or "centerGroup" or "left" or "right" ? equation : null;
    }

    public static DocxEquation? GetInlineRow(DocxParagraph paragraph) =>
        HasSupportedOwner(paragraph) && !paragraph.Equations[0].Display &&
        paragraph.Spacing.LineRuleValue is null or "auto" &&
        paragraph.Runs.All(run => !run.Text.Contains('\t') && !run.Text.Contains('\n')) &&
        paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text))
            ? paragraph.Equations[0] : null;

    private static bool HasSupportedOwner(DocxParagraph paragraph) =>
        paragraph.Equations.Count == 1 && paragraph.Equations[0].Revision is null && paragraph.Revisions.Count == 0 &&
        paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0 &&
        paragraph.InlineReferences.Count == 0 && paragraph.CommentRanges.Count == 0 &&
        paragraph.FieldReferences.Count == 0 && paragraph.Hyperlinks.Count == 0 && paragraph.BookmarkAnchors.Count == 0 &&
        paragraph.RevisionRanges.Count == 0 && paragraph.ListLabel is null;
}
