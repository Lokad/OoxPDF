using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IReadOnlyList<DocxMarkupBalloonParagraph>? ResolveCommentBalloonParagraphs(
        DocxRelatedStoryLayout? storyLayout,
        string preview,
        DocxFontResources? fonts,
        CancellationToken cancellationToken)
    {
        // Office prints each plain paragraph on its own row, with the paragraph
        // mark's prepared face. Wider paragraphs remain on the existing path.
        if (fonts is null || storyLayout is null || storyLayout.Story.BodyElements.Count != 2 ||
            storyLayout.InlineImages.Count != 0 || storyLayout.FloatingDrawings.Count != 0) { return null; }
        var paragraphs = new List<DocxMarkupBalloonParagraph>(2);
        foreach (DocxBodyElement item in storyLayout.Story.BodyElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item is not DocxParagraphElement element) { return null; }
            DocxParagraph paragraph = element.Paragraph;
            if (paragraph.Images.Count != 0 || paragraph.InlineTextBoxes.Count != 0 ||
                paragraph.FieldReferences.Count != 0 || paragraph.Revisions.Count != 0 ||
                paragraph.ListLabel is not null || paragraph.Alignment != DocxTextAlignment.Left ||
                paragraph.Indent != DocxParagraphIndent.Empty || paragraph.TabStops.Count != 0 ||
                paragraph.Spacing != DocxParagraphSpacing.Empty || paragraph.LineSpacingPoints is not null ||
                paragraph.ParagraphMarkRun is not { } mark || !IsPlainBalloonParagraphRun(mark) ||
                !fonts.RunResources.TryGetValue(mark, out DocxRunFontResource? markResource) ||
                markResource.Resolution.Bold || markResource.Resolution.Italic || markResource.Resolution.IsFallback ||
                !HasEncodedBalloonParagraphText(markResource, " ", cancellationToken)) { return null; }
            DocxRunFontResource? bodyResource = null;
            var text = new StringBuilder();
            foreach (DocxTextRun run in paragraph.Runs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (run.Text.Length == 0) { continue; }
                if (!IsPlainBalloonParagraphRun(run) ||
                    !fonts.RunResources.TryGetValue(run, out DocxRunFontResource? resource) ||
                    resource.Resolution.Bold || resource.Resolution.Italic || resource.Resolution.IsFallback ||
                    (bodyResource is not null && !ReferenceEquals(bodyResource.Embedded.Font, resource.Embedded.Font)) ||
                    !HasEncodedBalloonParagraphText(resource, run.Text, cancellationToken)) { return null; }
                bodyResource = resource;
                text.Append(run.Text);
            }
            string body = text.ToString().Trim(' ');
            if (bodyResource is null || body.Length == 0 || body.Contains("  ", StringComparison.Ordinal)) { return null; }
            paragraphs.Add(new(new(body, bodyResource), markResource));
        }
        return string.Join(" ", paragraphs.Select(paragraph => paragraph.Body.Text)) == preview ? paragraphs : null;
    }

    private static bool IsPlainBalloonParagraphRun(DocxTextRun run) =>
        !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
        !run.AllCaps && !run.SmallCaps && !run.Hidden && run.VerticalAlignmentValue is null &&
        run.HighlightValue is null && run.ShadingFillHex is null && run.CharacterSpacingPoints == 0d &&
        run.FieldKind is null && (run.ColorHex is null || run.ColorHex == "000000");

    private static bool HasEncodedBalloonParagraphText(DocxRunFontResource resource, string text, CancellationToken cancellationToken)
    {
        int index = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            if ((index++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            ushort glyph = resource.Embedded.Font.MapCodePoint(rune.Value);
            if ((Rune.IsWhiteSpace(rune) && rune.Value != ' ') || glyph == 0 ||
                !resource.Embedded.TryGetEncodedCid(glyph, out _)) { return false; }
        }
        return true;
    }

    private static double? ResolveWordCompatibleParagraphGap(
        IReadOnlyList<DocxMarkupBalloonParagraph>? paragraphs, double fontSize,
        double firstWidth, double continuationWidth)
    {
        if (paragraphs is not { Count: 2 }) { return null; }
        for (int index = 0; index < 2; index++)
        {
            DocxMarkupBalloonBodyPart body = paragraphs[index].Body;
            double width = body.Resource.Embedded.MeasureTextPoints(body.Text, fontSize);
            if (!double.IsFinite(width) || width <= 0d || width > (index == 0 ? firstWidth : continuationWidth)) { return null; }
        }
        var first = paragraphs[0].Body.Resource.Embedded.Font;
        var second = paragraphs[1].Body.Resource.Embedded.Font;
        if (first.UnitsPerEm <= 0 || second.UnitsPerEm <= 0) { return null; }
        double gap = (-first.Hhea.HorizontalDescender / (double)first.UnitsPerEm +
            (second.Hhea.HorizontalAscender + second.Hhea.HorizontalLineGap) / (double)second.UnitsPerEm) * fontSize;
        return double.IsFinite(gap) && gap > 0d ? gap : null;
    }

    private static void RenderWordCompatibleParagraphs(
        IReadOnlyList<DocxMarkupBalloonParagraph> paragraphs, DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics, double firstX, double continuationX,
        double firstY, double gap, double fontSize, CancellationToken cancellationToken)
    {
        for (int index = 0; index < 2; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxMarkupBalloonParagraph paragraph = paragraphs[index];
            double x = index == 0 ? firstX : continuationX;
            double y = firstY - index * gap;
            DrawBalloonText(graphics, paragraph.Body.Resource, paragraph.Body.Text, x, y, fontSize,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            DrawBalloonText(graphics, paragraph.Mark, " ", x + paragraph.Body.Resource.Embedded.MeasureTextPoints(paragraph.Body.Text, fontSize),
                y, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        }
    }
}
