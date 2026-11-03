using static Lokad.OoxPdf.Docx.DocxBalloonParagraphPolicy;
using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IReadOnlyList<DocxMarkupBalloonParagraph>? ResolveCommentBalloonParagraphs(
        DocxRelatedStoryLayout? storyLayout,
        string preview,
        DocxFontResources? fonts,
        CancellationToken cancellationToken,
        bool allowSingleParagraph = false)
    {
        // Preserve each plain paragraph's prepared body and mark faces. The
        // measured rendering paths separately admit fitting and wrapped bodies.
        if (fonts is null || storyLayout is null || storyLayout.Story.BodyElements.Count < (allowSingleParagraph ? 1 : 2) ||
            storyLayout.InlineImages.Count != 0 || storyLayout.FloatingDrawings.Count != 0) { return null; }
        var paragraphs = new List<DocxMarkupBalloonParagraph>(storyLayout.Story.BodyElements.Count);
        foreach (DocxBodyElement item in storyLayout.Story.BodyElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item is not DocxParagraphElement element) { return null; }
            DocxParagraph paragraph = element.Paragraph;
            if (!HasPlainBalloonParagraphShape(paragraph) ||
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

    private sealed record DocxMarkupParagraphRows(
        DocxMarkupBalloonParagraph[] Paragraphs, DocxUniformBalloonRow[][] BodyRows,
        double[] Pitches, double[] Transitions)
    {
        public double ContinuationsHeight => Enumerable.Range(0, Paragraphs.Length)
            .Sum(index => (BodyRows[index].Length - 1) * Pitches[index]) + Transitions.Sum();
    }

    private static DocxMarkupParagraphRows? ResolveWordCompatibleParagraphRows(
        IReadOnlyList<DocxMarkupBalloonParagraph>? paragraphs, double fontSize,
        double firstWidth, double continuationWidth, CancellationToken cancellationToken)
    {
        if (paragraphs is null || paragraphs.Count < 2 || !double.IsFinite(firstWidth) || firstWidth <= 0d ||
            !double.IsFinite(continuationWidth) || continuationWidth <= 0d) { return null; }
        var bodies = new DocxUniformBalloonRow[paragraphs.Count][];
        var pitches = new double[paragraphs.Count];
        var transitions = new double[paragraphs.Count - 1];
        for (int index = 0; index < paragraphs.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxMarkupBalloonParagraph paragraph = paragraphs[index];
            var font = paragraph.Body.Resource.Embedded.Font;
            double width = paragraph.Body.Resource.Embedded.MeasureTextPoints(paragraph.Body.Text, fontSize);
            if (font.UnitsPerEm <= 0 || !double.IsFinite(width) || width <= 0d) { return null; }
            int scalarIndex = 0;
            foreach (Rune rune in paragraph.Body.Text.EnumerateRunes())
            {
                if ((scalarIndex++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                double scalarWidth = font.GetAdvanceWidth(font.MapCodePoint(rune.Value)) * fontSize / font.UnitsPerEm;
                if (!double.IsFinite(scalarWidth) || scalarWidth > continuationWidth) { return null; }
            }
            double pitch = DocxLineMetrics.MeasureHheaLineHeight(font, fontSize);
            if (!double.IsFinite(pitch) || pitch <= 0d) { return null; }
            pitches[index] = pitch;
            bodies[index] = WrapUniformBalloonWords(paragraph.Body.Text, paragraph.Body.Resource.Embedded, fontSize,
                index == 0 ? firstWidth : continuationWidth, continuationWidth, cancellationToken,
                reserveFirstRowBreakSpace: false);
            if (bodies[index].Length == 0 || bodies[index][0].Text.Length == 0) { return null; }
            if (index > 0)
            {
                double? transition = MeasureWordCompatibleParagraphTransition(paragraphs[index - 1], paragraph, fontSize);
                if (transition is null) { return null; }
                transitions[index - 1] = transition.Value;
            }
        }
        var result = new DocxMarkupParagraphRows(paragraphs.ToArray(), bodies, pitches, transitions);
        return double.IsFinite(result.ContinuationsHeight) ? result : null;
    }

    private static void RenderWordCompatibleParagraphRows(
        DocxMarkupParagraphRows rows, DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        double firstX, double continuationX, double firstY, double fontSize, CancellationToken cancellationToken)
    {
        double y = firstY;
        for (int index = 0; index < rows.Paragraphs.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxMarkupBalloonParagraph paragraph = rows.Paragraphs[index];
            RenderUniformBalloonRows(rows.BodyRows[index], placement, graphics, paragraph.Body.Resource, paragraph.Mark,
                index == 0 ? firstX : continuationX, continuationX, y, rows.Pitches[index], fontSize, cancellationToken);
            y -= (rows.BodyRows[index].Length - 1) * rows.Pitches[index];
            if (index < rows.Transitions.Length) { y -= rows.Transitions[index]; }
        }
    }


    private static double? MeasureWordCompatibleParagraphTransition(
        DocxMarkupBalloonParagraph previous, DocxMarkupBalloonParagraph next, double fontSize)
    {
        var first = previous.Body.Resource.Embedded.Font;
        var second = next.Body.Resource.Embedded.Font;
        if (first.UnitsPerEm <= 0 || second.UnitsPerEm <= 0) { return null; }
        double gap = (-first.Hhea.HorizontalDescender / (double)first.UnitsPerEm +
            (second.Hhea.HorizontalAscender + second.Hhea.HorizontalLineGap) / (double)second.UnitsPerEm) * fontSize;
        return double.IsFinite(gap) && gap > 0d ? gap : null;
    }

}
