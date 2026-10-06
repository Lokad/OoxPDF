using System.Text;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IReadOnlyList<DocxMarkupBalloonBodyPart>? ResolveMixedCommentBalloonBodyParts(
        DocxRelatedStoryLayout? storyLayout,
        string body,
        DocxFontResources? fontResources,
        CancellationToken cancellationToken)
    {
        if (fontResources is null || storyLayout is null ||
            storyLayout.Story.BodyElements.Count != 1 ||
            storyLayout.Story.BodyElements[0] is not DocxParagraphElement element ||
            storyLayout.InlineImages.Count != 0 || storyLayout.FloatingDrawings.Count != 0)
        {
            return null;
        }
        DocxParagraph paragraph = element.Paragraph;
        if (paragraph.Images.Count != 0 || paragraph.InlineTextBoxes.Count != 0 ||
            paragraph.FieldReferences.Count != 0 || paragraph.Revisions.Count != 0)
        {
            return null;
        }
        var parts = new List<DocxMarkupBalloonBodyPart>();
        foreach (DocxTextRun run in paragraph.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Text.Length == 0) { continue; }
            if (run.Bold || run.Italic || run.Underline || run.Strike || run.DoubleStrike ||
                run.AllCaps || run.SmallCaps || run.Hidden || run.VerticalAlignmentValue is not null ||
                run.HighlightValue is not null || run.ShadingFillHex is not null ||
                run.CharacterSpacingPoints != 0d || run.FieldKind is not null ||
                (run.ColorHex is not null && run.ColorHex != "000000") ||
                !fontResources.RunResources.TryGetValue(run, out DocxRunFontResource? resource) ||
                resource.Resolution.Bold || resource.Resolution.Italic || resource.Resolution.IsFallback)
            {
                return null;
            }
            int index = 0;
            foreach (Rune rune in run.Text.EnumerateRunes())
            {
                if ((index++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                ushort glyph = resource.Embedded.Font.MapCodePoint(rune.Value);
                if ((Rune.IsWhiteSpace(rune) && rune.Value != ' ') || glyph == 0 ||
                    !resource.Embedded.TryGetEncodedCid(glyph, out _)) { return null; }
            }
            parts.Add(new(run.Text, resource));
        }
        if (parts.Count < 2 || parts.All(part => ReferenceEquals(part.Resource.Embedded.Font, parts[0].Resource.Embedded.Font)))
        {
            return null;
        }
        parts[0] = parts[0] with { Text = parts[0].Text.TrimStart(' ') };
        parts[^1] = parts[^1] with { Text = parts[^1].Text.TrimEnd(' ') };
        int offset = 0;
        foreach (DocxMarkupBalloonBodyPart part in parts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (part.Text.Length > body.Length - offset || !body.AsSpan(offset, part.Text.Length).SequenceEqual(part.Text)) { return null; }
            offset += part.Text.Length;
        }
        if (offset != body.Length || body.Contains("  ", StringComparison.Ordinal)) { return null; }
        return parts.ToArray();
    }

    private static bool TryMeasureWordCompatibleSingleRowParts(
        IReadOnlyList<DocxMarkupBalloonBodyPart>? parts,
        double fontSize,
        double availableWidth,
        out double width)
    {
        width = 0d;
        if (parts is null || parts.Count == 0) { return false; }
        foreach (DocxMarkupBalloonBodyPart part in parts)
        {
            width += part.Resource.Embedded.MeasureTextPoints(part.Text, fontSize);
        }
        return double.IsFinite(width) && width > 0d && width <= availableWidth;
    }
}
