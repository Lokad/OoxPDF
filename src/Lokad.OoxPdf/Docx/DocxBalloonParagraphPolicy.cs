using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxBalloonParagraphPolicy
{
    internal static bool HasMixedFiveParagraphTerminal(IReadOnlyList<DocxBodyElement> elements)
    {
        if (elements.Count != 5 || elements[^1] is not DocxParagraphElement terminal) { return false; }
        DocxTextRun[] runs = terminal.Paragraph.Runs.Where(run => run.Text.Length != 0).ToArray();
        return runs.Length == 2 && (runs[0].FontFamily != runs[1].FontFamily || runs[0].Fonts != runs[1].Fonts);
    }

    internal static bool HasPlainBalloonParagraphShape(DocxParagraph paragraph) =>
        paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0 &&
        paragraph.FieldReferences.Count == 0 && paragraph.Revisions.Count == 0 &&
        paragraph.ListLabel is null && paragraph.Alignment == DocxTextAlignment.Left &&
        paragraph.Indent == DocxParagraphIndent.Empty && paragraph.TabStops.Count == 0 &&
        IsWordCompatibleBalloonParagraphSpacing(paragraph.Spacing) && paragraph.LineSpacingPoints is null;

    internal static bool CanRetainFiveParagraphMarks(IReadOnlyList<DocxBodyElement> elements)
    {
        if (elements.Count != 5) { return false; }
        int paragraphIndex = 0;
        foreach (DocxBodyElement element in elements)
        {
            if (element is not DocxParagraphElement item || !HasPlainBalloonParagraphShape(item.Paragraph) ||
                item.Paragraph.ParagraphMarkRun is not { } mark || !IsPlainBalloonParagraphRun(mark)) { return false; }
            bool mixedTerminal = paragraphIndex == 4 && item.Paragraph.Runs.Count(run => run.Text.Length != 0) == 2;
            DocxTextRun? first = null;
            foreach (DocxTextRun run in item.Paragraph.Runs)
            {
                if (run.Text.Length == 0) { continue; }
                if (!IsPlainBalloonParagraphRun(run) ||
                    (!mixedTerminal && first is not null && (run.FontFamily != first.FontFamily || run.Fonts != first.Fonts))) { return false; }
                first ??= run;
            }
            string body = string.Concat(item.Paragraph.Runs.Select(run => run.Text)).Trim(' ');
            if (first is null || body.Length == 0 || body.Contains("  ", StringComparison.Ordinal)) { return false; }
            paragraphIndex++;
        }
        return true;
    }

    internal static bool IsPlainBalloonParagraphRun(DocxTextRun run) =>
        !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
        !run.AllCaps && !run.SmallCaps && !run.Hidden && run.VerticalAlignmentValue is null &&
        run.HighlightValue is null && run.ShadingFillHex is null && run.CharacterSpacingPoints == 0d &&
        run.FieldKind is null && (run.ColorHex is null || run.ColorHex == "000000");

    internal static bool IsWordCompatibleBalloonParagraphSpacing(DocxParagraphSpacing spacing) =>
        (spacing.BeforeLinesValue is null or "100" or "200" or "300") && (spacing.AfterLinesValue is null or "100" or "200" or "300") &&
        (spacing.BeforeAutoSpacingValue is null or "0" or "1" or "true" or "false" or "on" or "off") &&
        (spacing.AfterAutoSpacingValue is null or "0" or "1" or "true" or "false" or "on" or "off") &&
        spacing.LineValue is null && spacing.LineRuleValue is null && spacing.ContextualSpacing is null &&
        IsBalloonParagraphTwipsToken(spacing.BeforeValue) && IsBalloonParagraphTwipsToken(spacing.AfterValue);

    private static bool IsBalloonParagraphTwipsToken(string? value) =>
        value is null || uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out _);
}
