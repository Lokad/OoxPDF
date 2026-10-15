using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static DocxMidLineMath? WrapTableInlineMath(DocxParagraph paragraph, IReadOnlyList<DocxTextSpan> spans,
        ref DocxWrappedTextLine[] lines, double firstWidth, double continuationWidth, double fontSize,
        DocxTableCellLayoutContext context)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        return context.MathParagraphs is { } prepared && prepared.TryGetValue(paragraph, out var math) && !math.Equation.Display
            ? WrapInlineMath(paragraph, math, spans, ref lines, firstWidth, continuationWidth, fontSize,
                context.TextMeasurer, context.ParagraphSpacingScale, context.DefaultTabStopPoints,
                context.PageNumber ?? 1, context.CancellationToken) : null;
    }

    private static (double Above, double Below) ResolveTableInlineMathExpansion(DocxPreparedMathParagraph math,
        DocxParagraph paragraph, IReadOnlyList<DocxParagraph> paragraphs, double fontSize, double lineHeight,
        DocxTableCellLayoutContext context)
    {
        double firstInset = ResolveTableCellFirstBaselineInset(paragraphs, context.TextMeasurer);
        double paragraphInset = DocxLineMetrics.ResolveTableCellBodyBaselineOffset(fontSize, lineHeight,
            IsExactLineSpacing(paragraph.EffectiveProperties), paragraphs, paragraph, context.TextMeasurer);
        double inset = firstInset + (paragraphInset - firstInset) * context.ParagraphSpacingScale;
        return ResolveInlineMathExpansion(math, fontSize, lineHeight, lineHeight, context.ParagraphSpacingScale,
            IsExactLineSpacing(paragraph.EffectiveProperties), baselineInset: inset);
    }

    private static bool TryGetStandaloneMath(DocxParagraph paragraph, double width, double scale,
        IReadOnlyDictionary<DocxParagraph, DocxPreparedMathParagraph>? prepared,
        [NotNullWhen(true)] out DocxPreparedMathParagraph? math)
    {
        math = null;
        if (prepared is null || !prepared.TryGetValue(paragraph, out var candidate) ||
            DocxMathParagraphPolicy.GetStandaloneRow(paragraph) is null || !FitsMathDisplayWidth(paragraph, candidate, width, scale)) { return false; }
        math = candidate; return true;
    }

    private static DocxTextLineLayout CreateStandaloneMathLine(DocxParagraph paragraph, DocxPreparedMathParagraph math,
        double x, double textWidth, double top, double scale, int paragraphIndex, DocxParagraphSpacingProfile spacing)
    {
        double inset = GetParagraphTextStartOffset(paragraph, scale);
        double width = textWidth - inset - GetParagraphRightInset(paragraph, scale);
        double boxWidth = math.Box.AdvanceWidth * scale;
        double alignment = math.Equation.AlignmentValue switch
        { "left" => 0d, "right" => width - boxWidth, _ => (width - boxWidth) / 2d };
        double size = math.Equation.ParentFontSize ?? GetParagraphFontSize(paragraph);
        var style = paragraph.ParagraphMarkRun ?? new DocxTextRun(string.Empty, size, null, false, false, false, null, null);
        return new DocxTextLineLayout(
            string.Concat(math.Box.Glyphs.Select(glyph => new System.Text.Rune(glyph.RenderCodePoint).ToString())),
            style, size, x + inset + alignment, top - math.BaselineInset * scale, boxWidth, [],
            SourceBlockIndex: null, SourceParagraphIndex: paragraphIndex, SourceLineIndex: 0,
            Story: DocxStoryId.TableCell(), LineHeight: math.LineHeight * scale,
            AppliedBeforeSpacing: spacing.AppliedBeforeSpacing, IsFirstParagraphLine: true, EndsWithIntraTokenBreak: false,
            SingleLineHeight: null, ListLabelSingleLineHeight: null, BodyWindowsLineHeight: null,
            ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null, LineSpacingFactorFloorApplied: null,
            PendingAfterSpacing: spacing.PendingAfterSpacing, ParagraphBeforeSpacing: spacing.ParagraphBeforeSpacing,
            ParagraphAfterSpacing: spacing.ParagraphAfterSpacing, ContextualSpacingSuppressed: spacing.ContextualSpacingSuppressed,
            SourceParagraph: paragraph, LineHeightSource: null, EmitsTerminalParagraphMark: false)
        { Math = math, MathScale = scale };
    }
}
