using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static DocxMidLineMath? WrapTableInlineMath(DocxParagraph paragraph, IReadOnlyList<DocxTextSpan> spans,
        ref DocxWrappedTextLine[] lines, double firstWidth, double continuationWidth, double fontSize,
        DocxTableCellLayoutContext context)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        return context.TextMeasurer is { } measurer &&
            TryGetSinglePreparedMath(paragraph, context.MathEquations, out var math) && !math.Equation.Display
            ? WrapInlineMath(paragraph, math, spans, ref lines, firstWidth, continuationWidth, fontSize,
                measurer, context.ParagraphSpacingScale, context.DefaultTabStopPoints,
                context.PageNumber ?? 1, context.CancellationToken) : null;
    }

    private static (double Above, double Below) ResolveTableInlineMathExpansion(DocxPreparedMathParagraph math,
        double fontSize, double lineHeight, bool exactSpacing, double firstInset, double paragraphInset, double scale)
    {
        double inset = firstInset + (paragraphInset - firstInset) * scale;
        return ResolveInlineMathExpansion(math, fontSize, lineHeight, lineHeight, scale, exactSpacing, baselineInset: inset);
    }

    private static bool TryGetSinglePreparedMath(DocxParagraph paragraph,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? prepared,
        [NotNullWhen(true)] out DocxPreparedMathParagraph? math)
    {
        math = null;
        return paragraph.Equations.Count == 1 && prepared is not null &&
            (DocxMathParagraphPolicy.GetStandaloneRow(paragraph) is not null || DocxMathParagraphPolicy.GetInlineRow(paragraph) is not null) &&
            prepared.TryGetValue(paragraph.Equations[0], out math);
    }

    private static bool TryGetStandaloneMath(DocxParagraph paragraph, double width, double scale,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? prepared,
        [NotNullWhen(true)] out DocxPreparedMathParagraph? math)
    {
        math = null;
        if (!TryGetSinglePreparedMath(paragraph, prepared, out var candidate) ||
            DocxMathParagraphPolicy.GetStandaloneRow(paragraph) is null || !FitsMathDisplayWidth(paragraph, candidate, width, scale)) { return false; }
        math = candidate; return true;
    }

    private static DocxTextLineLayout CreateStandaloneMathLine(DocxParagraph paragraph, DocxPreparedMathParagraph math,
        double x, double textWidth, double top, double scale, int paragraphIndex, DocxParagraphSpacingProfile spacing, bool continuesShading,
        double beforeSpacing, double baselineAnchor = 0d)
    {
        double inset = GetParagraphTextStartOffset(paragraph, scale);
        double width = textWidth - inset - GetParagraphRightInset(paragraph, scale);
        double boxWidth = math.Box.AdvanceWidth * scale;
        double alignment = math.Equation.EffectiveAlignmentValue switch
        { "left" => 0d, "right" => width - boxWidth, _ => (width - boxWidth) / 2d };
        double size = math.Equation.ParentFontSize ?? GetParagraphFontSize(paragraph);
        var style = paragraph.ParagraphMarkRun ?? new DocxTextRun(string.Empty, size, null, false, false, false, null, null);
        // An explicit nominal anchor scales only the transition to the math
        // baseline. Callers with an already physical top keep the zero default.
        double baseline = top - baselineAnchor * (1d - scale) - math.BaselineInset * scale;
        return new DocxTextLineLayout(
            string.Concat(math.Box.Glyphs.Select(glyph => new System.Text.Rune(glyph.RenderCodePoint).ToString())),
            style, size, x + inset + alignment, baseline, boxWidth, [],
            SourceBlockIndex: null, SourceParagraphIndex: paragraphIndex, SourceLineIndex: 0,
            Story: DocxStoryId.TableCell(), LineHeight: math.LineHeight * scale,
            AppliedBeforeSpacing: spacing.AppliedBeforeSpacing, IsFirstParagraphLine: true, EndsWithIntraTokenBreak: false,
            SingleLineHeight: null, ListLabelSingleLineHeight: null, BodyWindowsLineHeight: null,
            ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null, LineSpacingFactorFloorApplied: null,
            PendingAfterSpacing: spacing.PendingAfterSpacing, ParagraphBeforeSpacing: spacing.ParagraphBeforeSpacing,
            ParagraphAfterSpacing: spacing.ParagraphAfterSpacing, ContextualSpacingSuppressed: spacing.ContextualSpacingSuppressed,
            SourceParagraph: paragraph, LineHeightSource: null, EmitsTerminalParagraphMark: false)
        { Math = math, MathScale = scale, FrameOriginOffsetX = -inset - alignment,
            ParagraphShadingFrame = CreateParagraphShadingFrame(paragraph, continuesShading, x, textWidth, x + inset + alignment,
                top - baseline, math.LineHeight * scale, scale, spacing.ParagraphAfterSpacing,
                beforeSpacing) };
    }
}
