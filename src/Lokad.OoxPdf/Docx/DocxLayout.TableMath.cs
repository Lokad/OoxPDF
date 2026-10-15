using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool TryGetTableDisplayMath(DocxParagraph paragraph, double width, double scale,
        IReadOnlyDictionary<DocxParagraph, DocxPreparedMathParagraph>? prepared,
        [NotNullWhen(true)] out DocxPreparedMathParagraph? math)
    {
        math = null;
        if (prepared is null || !prepared.TryGetValue(paragraph, out var candidate) ||
            !candidate.Equation.Display || !FitsMathDisplayWidth(paragraph, candidate, width, scale)) { return false; }
        math = candidate; return true;
    }

    private static DocxTextLineLayout CreateTableMathLine(DocxParagraph paragraph, DocxPreparedMathParagraph math,
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
