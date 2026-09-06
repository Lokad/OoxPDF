using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool IsOfficeTextOperationBoundaryPunctuation(char value)
    {
        return CharUnicodeInfo.GetUnicodeCategory(value) == UnicodeCategory.DashPunctuation;
    }

    private static double MeasureFlowSegmentBoundaryAdjustment(TextAdvanceEstimator advanceEstimator, string advanceText, int? previousCodePoint, double fontSize, TextAdvanceOptions options)
    {
        if (previousCodePoint is int previous && FirstCodePoint() is int first)
        {
            return advanceEstimator.MeasureBoundaryAdvance(previous, first, fontSize, options);
        }

        return 0d;

        int? FirstCodePoint()
        {
            foreach (Rune rune in advanceText.EnumerateRunes())
            {
                return rune.Value;
            }

            return null;
        }
    }

    private static int? LastCodePoint(string text)
    {
        int? last = null;
        foreach (Rune rune in text.EnumerateRunes())
        {
            last = rune.Value;
        }

        return last;
    }

    private static IReadOnlyList<PptxTextAtomLayout> BuildTextAtoms(TextRun run, TextAdvanceEstimator advanceEstimator, PptxTextAtomKind? forcedKind)
    {
        if (run.Text.Length == 0)
        {
            return [];
        }

        if (forcedKind is { } kind)
        {
            return [new PptxTextAtomLayout(kind, run.Text, run.X, run.Width, Draw: kind != PptxTextAtomKind.HiddenAdvance)];
        }

        var atoms = new List<PptxTextAtomLayout>();
        double cursorX = run.X;
        int index = 0;
        while (index < run.Text.Length)
        {
            int start = index;
            bool isSpace = run.Text[index] == ' ';
            while (index < run.Text.Length && (run.Text[index] == ' ') == isSpace)
            {
                index++;
            }

            string text = run.Text[start..index];
            double width = advanceEstimator.Measure(text, run.FontSize, run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
            atoms.Add(new PptxTextAtomLayout(isSpace ? PptxTextAtomKind.Space : PptxTextAtomKind.Word, text, cursorX, width, Draw: true));
            cursorX += width;
        }

        if (atoms.Count > 0)
        {
            PptxTextAtomLayout last = atoms[^1];
            double delta = run.X + run.Width - (last.X + last.Width);
            if (Math.Abs(delta) > PptxTextMetricRules.TextStateTolerance)
            {
                atoms[^1] = last with { Width = Math.Max(0d, last.Width + delta) };
            }
        }

        return atoms;
    }

    private static PptxTextGlyphSpanLayout BuildGlyphSpan(TextRun run, TextAdvanceEstimator advanceEstimator, double leadingAdjustment)
    {
        var glyphs = new List<PptxTextGlyphLayout>();
        OpenTypeFont? previousFont = null;
        ushort previousGlyph = 0;
        foreach (Rune rune in run.Text.EnumerateRunes())
        {
            ResolvedGlyphFont? resolved = advanceEstimator.ResolveGlyphFont(run.FontFamily, run.Bold, run.Italic, rune.Value);
            if (resolved is null || resolved.Font.UnitsPerEm == 0)
            {
                continue;
            }

            OpenTypeFont font = resolved.Font;
            ushort glyph = font.MapCodePoint(rune.Value);
            if (glyph == 0)
            {
                continue;
            }

            double adjustmentBefore = 0d;
            if (glyphs.Count > 0)
            {
                adjustmentBefore += run.CharacterSpacing;
                if (run.KerningEnabled && previousFont == font && previousGlyph != 0)
                {
                    adjustmentBefore += font.GetKerning(previousGlyph, glyph) * run.FontSize / font.UnitsPerEm;
                }

                if (previousFont == font && previousGlyph != 0 && resolved.SyntheticBold)
                {
                    adjustmentBefore -= PptxTextMetricRules.OfficeSyntheticBoldAdvanceTightening(run.FontSize);
                }
            }

            double advance = font.GetAdvanceWidth(glyph) * run.FontSize / font.UnitsPerEm;
            glyphs.Add(new PptxTextGlyphLayout(rune.Value, resolved.Typeface, resolved.Source, glyph, advance, adjustmentBefore));
            previousFont = font;
            previousGlyph = glyph;
        }

        if (glyphs.Count == 0)
        {
            return PptxTextGlyphSpanLayout.Empty(run);
        }

        double naturalWidth = glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore);
        return new PptxTextGlyphSpanLayout(
            run.Text,
            run.FontFamily,
            run.Bold,
            run.Italic,
            run.FontSize,
            run.CharacterSpacing,
            run.KerningEnabled,
            leadingAdjustment,
            Math.Max(0d, naturalWidth),
            run.Width,
            glyphs);
    }

    private static PptxTextLineBoxLayout CreateLineBox(
        double lineTopY,
        double baselineY,
        LineSpacing lineSpacing,
        double maxFontSize,
        TextLayoutLine line,
        TextAdvanceEstimator advanceEstimator,
        bool useOfficeBaselineFloor)
    {
        double advance = ReadLineAdvance(lineSpacing, maxFontSize);
        PptxTextLineMetrics lineMetrics = ResolvePositionedLineMetrics(lineTopY, baselineY, advance);
        PptxTextSpanLayout? baselineSpan = line.Spans.FirstOrDefault();
        ResolvedRunTextStyle? baselineStyle = baselineSpan?.SourceRun?.Style;
        double baselineFontSize = baselineSpan?.Run.FontSize ?? maxFontSize;
        PptxTextBaselineMetricLayout baselineMetric = ReadBaselineMetric(baselineFontSize, baselineStyle, advanceEstimator, useOfficeBaselineFloor, lineSpacing);
        return new PptxTextLineBoxLayout(
            lineTopY,
            baselineY,
            lineMetrics.LineAdvance,
            lineMetrics.BaselineOffset,
            maxFontSize,
            lineSpacing,
            baselineMetric);
    }

    private static PptxTextLineMetrics ResolvePositionedLineMetrics(double lineTopY, double baselineY, double lineAdvance)
    {
        return new PptxTextLineMetrics(lineTopY - baselineY, lineAdvance, "positioned-line");
    }

    private static void AddAlignedParagraphLine(
        List<PptxTextLineLayout> lines,
        TextLayoutLine line,
        PptxTextLineBoxLayout box,
        TextAlignment alignment,
        double textX,
        double textWidth,
        bool justify,
        bool distribute,
        TextAdvanceEstimator advanceEstimator)
    {
        if (line.Spans.Count == 0)
        {
            return;
        }

        double alignmentEndX = ReadAlignmentEndX(line);
        double paragraphWidth = Math.Max(0d, alignmentEndX - textX);
        bool justifyLine = justify && paragraphWidth > 0d && paragraphWidth < textWidth;
        double offset = alignment switch
        {
            TextAlignment.Center => Math.Max(0d, textWidth - paragraphWidth) / 2d,
            TextAlignment.Right => Math.Max(0d, textWidth - paragraphWidth),
            _ => 0d
        };

        if (justifyLine)
        {
            PptxTextLineLayout? justified = TryJustifyLine(line, box, textX, textWidth, advanceEstimator);
            if (justified is not null)
            {
                lines.Add(justified);
                return;
            }
        }

        if (distribute && paragraphWidth > 0d && paragraphWidth < textWidth)
        {
            PptxTextLineLayout? distributed = TryDistributeLine(line, box, textX, textWidth);
            if (distributed is not null)
            {
                lines.Add(distributed);
                return;
            }
        }

        IEnumerable<PptxTextSpanLayout> spans = line.Spans
            .Select(span => span with
            {
                Run = span.Run with
                {
                    X = span.Run.X + offset,
                    Width = span.Run.Width,
                    Alignment = TextAlignment.Left
                },
                EndX = span.EndX + offset,
                Atoms = OffsetAtoms(span.Atoms, offset)
            });
        if (IsWordJustifiedAlignment(alignment))
        {
            spans = spans.SelectMany(span => SplitJustifiedWordSpans(span, advanceEstimator));
        }

        lines.Add(new PptxTextLineLayout(box, textX + offset, line.EndX + offset, line.EndX + offset, alignment, spans.ToArray()));
    }

    private static double ReadAlignmentEndX(TextLayoutLine line)
    {
        for (int spanIndex = line.Spans.Count - 1; spanIndex >= 0; spanIndex--)
        {
            PptxTextSpanLayout span = line.Spans[spanIndex];
            for (int atomIndex = span.Atoms.Count - 1; atomIndex >= 0; atomIndex--)
            {
                PptxTextAtomLayout atom = span.Atoms[atomIndex];
                if (!atom.Draw || atom.Kind == PptxTextAtomKind.HiddenAdvance || atom.Kind == PptxTextAtomKind.Space)
                {
                    continue;
                }

                return atom.X + atom.Width;
            }
        }

        return line.EndX;
    }

    private static IReadOnlyList<PptxTextAtomLayout> OffsetAtoms(IReadOnlyList<PptxTextAtomLayout> atoms, double offset)
    {
        if (Math.Abs(offset) <= PptxTextMetricRules.TextStateTolerance)
        {
            return atoms;
        }

        return atoms.Select(atom => atom with { X = atom.X + offset }).ToArray();
    }

    private static bool IsWordJustifiedAlignment(TextAlignment alignment)
    {
        return alignment is TextAlignment.Justify or TextAlignment.JustLow or TextAlignment.ThaiDistributed;
    }

    private static PptxTextLineLayout? TryJustifyLine(TextLayoutLine line, PptxTextLineBoxLayout box, double textX, double textWidth, TextAdvanceEstimator advanceEstimator)
    {
        double drawableEndX = ReadAlignmentEndX(line);
        int spaceCount = CountStretchableJustificationSpaces(line, drawableEndX);
        if (spaceCount == 0)
        {
            return null;
        }

        double extraWidth = textWidth - Math.Max(0d, drawableEndX - textX);
        if (extraWidth <= PptxTextMetricRules.TextStateTolerance)
        {
            return null;
        }

        double extraPerSpace = extraWidth / spaceCount;
        double shift = 0d;
        var spans = new List<PptxTextSpanLayout>(line.Spans.Count);
        foreach (PptxTextSpanLayout span in line.Spans)
        {
            int spanSpaces = span.Run.Text.Count(static c => c == ' ');
            double spanExtra = spanSpaces * extraPerSpace;
            PptxTextSpanLayout justifiedSpan = span with
            {
                Run = span.Run with
                {
                    X = span.Run.X + shift,
                    Width = span.Run.Width + spanExtra,
                    Alignment = TextAlignment.Left,
                    PreventCoalesce = true
                },
                EndX = span.EndX + shift + spanExtra,
                Atoms = JustifyAtoms(span.Atoms, shift, extraPerSpace),
                GlyphSpan = span.GlyphSpan with { LayoutWidth = span.GlyphSpan.LayoutWidth + spanExtra }
            };
            foreach (PptxTextSpanLayout wordSpan in SplitJustifiedWordSpans(justifiedSpan, advanceEstimator))
            {
                spans.Add(wordSpan);
            }

            shift += spanExtra;
        }

        return new PptxTextLineLayout(box, textX, textX + textWidth, line.EndX, TextAlignment.Justify, spans);
    }

    private static int CountStretchableJustificationSpaces(TextLayoutLine line, double drawableEndX)
    {
        int count = 0;
        foreach (PptxTextSpanLayout span in line.Spans)
        {
            foreach (PptxTextAtomLayout atom in span.Atoms)
            {
                if (!atom.Draw ||
                    atom.Kind != PptxTextAtomKind.Space ||
                    atom.X >= drawableEndX - PptxTextMetricRules.TextStateTolerance)
                {
                    continue;
                }

                count += atom.Text.Count(static c => c == ' ');
            }
        }

        return count;
    }

    private static PptxTextLineLayout? TryDistributeLine(TextLayoutLine line, PptxTextLineBoxLayout box, double textX, double textWidth)
    {
        int glyphCount = line.Spans.Sum(span => span.GlyphSpan.Glyphs.Count);
        if (glyphCount <= 1)
        {
            return null;
        }

        double extraWidth = textWidth - Math.Max(0d, line.EndX - textX);
        if (extraWidth <= PptxTextMetricRules.TextStateTolerance)
        {
            return null;
        }

        double extraPerGlyphGap = extraWidth / (glyphCount - 1);
        double cursor = textX;
        int globalGlyphIndex = 0;
        var spans = new List<PptxTextSpanLayout>(glyphCount);
        foreach (PptxTextSpanLayout span in line.Spans)
        {
            TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(span.Run.Text);
            foreach (PptxTextGlyphLayout glyph in span.GlyphSpan.Glyphs)
            {
                if (!elements.MoveNext())
                {
                    return null;
                }

                string text = elements.GetTextElement();
                cursor += glyph.AdjustmentBefore;
                TextRun glyphRun = span.Run with
                {
                    Text = text,
                    X = cursor,
                    Width = glyph.Advance,
                    CharacterSpacing = 0d,
                    Alignment = TextAlignment.Left,
                    PreventCoalesce = true
                };
                var glyphSpan = new PptxTextGlyphSpanLayout(
                    text,
                    span.GlyphSpan.Typeface,
                    span.GlyphSpan.Bold,
                    span.GlyphSpan.Italic,
                    span.GlyphSpan.FontSize,
                    0d,
                    span.GlyphSpan.KerningEnabled,
                    0d,
                    glyph.Advance,
                    glyph.Advance,
                    [glyph with { AdjustmentBefore = 0d }]);
                spans.Add(span with
                {
                    Run = glyphRun,
                    EndX = cursor + glyph.Advance,
                    Atoms = [new PptxTextAtomLayout(char.IsWhiteSpace(text, 0) ? PptxTextAtomKind.Space : PptxTextAtomKind.Word, text, cursor, glyph.Advance, Draw: true)],
                    GlyphSpan = glyphSpan
                });

                cursor += glyph.Advance;
                if (++globalGlyphIndex < glyphCount)
                {
                    cursor += extraPerGlyphGap;
                }
            }
        }

        return new PptxTextLineLayout(box, textX, textX + textWidth, line.EndX, TextAlignment.Distributed, spans);
    }

    private static IEnumerable<PptxTextSpanLayout> SplitJustifiedWordSpans(PptxTextSpanLayout span, TextAdvanceEstimator advanceEstimator)
    {
        PptxTextAtomLayout[] words = span.Atoms
            .Where(static atom => atom.Kind == PptxTextAtomKind.Word && atom.Draw && atom.Text.Length != 0)
            .ToArray();
        if (words.Length == 0)
        {
            if (span.Atoms.Any(static atom => atom.Kind is not PptxTextAtomKind.Space and not PptxTextAtomKind.HiddenAdvance))
            {
                yield return span;
            }

            yield break;
        }

        foreach (PptxTextAtomLayout word in words
                     .SelectMany(word => SplitWordAtomOnSpaces(span.Run, word, advanceEstimator))
                     .SelectMany(word => SplitJustifiedWordAtomOnSentencePeriod(span.Run, word, advanceEstimator)))
        {
            TextRun wordRun = span.Run with
            {
                Text = word.Text,
                X = word.X,
                Width = word.Width,
                PreventCoalesce = true
            };
            yield return span with
            {
                Run = wordRun,
                EndX = word.X + word.Width,
                Atoms = [word],
                GlyphSpan = BuildGlyphSpan(wordRun, advanceEstimator, 0d)
            };
        }
    }

    private static IEnumerable<PptxTextAtomLayout> SplitJustifiedWordAtomOnSentencePeriod(TextRun run, PptxTextAtomLayout atom, TextAdvanceEstimator advanceEstimator)
    {
        if (atom.Text.Length <= 1 || atom.Text[^1] != '.')
        {
            yield return atom;
            yield break;
        }

        string wordText = atom.Text[..^1];
        double wordWidth = advanceEstimator.Measure(wordText, run.FontSize, run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
        double periodBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(
            advanceEstimator,
            ".",
            LastCodePoint(wordText),
            run.FontSize,
            new TextAdvanceOptions(run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled));
        double periodX = atom.X + wordWidth + periodBoundaryAdjustment;
        yield return atom with { Text = wordText, Width = Math.Max(0d, periodX - atom.X) };
        yield return atom with { Text = ".", X = periodX, Width = Math.Max(0d, atom.X + atom.Width - periodX) };
    }

    private static IEnumerable<PptxTextAtomLayout> SplitWordAtomOnSpaces(TextRun run, PptxTextAtomLayout atom, TextAdvanceEstimator advanceEstimator)
    {
        if (atom.Text.IndexOf(' ') < 0)
        {
            yield return atom;
            yield break;
        }

        double cursorX = atom.X;
        int index = 0;
        while (index < atom.Text.Length)
        {
            int start = index;
            bool isSpace = atom.Text[index] == ' ';
            while (index < atom.Text.Length && (atom.Text[index] == ' ') == isSpace)
            {
                index++;
            }

            string text = atom.Text[start..index];
            double width = advanceEstimator.Measure(text, run.FontSize, run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
            if (!isSpace)
            {
                yield return atom with
                {
                    Kind = PptxTextAtomKind.Word,
                    Text = text,
                    X = cursorX,
                    Width = width,
                    Draw = true
                };
            }

            cursorX += width;
        }
    }

    private static IReadOnlyList<PptxTextAtomLayout> JustifyAtoms(IReadOnlyList<PptxTextAtomLayout> atoms, double initialShift, double extraPerSpace)
    {
        double shift = initialShift;
        var justified = new PptxTextAtomLayout[atoms.Count];
        for (int i = 0; i < atoms.Count; i++)
        {
            PptxTextAtomLayout atom = atoms[i];
            double extra = atom.Kind == PptxTextAtomKind.Space
                ? atom.Text.Count(static c => c == ' ') * extraPerSpace
                : 0d;
            justified[i] = atom with
            {
                X = atom.X + shift,
                Width = atom.Width + extra
            };
            shift += extra;
        }

        return justified;
    }
}
