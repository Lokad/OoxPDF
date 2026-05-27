using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static IReadOnlyList<PptxTextGlyphRunSnapshot> InspectTextGlyphRuns(PptxDocument document, OoxPackage package, int slideIndex)
    {
        IReadOnlyList<PptxPositionedTextSpan> textSpans = ReadSlideTextSpansForInspection(document, package, slideIndex);
        RenderedFonts renderedFonts = CreateRenderedFonts(textSpans, []);
        textSpans = SplitLeadingSpacesAtHighlightBoundaries(textSpans);
        textSpans = CoalesceAdjacentTextSpans(textSpans, compareHighlight: true);
        textSpans = CoalesceUnderlineSpans(textSpans);
        var glyphRuns = new List<PptxTextGlyphRunSnapshot>();
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            foreach (PptxPositionedTextSpan emissionSpan in SplitSpanByGlyphTypeface(span))
            {
                TextRun run = emissionSpan.Run;
                if (!renderedFonts.Fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
                {
                    continue;
                }

                TextGlyphRun? glyphRun = BuildTextGlyphRun(rendered.ResourceName, rendered.Font, emissionSpan, rendered.SyntheticBold, rendered.SyntheticItalic);
                if (glyphRun is null)
                {
                    continue;
                }

                glyphRuns.Add(new PptxTextGlyphRunSnapshot(
                    run.Text,
                    glyphRun.X,
                    glyphRun.BaselineY,
                    glyphRun.Width,
                    run.FontSize,
                    glyphRun.PdfFontSize,
                    glyphRun.Glyphs.Count,
                    glyphRun.Glyphs.Skip(1).FirstOrDefault()?.AdjustmentBefore ?? 0d,
                    glyphRun.Glyphs
                        .Select(glyph => new PptxTextGlyphRunAtomSnapshot(
                            glyph.CodePoint,
                            glyph.Typeface,
                            glyphRun.ResourceName,
                            glyph.GlyphId,
                            glyph.Advance,
                            glyph.AdjustmentBefore))
                        .ToArray()));
            }
        }

        return glyphRuns;
    }

    private static IReadOnlyList<PdfFontResource> RenderTextRuns(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, string resourcePrefix = "F")
    {
        if (textRuns.Count == 0)
        {
            return [];
        }

        RenderedFonts renderedFonts = CreateRenderedFonts(textRuns, resourcePrefix);
        DrawTextRunsWithFonts(textRuns, graphics, renderedFonts.Fonts);
        return renderedFonts.Resources;
    }

    private static IReadOnlyList<PdfFontResource> RenderPositionedTextSpans(
        IReadOnlyList<PptxPositionedTextSpan> textSpans,
        IReadOnlyList<TextRun> legacyTextRuns,
        PdfGraphicsBuilder graphics)
    {
        if (textSpans.Count == 0 && legacyTextRuns.Count == 0)
        {
            return [];
        }

        RenderedFonts renderedFonts = CreateRenderedFonts(textSpans, legacyTextRuns);
        DrawTextSpansWithFonts(textSpans, graphics, renderedFonts.Fonts);
        DrawTextRunsWithFonts(legacyTextRuns, graphics, renderedFonts.Fonts);
        return renderedFonts.Resources;
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<PptxPositionedTextSpan> textSpans, IReadOnlyList<TextRun> legacyTextRuns, string resourcePrefix = "F")
    {
        var uses = new List<TextFontUse>();
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            foreach (IGrouping<string, PptxTextGlyphLayout> group in span.GlyphSpan.Glyphs.GroupBy(
                         glyph => string.IsNullOrWhiteSpace(glyph.Typeface) ? (span.Run.FontFamily ?? "Arial") : glyph.Typeface!,
                         StringComparer.OrdinalIgnoreCase))
            {
                uses.Add(new TextFontUse(group.Key, span.GlyphSpan.Bold, span.GlyphSpan.Italic, group.Select(glyph => glyph.CodePoint).ToArray()));
            }
        }

        foreach (TextRun run in CoalesceUnderlineRuns(CoalesceAdjacentTextRuns(legacyTextRuns, compareHighlight: false)))
        {
            string familyName = string.IsNullOrWhiteSpace(run.FontFamily) ? "Arial" : run.FontFamily!;
            uses.Add(new TextFontUse(familyName, run.Bold, run.Italic, run.Text.EnumerateRunes().Select(rune => rune.Value).ToArray()));
        }

        return CreateRenderedFonts(uses, resourcePrefix);
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextRun> textRuns, string resourcePrefix = "F")
    {
        if (textRuns.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }

        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        return CreateRenderedFonts(textRuns
            .Select(run => new TextFontUse(
                string.IsNullOrWhiteSpace(run.FontFamily) ? "Arial" : run.FontFamily!,
                run.Bold,
                run.Italic,
                run.Text.EnumerateRunes().Select(rune => rune.Value).ToArray()))
            .ToArray(), resourcePrefix);
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextFontUse> uses, string resourcePrefix)
    {
        if (uses.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }

        var resolver = new WindowsFontResolver();
        var fonts = new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase);
        var resources = new List<PdfFontResource>();
        foreach (IGrouping<string, TextFontUse> group in uses.GroupBy(use => FontKey(use.FamilyName, use.Bold, use.Italic), StringComparer.OrdinalIgnoreCase))
        {
            TextFontUse first = group.First();
            FontResolution resolution = resolver.ResolvePresentationTextFace(new FontRequest(first.FamilyName, first.Bold, first.Italic));
            if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
            {
                continue;
            }

            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, group.SelectMany(use => use.CodePoints));
            string resourceName = resourcePrefix + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            fonts[group.Key] = new RenderedFont(resourceName, embedded, first.Bold && !resolution.Bold, first.Italic && !resolution.Italic);
            resources.Add(new PdfFontResource(resourceName, embedded));
        }

        return new RenderedFonts(fonts, resources);
    }

    private static void DrawTextRunsWithFonts(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        DrawHighlightRunsWithFonts(textRuns, graphics, fonts);
        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        foreach (TextRun run in textRuns)
        {
            if (fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                DrawWrappedRun(graphics, rendered.ResourceName, rendered.Font, run, rendered.SyntheticBold, rendered.SyntheticItalic);
            }
        }
    }

    private static void DrawTextSpansWithFonts(IReadOnlyList<PptxPositionedTextSpan> textSpans, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        DrawHighlightSpansWithFonts(textSpans, graphics, fonts);
        textSpans = SplitLeadingSpacesAtHighlightBoundaries(textSpans);
        textSpans = CoalesceAdjacentTextSpans(textSpans, compareHighlight: true);
        textSpans = CoalesceUnderlineSpans(textSpans);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            foreach (PptxPositionedTextSpan emissionSpan in SplitSpanByGlyphTypeface(span))
            {
                TextRun run = emissionSpan.Run;
                if (fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
                {
                    DrawWrappedSpan(graphics, rendered.ResourceName, rendered.Font, emissionSpan, rendered.SyntheticBold, rendered.SyntheticItalic);
                }
            }
        }
    }

    private static void DrawHighlightSpansWithFonts(IReadOnlyList<PptxPositionedTextSpan> textSpans, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        foreach (PptxPositionedTextSpan span in CoalesceHighlightSpans(textSpans))
        {
            TextRun run = span.Run;
            if (run.HighlightColor is null || !fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                continue;
            }

            DrawHighlightSpan(graphics, rendered.Font, span);
        }
    }

    private static IReadOnlyList<PptxPositionedTextSpan> SplitLeadingSpacesAtHighlightBoundaries(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        if (textSpans.Count < 2)
        {
            return textSpans;
        }

        var split = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            int leadingSpaceCount = CountLeadingSpaces(span.Run.Text);
            if (split.Count == 0 ||
                leadingSpaceCount == 0 ||
                leadingSpaceCount >= span.Run.Text.Length ||
                !PreservesHighlightTextOperationBoundaries(split[^1], span) ||
                split[^1].Run.HighlightColor.Equals(span.Run.HighlightColor))
            {
                split.Add(span);
                continue;
            }

            split.AddRange(SplitLeadingSpaceSpan(span, leadingSpaceCount));
        }

        return split;
    }

    private static IEnumerable<PptxPositionedTextSpan> SplitLeadingSpaceSpan(PptxPositionedTextSpan span, int leadingSpaceCount)
    {
        string leadingText = span.Run.Text[..leadingSpaceCount];
        string remainingText = span.Run.Text[leadingSpaceCount..];
        PptxTextGlyphLayout[] leadingGlyphs = span.GlyphSpan.Glyphs.Take(leadingSpaceCount).ToArray();
        PptxTextGlyphLayout[] remainingGlyphs = span.GlyphSpan.Glyphs.Skip(leadingSpaceCount).ToArray();
        double leadingWidth = leadingGlyphs.Length == 0
            ? Math.Min(span.Run.Width, span.GlyphSpan.LayoutWidth * leadingSpaceCount / Math.Max(1, span.Run.Text.Length))
            : GlyphNaturalWidth(leadingGlyphs);
        double remainingLeadingAdjustment = remainingGlyphs.Length == 0 ? 0d : remainingGlyphs[0].AdjustmentBefore;
        if (remainingGlyphs.Length > 0 && Math.Abs(remainingLeadingAdjustment) > PptxTextMetricRules.TextStateTolerance)
        {
            remainingGlyphs[0] = remainingGlyphs[0] with { AdjustmentBefore = 0d };
        }

        TextRun leadingRun = span.Run with
        {
            Text = leadingText,
            Width = Math.Max(0d, leadingWidth),
            PreventCoalesce = true
        };
        yield return span with
        {
            Run = leadingRun,
            EndX = leadingRun.X + leadingWidth,
            Atoms = [new PptxTextAtomLayout(PptxTextAtomKind.Space, leadingText, leadingRun.X, leadingWidth, Draw: true)],
            GlyphSpan = SliceGlyphSpan(span.GlyphSpan, leadingText, leadingWidth, leadingGlyphs)
        };

        double remainingX = span.Run.X + leadingWidth + remainingLeadingAdjustment;
        double remainingWidth = Math.Max(0d, span.EndX - remainingX);
        TextRun remainingRun = span.Run with
        {
            Text = remainingText,
            X = remainingX,
            Width = remainingWidth
        };
        yield return span with
        {
            Run = remainingRun,
            EndX = span.EndX,
            Atoms = [new PptxTextAtomLayout(PptxTextAtomKind.Word, remainingText, remainingRun.X, remainingWidth, Draw: true)],
            GlyphSpan = SliceGlyphSpan(span.GlyphSpan, remainingText, remainingWidth, remainingGlyphs)
        };
    }

    private static PptxTextGlyphSpanLayout SliceGlyphSpan(PptxTextGlyphSpanLayout source, string text, double width, IReadOnlyList<PptxTextGlyphLayout> glyphs)
    {
        double naturalWidth = GlyphNaturalWidth(glyphs);
        return source with
        {
            Text = text,
            NaturalWidth = naturalWidth,
            LayoutWidth = width,
            Glyphs = glyphs
        };
    }

    private static double GlyphNaturalWidth(IReadOnlyList<PptxTextGlyphLayout> glyphs)
    {
        return Math.Max(0d, glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore));
    }

    private static IReadOnlyList<TextRun> CoalesceAdjacentTextRuns(IReadOnlyList<TextRun> textRuns, bool compareHighlight = true)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (run.Text.Length == 0)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run, compareHighlight))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = run.X + run.Width - previous.X
                };
            }
            else
            {
                coalesced.Add(run);
            }
        }

        return coalesced;
    }

    private static IReadOnlyList<PptxPositionedTextSpan> CoalesceAdjacentTextSpans(IReadOnlyList<PptxPositionedTextSpan> textSpans, bool compareHighlight = true)
    {
        var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            TextRun run = span.Run;
            if (run.Text.Length == 0)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextSpan(coalesced[^1], span, compareHighlight))
            {
                PptxPositionedTextSpan previous = coalesced[^1];
                TextRun mergedRun = previous.Run with
                {
                    Text = previous.Run.Text + run.Text,
                    Width = run.X + run.Width - previous.Run.X
                };
                coalesced[^1] = previous with
                {
                    Run = mergedRun,
                    EndX = span.EndX,
                    Atoms = previous.Atoms.Concat(span.Atoms).ToArray(),
                    GlyphSpan = MergeGlyphSpans(mergedRun, previous.Run, previous.GlyphSpan, run, span.GlyphSpan)
                };
            }
            else
            {
                coalesced.Add(span);
            }
        }

        return coalesced;
    }

    private static bool CanCoalesceTextSpan(PptxPositionedTextSpan left, PptxPositionedTextSpan right, bool compareHighlight = true)
    {
        return CanCoalesceTextRun(left.Run, right.Run, compareHighlight, PreservesHighlightTextOperationBoundaries(left, right));
    }

    private static bool CanCoalesceTextRun(TextRun left, TextRun right, bool compareHighlight = true, bool preserveHighlightBoundary = true)
    {
        return Math.Abs(left.Y - right.Y) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.FontSize - right.FontSize) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.CharacterSpacing - right.CharacterSpacing) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.BaselineOffset - right.BaselineOffset) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.RotationDegrees - right.RotationDegrees) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.RotationCenterX - right.RotationCenterX) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.RotationCenterY - right.RotationCenterY) < PptxTextMetricRules.CoordinateTolerance &&
            left.FlipHorizontal == right.FlipHorizontal &&
            left.FlipVertical == right.FlipVertical &&
            Math.Abs(left.ClipX - right.ClipX) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.ClipY - right.ClipY) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.ClipWidth - right.ClipWidth) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.ClipHeight - right.ClipHeight) < PptxTextMetricRules.CoordinateTolerance &&
            left.Color.Equals(right.Color) &&
            Math.Abs(left.Alpha - right.Alpha) < PptxTextMetricRules.TextStateTolerance &&
            TextOutlinesEqual(left.Outline, right.Outline) &&
            (!compareHighlight || !preserveHighlightBoundary || left.HighlightColor.Equals(right.HighlightColor)) &&
            !left.PreventCoalesce &&
            !right.PreventCoalesce &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.Underline == right.Underline &&
            left.Strike == right.Strike &&
            left.KerningEnabled == right.KerningEnabled &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            right.X >= left.X &&
            Math.Abs(right.X - (left.X + left.Width)) < PptxTextMetricRules.TextCoalesceGap(left.FontSize);
    }

    private static bool PreservesHighlightTextOperationBoundaries(PptxPositionedTextSpan left, PptxPositionedTextSpan right)
    {
        return left.SourceAlignment != TextAlignment.Left || right.SourceAlignment != TextAlignment.Left;
    }

    private static void DrawHighlightRunsWithFonts(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        foreach (TextRun run in CoalesceHighlightRuns(textRuns))
        {
            if (run.HighlightColor is null || !fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                continue;
            }

            DrawHighlightRun(graphics, rendered.Font, run);
        }
    }

    private static IReadOnlyList<TextRun> CoalesceHighlightRuns(IReadOnlyList<TextRun> textRuns)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (run.Text.Length == 0 || run.HighlightColor is null)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = run.X + run.Width - previous.X
                };
            }
            else
            {
                coalesced.Add(run);
            }
        }

        return coalesced;
    }

    private static IReadOnlyList<PptxPositionedTextSpan> CoalesceHighlightSpans(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            TextRun run = span.Run;
            if (run.Text.Length == 0 || run.HighlightColor is null)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1].Run, run))
            {
                PptxPositionedTextSpan previous = coalesced[^1];
                TextRun mergedRun = previous.Run with
                {
                    Text = previous.Run.Text + run.Text,
                    Width = run.X + run.Width - previous.Run.X
                };
                coalesced[^1] = previous with
                {
                    Run = mergedRun,
                    EndX = span.EndX,
                    Atoms = previous.Atoms.Concat(span.Atoms).ToArray(),
                    GlyphSpan = MergeGlyphSpans(mergedRun, previous.Run, previous.GlyphSpan, run, span.GlyphSpan)
                };
            }
            else
            {
                coalesced.Add(span);
            }
        }

        return coalesced;
    }

    private static IReadOnlyList<TextRun> CoalesceUnderlineRuns(IReadOnlyList<TextRun> textRuns)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (coalesced.Count > 0 && CanCoalesceUnderlineRun(coalesced[^1], run))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = Math.Max(previous.Width, run.X + run.Width - previous.X)
                };
                continue;
            }

            coalesced.Add(run);
        }

        return coalesced;
    }

    private static IReadOnlyList<PptxPositionedTextSpan> CoalesceUnderlineSpans(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            if (coalesced.Count > 0 && CanCoalesceUnderlineRun(coalesced[^1].Run, span.Run))
            {
                PptxPositionedTextSpan previous = coalesced[^1];
                TextRun mergedRun = previous.Run with
                {
                    Text = previous.Run.Text + span.Run.Text,
                    Width = Math.Max(previous.Run.Width, span.Run.X + span.Run.Width - previous.Run.X)
                };
                coalesced[^1] = previous with
                {
                    Run = mergedRun,
                    EndX = span.EndX,
                    Atoms = previous.Atoms.Concat(span.Atoms).ToArray(),
                    GlyphSpan = MergeGlyphSpans(mergedRun, previous.Run, previous.GlyphSpan, span.Run, span.GlyphSpan)
                };
                continue;
            }

            coalesced.Add(span);
        }

        return coalesced;
    }

    private static PptxTextGlyphSpanLayout MergeGlyphSpans(TextRun mergedRun, TextRun leftRun, PptxTextGlyphSpanLayout left, TextRun rightRun, PptxTextGlyphSpanLayout right)
    {
        if (left.Glyphs.Count == 0 && right.Glyphs.Count == 0)
        {
            return PptxTextGlyphSpanLayout.Empty(mergedRun);
        }

        if (left.Glyphs.Count == 0)
        {
            return right with
            {
                Text = mergedRun.Text,
                LayoutWidth = mergedRun.Width
            };
        }

        if (right.Glyphs.Count == 0)
        {
            return left with
            {
                Text = mergedRun.Text,
                LayoutWidth = mergedRun.Width
            };
        }

        double leftVisualEnd = leftRun.X + left.NaturalWidth;
        double interSpanGap = rightRun.X - leftVisualEnd;
        PptxTextGlyphLayout[] glyphs = left.Glyphs
            .Concat(right.Glyphs.Select((glyph, index) => index == 0
                ? glyph with { AdjustmentBefore = glyph.AdjustmentBefore + interSpanGap }
                : glyph))
            .ToArray();
        double naturalWidth = glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore);
        return left with
        {
            Text = mergedRun.Text,
            NaturalWidth = Math.Max(0d, naturalWidth),
            LayoutWidth = mergedRun.Width,
            Glyphs = glyphs
        };
    }

    private static bool CanCoalesceUnderlineRun(TextRun left, TextRun right)
    {
        return left.Underline &&
            right.Underline &&
            !left.Strike &&
            !right.Strike &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.KerningEnabled == right.KerningEnabled &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            left.Color.Equals(right.Color) &&
            NearlyEqual(left.Alpha, right.Alpha) &&
            TextOutlinesEqual(left.Outline, right.Outline) &&
            left.HighlightColor.Equals(right.HighlightColor) &&
            !left.PreventCoalesce &&
            !right.PreventCoalesce &&
            NearlyEqual(left.Y, right.Y) &&
            NearlyEqual(left.Height, right.Height) &&
            NearlyEqual(left.ClipX, right.ClipX) &&
            NearlyEqual(left.ClipY, right.ClipY) &&
            NearlyEqual(left.ClipWidth, right.ClipWidth) &&
            NearlyEqual(left.ClipHeight, right.ClipHeight) &&
            NearlyEqual(left.FontSize, right.FontSize) &&
            NearlyEqual(left.CharacterSpacing, right.CharacterSpacing) &&
            NearlyEqual(left.BaselineOffset, right.BaselineOffset) &&
            NearlyEqual(left.RotationDegrees, right.RotationDegrees) &&
            NearlyEqual(left.RotationCenterX, right.RotationCenterX) &&
            NearlyEqual(left.RotationCenterY, right.RotationCenterY) &&
            left.FlipHorizontal == right.FlipHorizontal &&
            left.FlipVertical == right.FlipVertical &&
            Math.Abs((left.X + left.Width) - right.X) <= PptxTextMetricRules.UnderlineCoalesceGap(left.FontSize);
    }

    private static bool NearlyEqual(double left, double right)
    {
        return Math.Abs(left - right) <= PptxTextMetricRules.TextStateTolerance;
    }

    private static bool TextOutlinesEqual(TextOutline? left, TextOutline? right)
    {
        if (left is null || right is null)
        {
            return left is null && right is null;
        }

        TextOutline leftOutline = left.Value;
        TextOutline rightOutline = right.Value;
        return leftOutline.Color.Equals(rightOutline.Color) &&
            NearlyEqual(leftOutline.Alpha, rightOutline.Alpha) &&
            NearlyEqual(leftOutline.Width, rightOutline.Width);
    }

    private static IEnumerable<PptxPositionedTextSpan> SplitSpanByGlyphTypeface(PptxPositionedTextSpan span)
    {
        if (span.GlyphSpan.Glyphs.Count == 0)
        {
            yield return span;
            yield break;
        }

        double cursor = span.Run.X;
        int index = 0;
        while (index < span.GlyphSpan.Glyphs.Count)
        {
            PptxTextGlyphLayout first = span.GlyphSpan.Glyphs[index];
            string typeface = string.IsNullOrWhiteSpace(first.Typeface) ? (span.Run.FontFamily ?? "Arial") : first.Typeface!;
            int start = index;
            index++;
            while (index < span.GlyphSpan.Glyphs.Count &&
                string.Equals(span.GlyphSpan.Glyphs[index].Typeface ?? span.Run.FontFamily ?? "Arial", typeface, StringComparison.OrdinalIgnoreCase))
            {
                index++;
            }

            PptxTextGlyphLayout[] glyphs = span.GlyphSpan.Glyphs.Skip(start).Take(index - start).ToArray();
            double leadingOffset = start == 0 ? 0d : glyphs[0].AdjustmentBefore;
            if (start != 0 && Math.Abs(leadingOffset) > PptxTextMetricRules.TextStateTolerance)
            {
                glyphs[0] = glyphs[0] with { AdjustmentBefore = 0d };
            }

            double naturalWidth = glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore);
            double x = cursor + leadingOffset;
            string text = string.Concat(glyphs.Select(glyph => char.ConvertFromUtf32(glyph.CodePoint)));
            TextRun run = span.Run with
            {
                Text = text,
                X = x,
                Width = naturalWidth,
                FontFamily = typeface,
                Alignment = TextAlignment.Left,
                PreventCoalesce = true
            };
            yield return span with
            {
                Run = run,
                EndX = x + naturalWidth,
                GlyphSpan = new PptxTextGlyphSpanLayout(
                    text,
                    typeface,
                    span.GlyphSpan.Bold,
                    span.GlyphSpan.Italic,
                    span.GlyphSpan.FontSize,
                    span.GlyphSpan.CharacterSpacing,
                    span.GlyphSpan.KerningEnabled,
                    start == 0 ? span.GlyphSpan.LeadingAdjustment : 0d,
                    naturalWidth,
                    naturalWidth,
                    glyphs)
            };

            cursor = x + naturalWidth;
        }
    }

    private static string FontKey(TextRun run)
    {
        string familyName = string.IsNullOrWhiteSpace(run.FontFamily) ? "Arial" : run.FontFamily!;
        return FontKey(familyName, run.Bold, run.Italic);
    }

    private static string FontKey(string familyName, bool bold, bool italic)
    {
        return familyName + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture);
    }

    private static void DrawWrappedRun(PdfGraphicsBuilder graphics, string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
    {
        graphics.SaveState();
        if (HasTextTransform(run))
        {
            ApplyTextTransform(graphics, run);
        }

        graphics.ClipRectangleEvenOdd(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        TextGlyphRun? glyphRun = BuildTextGlyphRun(resourceName, embedded, run, syntheticBold, syntheticItalic);
        if (glyphRun is not null)
        {
            bool needsTextAlpha = run.Alpha < 1d - PptxTextMetricRules.TextStateTolerance ||
                (run.Outline is { } runOutline && runOutline.Alpha < 1d - PptxTextMetricRules.TextStateTolerance);
            if (needsTextAlpha)
            {
                graphics.SaveState();
                graphics.SetAlpha(run.Alpha, run.Outline?.Alpha ?? 1d);
            }

            DrawGlyphText(graphics, glyphRun);

            if (run.Underline)
            {
                graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                double underlineScale = run.FontSize / embedded.Font.UnitsPerEm;
                double underlineThickness = Math.Max(PptxTextMetricRules.MinimumStrokeWidth, Math.Abs(embedded.Font.Post.UnderlineThickness) * underlineScale);
                double underlineY = glyphRun.BaselineY + (embedded.Font.Post.UnderlinePosition - Math.Abs(embedded.Font.Post.UnderlineThickness)) * underlineScale;
                graphics.FillRectangle(glyphRun.X, underlineY, glyphRun.Width, underlineThickness);
            }

            if (run.Strike)
            {
                graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                graphics.FillRectangle(glyphRun.X, PptxTextMetricRules.StrikeY(glyphRun.BaselineY, run.FontSize), glyphRun.Width, PptxTextMetricRules.StrikeThickness(run.FontSize));
            }

            if (needsTextAlpha)
            {
                graphics.RestoreState();
            }
        }

        graphics.RestoreState();
    }

    private static void DrawWrappedSpan(PdfGraphicsBuilder graphics, string resourceName, PdfEmbeddedFont embedded, PptxPositionedTextSpan span, bool syntheticBold, bool syntheticItalic)
    {
        TextRun run = span.Run;
        graphics.SaveState();
        if (HasTextTransform(run))
        {
            ApplyTextTransform(graphics, run);
        }

        graphics.ClipRectangleEvenOdd(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        TextGlyphRun? glyphRun = BuildTextGlyphRun(resourceName, embedded, span, syntheticBold, syntheticItalic);
        if (glyphRun is not null)
        {
            bool needsTextAlpha = run.Alpha < 1d - PptxTextMetricRules.TextStateTolerance ||
                (run.Outline is { } runOutline && runOutline.Alpha < 1d - PptxTextMetricRules.TextStateTolerance);
            if (needsTextAlpha)
            {
                graphics.SaveState();
                graphics.SetAlpha(run.Alpha, run.Outline?.Alpha ?? 1d);
            }

            DrawGlyphText(graphics, glyphRun);

            if (run.Underline)
            {
                graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                double underlineScale = run.FontSize / embedded.Font.UnitsPerEm;
                double underlineThickness = Math.Max(PptxTextMetricRules.MinimumStrokeWidth, Math.Abs(embedded.Font.Post.UnderlineThickness) * underlineScale);
                double underlineY = glyphRun.BaselineY + (embedded.Font.Post.UnderlinePosition - Math.Abs(embedded.Font.Post.UnderlineThickness)) * underlineScale;
                graphics.FillRectangle(glyphRun.X, underlineY, glyphRun.Width, underlineThickness);
            }

            if (run.Strike)
            {
                graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                graphics.FillRectangle(glyphRun.X, PptxTextMetricRules.StrikeY(glyphRun.BaselineY, run.FontSize), glyphRun.Width, PptxTextMetricRules.StrikeThickness(run.FontSize));
            }

            if (needsTextAlpha)
            {
                graphics.RestoreState();
            }
        }

        graphics.RestoreState();
    }

    private static TextGlyphRun? BuildTextGlyphRun(string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
    {
        string glyphHex = embedded.EncodeGlyphHex(run.Text);
        double baselineY = run.Y + run.BaselineOffset;
        if (glyphHex.Length == 0 || !BaselineIntersectsClip(run, baselineY))
        {
            return null;
        }

        double lineWidth = MeasureRenderedText(embedded, run.Text, run.FontSize, run.CharacterSpacing, run.KerningEnabled);
        double x = run.Alignment switch
        {
            TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
            TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
            _ => run.X
        };
        IReadOnlyList<TextGlyphAtom> glyphs = BuildTextGlyphAtoms(embedded, run);
        double pdfFontSize = PptxPdfTextEmissionProfile.FontSize(run.FontSize);
        string? positioningArray = EncodeGlyphPositioningArray(glyphs, run.FontSize, pdfFontSize, forcePositioningArray: true);
        return new TextGlyphRun(run, resourceName, embedded, glyphHex, positioningArray, glyphs, x, baselineY, lineWidth, pdfFontSize, syntheticBold, syntheticItalic);
    }

    private static TextGlyphRun? BuildTextGlyphRun(string resourceName, PdfEmbeddedFont embedded, PptxPositionedTextSpan span, bool syntheticBold, bool syntheticItalic)
    {
        TextRun run = span.Run;
        string glyphHex = EncodeGlyphHex(span.GlyphSpan);
        double baselineY = run.Y + run.BaselineOffset;
        if (glyphHex.Length == 0 || !BaselineIntersectsClip(run, baselineY))
        {
            return null;
        }

        double lineWidth = span.GlyphSpan.NaturalWidth;
        double x = run.Alignment switch
        {
            TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
            TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
            _ => run.X
        };
        double pdfFontSize = PptxPdfTextEmissionProfile.FontSize(run.FontSize);
        string? positioningArray = EncodeGlyphPositioningArray(span.GlyphSpan, pdfFontSize, forcePositioningArray: true);
        IReadOnlyList<TextGlyphAtom> glyphs = span.GlyphSpan.Glyphs
            .Select(glyph => new TextGlyphAtom(glyph.CodePoint, glyph.Typeface, glyph.GlyphId, glyph.Advance, glyph.AdjustmentBefore))
            .ToArray();
        return new TextGlyphRun(run, resourceName, embedded, glyphHex, positioningArray, glyphs, x, baselineY, lineWidth, pdfFontSize, syntheticBold, syntheticItalic);
    }

    private static string EncodeGlyphHex(PptxTextGlyphSpanLayout span)
    {
        var builder = new StringBuilder(span.Glyphs.Count * 4);
        foreach (PptxTextGlyphLayout glyph in span.Glyphs)
        {
            builder.Append(glyph.GlyphId.ToString("X4", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static string? EncodeGlyphPositioningArray(PptxTextGlyphSpanLayout span, double pdfFontSize, bool forcePositioningArray)
    {
        if (span.Glyphs.Count == 0)
        {
            return null;
        }

        bool hasPositioning = false;
        var builder = new StringBuilder("[");
        for (int i = 0; i < span.Glyphs.Count; i++)
        {
            PptxTextGlyphLayout glyph = span.Glyphs[i];
            if (i > 0)
            {
                PptxTextGlyphLayout previousGlyph = span.Glyphs[i - 1];
                double adjustmentBefore = PdfTextAdjustmentBefore(
                    glyph.AdjustmentBefore,
                    previousGlyph.Advance,
                    span.FontSize,
                    pdfFontSize);
                double adjustment = pdfFontSize <= 0d ? 0d : -adjustmentBefore * 1000d / pdfFontSize;
                if (Math.Abs(adjustment) > PptxTextMetricRules.TextStateTolerance)
                {
                    builder.Append(' ').Append(adjustment.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ');
                    hasPositioning = true;
                }
            }

            builder.Append('<').Append(glyph.GlyphId.ToString("X4", CultureInfo.InvariantCulture)).Append('>');
        }

        builder.Append(']');
        return hasPositioning || forcePositioningArray ? builder.ToString() : null;
    }

    private static string? EncodeGlyphPositioningArray(IReadOnlyList<TextGlyphAtom> glyphs, double layoutFontSize, double pdfFontSize, bool forcePositioningArray)
    {
        if (glyphs.Count == 0)
        {
            return null;
        }

        bool hasPositioning = false;
        var builder = new StringBuilder("[");
        for (int i = 0; i < glyphs.Count; i++)
        {
            TextGlyphAtom glyph = glyphs[i];
            if (i > 0)
            {
                TextGlyphAtom previousGlyph = glyphs[i - 1];
                double adjustmentBefore = PdfTextAdjustmentBefore(
                    glyph.AdjustmentBefore,
                    previousGlyph.Advance,
                    layoutFontSize,
                    pdfFontSize);
                double adjustment = pdfFontSize <= 0d ? 0d : -adjustmentBefore * 1000d / pdfFontSize;
                if (Math.Abs(adjustment) > PptxTextMetricRules.TextStateTolerance)
                {
                    builder.Append(' ').Append(adjustment.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ');
                    hasPositioning = true;
                }
            }

            builder.Append('<').Append(glyph.GlyphId.ToString("X4", CultureInfo.InvariantCulture)).Append('>');
        }

        builder.Append(']');
        return hasPositioning || forcePositioningArray ? builder.ToString() : null;
    }

    private static double PdfTextAdjustmentBefore(double layoutAdjustmentBefore, double previousLayoutAdvance, double layoutFontSize, double pdfFontSize)
    {
        if (layoutFontSize <= 0d || pdfFontSize <= 0d)
        {
            return layoutAdjustmentBefore;
        }

        double previousPdfAdvance = previousLayoutAdvance * pdfFontSize / layoutFontSize;
        return layoutAdjustmentBefore + previousLayoutAdvance - previousPdfAdvance;
    }

    private static IReadOnlyList<TextGlyphAtom> BuildTextGlyphAtoms(PdfEmbeddedFont embedded, TextRun run)
    {
        var atoms = new List<TextGlyphAtom>();
        ushort previousGlyph = 0;
        foreach (Rune rune in run.Text.EnumerateRunes())
        {
            ushort glyph = embedded.Font.MapCodePoint(rune.Value);
            if (glyph == 0)
            {
                continue;
            }

            double adjustmentBefore = 0d;
            if (atoms.Count > 0)
            {
                adjustmentBefore += run.CharacterSpacing;
                if (run.KerningEnabled && previousGlyph != 0)
                {
                    adjustmentBefore += embedded.Font.GetKerning(previousGlyph, glyph) * run.FontSize / embedded.Font.UnitsPerEm;
                }
            }

            double advance = embedded.Font.GetAdvanceWidth(glyph) * run.FontSize / embedded.Font.UnitsPerEm;
            atoms.Add(new TextGlyphAtom(rune.Value, run.FontFamily, glyph, advance, adjustmentBefore));
            previousGlyph = glyph;
        }

        return atoms;
    }

    private static void DrawHighlightRun(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run)
    {
        if (run.HighlightColor is not { } highlight)
        {
            return;
        }

        double baselineY = run.Y + run.BaselineOffset;
        double lineWidth = MeasureRenderedText(embedded, run.Text, run.FontSize, run.CharacterSpacing, run.KerningEnabled);
        DrawHighlightRectangle(graphics, embedded, run, highlight, baselineY, lineWidth);
    }

    private static void DrawHighlightSpan(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, PptxPositionedTextSpan span)
    {
        TextRun run = span.Run;
        if (run.HighlightColor is not { } highlight)
        {
            return;
        }

        double baselineY = span.LineBox?.BaselineY ?? run.Y + run.BaselineOffset;
        DrawHighlightRectangle(graphics, embedded, run, highlight, baselineY, span.GlyphSpan.NaturalWidth);
    }

    private static void DrawHighlightRectangle(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run, RgbColor highlight, double baselineY, double lineWidth)
    {
        if (!BaselineIntersectsClip(run, baselineY))
        {
            return;
        }

        graphics.SaveState();
        if (HasTextTransform(run))
        {
            ApplyTextTransform(graphics, run);
        }

        graphics.ClipRectangleEvenOdd(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        graphics.SetFillRgb(highlight.Red, highlight.Green, highlight.Blue);
        double fontScale = run.FontSize / embedded.Font.UnitsPerEm;
        double highlightDescent = PptxTextMetricRules.HighlightDescent(embedded, run.FontSize, fontScale);
        double highlightHeight = PptxTextMetricRules.HighlightHeight(embedded, run.FontSize, fontScale);
        double highlightY = baselineY - highlightDescent;
        graphics.FillRectangleEvenOdd(run.X, highlightY, lineWidth, highlightHeight);
        graphics.RestoreState();
    }

    private static bool BaselineIntersectsClip(TextRun run, double baselineY)
    {
        if (HasTextTransform(run))
        {
            return true;
        }

        if (run.StrictClip)
        {
            return baselineY >= run.ClipY - PptxTextMetricRules.TextStateTolerance &&
                baselineY <= run.ClipY + run.ClipHeight + PptxTextMetricRules.TextStateTolerance;
        }

        return baselineY + run.FontSize >= run.ClipY &&
            baselineY - run.FontSize <= run.ClipY + run.ClipHeight;
    }

    private static bool HasTextTransform(TextRun run)
    {
        return Math.Abs(run.RotationDegrees) > PptxTextMetricRules.TextStateTolerance ||
            run.FlipHorizontal ||
            run.FlipVertical;
    }

    private static void ApplyTextTransform(PdfGraphicsBuilder graphics, TextRun run)
    {
        double radians = -run.RotationDegrees * Math.PI / 180d;
        double sx = run.FlipHorizontal ? -1d : 1d;
        double sy = run.FlipVertical ? -1d : 1d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double a = cos * sx;
        double b = sin * sx;
        double c = -sin * sy;
        double d = cos * sy;
        double e = run.RotationCenterX - a * run.RotationCenterX - c * run.RotationCenterY;
        double f = run.RotationCenterY - b * run.RotationCenterX - d * run.RotationCenterY;
        graphics.Transform(a, b, c, d, e, f);
    }

    private static void DrawGlyphText(PdfGraphicsBuilder graphics, TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        double pdfFontSize = glyphRun.PdfFontSize;
        if (glyphRun.PositioningArray is null)
        {
            graphics.DrawGlyphText(
                glyphRun.ResourceName,
                pdfFontSize,
                glyphRun.X,
                glyphRun.BaselineY,
                run.Color.Red,
                run.Color.Green,
                run.Color.Blue,
                glyphRun.GlyphHex,
                glyphRun.SyntheticItalic,
                textRenderingMode: TextRenderingMode(glyphRun),
                strokeRed: TextStrokeColor(glyphRun).Red,
                strokeGreen: TextStrokeColor(glyphRun).Green,
                strokeBlue: TextStrokeColor(glyphRun).Blue,
                strokeWidth: TextStrokeWidth(glyphRun));
        }
        else
        {
            graphics.DrawGlyphPositionedText(
                glyphRun.ResourceName,
                pdfFontSize,
                glyphRun.X,
                glyphRun.BaselineY,
                run.Color.Red,
                run.Color.Green,
                run.Color.Blue,
                glyphRun.PositioningArray,
                glyphRun.SyntheticItalic,
                textRenderingMode: TextRenderingMode(glyphRun),
                strokeRed: TextStrokeColor(glyphRun).Red,
                strokeGreen: TextStrokeColor(glyphRun).Green,
                strokeBlue: TextStrokeColor(glyphRun).Blue,
                strokeWidth: TextStrokeWidth(glyphRun));
        }
    }

    private static int TextRenderingMode(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        if (run.Outline is null)
        {
            return glyphRun.SyntheticBold ? 2 : 0;
        }

        return run.Alpha <= PptxTextMetricRules.TextStateTolerance ? 1 : 2;
    }

    private static RgbColor TextStrokeColor(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        return run.Outline?.Color ?? run.Color;
    }

    private static double TextStrokeWidth(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        return run.Outline?.Width ?? PptxTextMetricRules.SyntheticBoldStrokeWidth(run.FontSize);
    }

    private static IEnumerable<string> WrapWords(string text, double maxWidth, double fontSize, double characterSpacing, PdfEmbeddedFont embedded)
    {
        if (MeasureRenderedText(embedded, text, fontSize, characterSpacing, kerningEnabled: true) <= maxWidth)
        {
            yield return text;
            yield break;
        }

        string[] words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
        {
            yield break;
        }

        var line = new StringBuilder();
        foreach (string word in words)
        {
            string candidate = line.Length == 0 ? word : line + " " + word;
            if (line.Length > 0 && MeasureRenderedText(embedded, candidate, fontSize, characterSpacing, kerningEnabled: true) > maxWidth)
            {
                yield return line.ToString();
                line.Clear();
                line.Append(word);
            }
            else
            {
                line.Clear();
                line.Append(candidate);
            }
        }

        if (line.Length > 0)
        {
            yield return line.ToString();
        }
    }

    private static double MeasureRenderedText(PdfEmbeddedFont embedded, string text, double fontSize, double characterSpacing, bool kerningEnabled)
    {
        double width = embedded.MeasureTextPoints(text, fontSize, kerningEnabled);
        int runeCount = text.EnumerateRunes().Count();
        return Math.Max(0d, width + Math.Max(0, runeCount - 1) * characterSpacing);
    }
}
