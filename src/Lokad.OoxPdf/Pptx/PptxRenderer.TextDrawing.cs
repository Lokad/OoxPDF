using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<PdfFontResource> RenderTextRuns(
        IReadOnlyList<TextRun> textRuns,
        PdfGraphicsBuilder graphics,
        string resourcePrefix,
        PresentationFontResolver? fontResolver)
    {
        if (textRuns.Count == 0)
        {
            return [];
        }

        RenderedFonts renderedFonts = CreateRenderedFonts(textRuns, fontResolver ?? new PresentationFontResolver(null), resourcePrefix, CancellationToken.None);
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

        RenderedFonts renderedFonts = CreateRenderedFonts(textSpans, legacyTextRuns, new PresentationFontResolver(null), "F", CancellationToken.None);
        DrawTextSpansWithFonts(textSpans, graphics, renderedFonts.Fonts);
        DrawTextRunsWithFonts(legacyTextRuns, graphics, renderedFonts.Fonts);
        return renderedFonts.Resources;
    }

    private static RenderedFonts CreateRenderedFonts(
        IReadOnlyList<PptxPositionedTextSpan> textSpans,
        IReadOnlyList<TextRun> legacyTextRuns,
        PresentationFontResolver fontResolver,
        string resourcePrefix,
        CancellationToken cancellationToken)
    {
        var uses = new List<TextFontUse>();
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (IGrouping<string, PptxTextGlyphLayout> group in span.GlyphSpan.Glyphs.GroupBy(
                        glyph => string.IsNullOrWhiteSpace(glyph.Typeface) ? PptxFontFallbackRules.ResolveDefaultLatinTypeface(span.Run.FontFamily) : glyph.Typeface,
                         StringComparer.OrdinalIgnoreCase))
            {
                uses.Add(new TextFontUse(group.Key, span.GlyphSpan.Bold, span.GlyphSpan.Italic, group.Select(glyph => glyph.CodePoint).ToArray()));
            }
        }

        foreach (TextRun run in CoalesceUnderlineRuns(CoalesceAdjacentTextRuns(legacyTextRuns, compareHighlight: false)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string familyName = PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily);
            uses.Add(new TextFontUse(familyName, run.Bold, run.Italic, run.Text.EnumerateRunes().Select(rune => rune.Value).ToArray()));
        }

        return CreateRenderedFonts(uses, fontResolver, resourcePrefix, cancellationToken);
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextRun> textRuns, PresentationFontResolver fontResolver, string resourcePrefix, CancellationToken cancellationToken)
    {
        if (textRuns.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }

        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        return CreateRenderedFonts(textRuns
            .Select(run => new TextFontUse(
                PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily),
                run.Bold,
                run.Italic,
                run.Text.EnumerateRunes().Select(rune => rune.Value).ToArray()))
            .ToArray(), fontResolver, resourcePrefix, cancellationToken);
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextFontUse> uses, PresentationFontResolver fontResolver, string resourcePrefix, CancellationToken cancellationToken)
    {
        if (uses.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }

        var fonts = new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase);
        var resources = new List<PdfFontResource>();
        foreach (IGrouping<string, TextFontUse> group in uses.GroupBy(use => FontKey(use.FamilyName, use.Bold, use.Italic), StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            TextFontUse first = group.First();
            (FontFaceResolution Resolution, OpenTypeFont Font)? resolved = fontResolver.ResolvePresentationOpenTypeFont(new FontRequest(first.FamilyName, first.Bold, first.Italic), cancellationToken);
            if (resolved is null)
            {
                continue;
            }

            FontFaceResolution resolution = resolved.Value.Resolution;
            OpenTypeFont font = resolved.Value.Font;
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, group.SelectMany(use => use.CodePoints), cancellationToken);
            string resourceName = resourcePrefix + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            fonts[group.Key] = new RenderedFont(resourceName, embedded, resolution, first.Bold && !resolution.Bold, first.Italic && !resolution.Italic);
            resources.Add(new PdfFontResource(resourceName, embedded));
        }

        return new RenderedFonts(fonts, resources);
    }

    private static void DrawTextRunsWithFonts(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        DrawHighlightRunsWithFonts();
        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        foreach (TextRun run in textRuns)
        {
            if (fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                DrawWrappedRun(rendered.ResourceName, rendered.Font, run, rendered.SyntheticBold, rendered.SyntheticItalic);
            }
        }

        void DrawWrappedRun(string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
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
                    if (TryGetUnderlineRectangle(embedded, glyphRun, out TextDecorationRectangle underline))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, underline);
                    }
                }

                if (run.Strike)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    if (TryGetStrikeRectangle(embedded, glyphRun, out TextDecorationRectangle strike))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, strike);
                    }
                }

                if (needsTextAlpha)
                {
                    graphics.RestoreState();
                }
            }

            graphics.RestoreState();
        }

        void DrawHighlightRunsWithFonts()
        {
            foreach (TextRun run in CoalesceHighlightRuns())
            {
                if (run.HighlightColor is null || !fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
                {
                    continue;
                }

                DrawHighlightRun(rendered.Font, run);
            }

            IReadOnlyList<TextRun> CoalesceHighlightRuns()
            {
                var coalesced = new List<TextRun>(textRuns.Count);
                foreach (TextRun run in textRuns)
                {
                    if (run.Text.Length == 0 || run.HighlightColor is null)
                    {
                        continue;
                    }

                    if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run, true, true))
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

            void DrawHighlightRun(PdfEmbeddedFont embedded, TextRun run)
            {
                if (run.HighlightColor is not { } highlight)
                {
                    return;
                }

                double baselineY = run.Y + run.BaselineOffset;
                double lineWidth = MeasureRenderedText(embedded, run.Text, run.FontSize, run.CharacterSpacing, run.KerningEnabled);
                DrawHighlightRectangle(graphics, embedded, run, highlight, baselineY, lineWidth);
            }
        }
    }

    private static void DrawTextSpansWithFonts(IReadOnlyList<PptxPositionedTextSpan> textSpans, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        DrawHighlightSpansWithFonts();
        textSpans = SplitLeadingSpacesAtHighlightBoundaries(textSpans);
        textSpans = CoalesceAdjacentTextSpans(textSpans, compareHighlight: true);
        textSpans = CoalesceUnderlineSpans(textSpans);
        textSpans = ApplyOfficePdfCharacterSpacing(textSpans);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            foreach (PptxPositionedTextSpan emissionSpan in SplitSpanByGlyphTypeface(span))
            {
                TextRun run = emissionSpan.Run;
                if (fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
                {
                    DrawWrappedSpan(rendered.ResourceName, rendered.Font, emissionSpan, rendered.SyntheticBold, rendered.SyntheticItalic);
                }
            }
        }

        void DrawWrappedSpan(string resourceName, PdfEmbeddedFont embedded, PptxPositionedTextSpan span, bool syntheticBold, bool syntheticItalic)
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
                    if (TryGetUnderlineRectangle(embedded, glyphRun, out TextDecorationRectangle underline))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, underline);
                    }
                }

                if (run.Strike)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    if (TryGetStrikeRectangle(embedded, glyphRun, out TextDecorationRectangle strike))
                    {
                        FillTextDecorationRectangleEvenOdd(graphics, strike);
                    }
                }

                if (needsTextAlpha)
                {
                    graphics.RestoreState();
                }
            }

            graphics.RestoreState();
        }

        void DrawHighlightSpansWithFonts()
        {
            foreach (PptxPositionedTextSpan span in CoalesceHighlightSpans())
            {
                TextRun run = span.Run;
                if (run.HighlightColor is null || !fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
                {
                    continue;
                }

                DrawHighlightSpan(rendered.Font, span);
            }

            IReadOnlyList<PptxPositionedTextSpan> CoalesceHighlightSpans()
            {
                var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
                foreach (PptxPositionedTextSpan span in textSpans)
                {
                    TextRun run = span.Run;
                    if (run.Text.Length == 0 || run.HighlightColor is null)
                    {
                        continue;
                    }

                    if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1].Run, run, true, true))
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

            void DrawHighlightSpan(PdfEmbeddedFont embedded, PptxPositionedTextSpan span)
            {
                TextRun run = span.Run;
                if (run.HighlightColor is not { } highlight)
                {
                    return;
                }

                double baselineY = span.LineBox?.BaselineY ?? run.Y + run.BaselineOffset;
                DrawHighlightRectangle(graphics, embedded, run, highlight, baselineY, span.GlyphSpan.NaturalWidth);
            }
        }
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ApplyOfficePdfCharacterSpacing(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        if (textSpans.Count == 0)
        {
            return textSpans;
        }

        var adjusted = new PptxPositionedTextSpan[textSpans.Count];
        Dictionary<int, (double SpacingEm, int FirstNumberedParagraphIndex)> numberedFrameSpacing = ReadOfficeNumberedFrameCharacterSpacing(textSpans);
        int currentFrame = -1;
        bool useHighlightContinuationSpacing = false;
        for (int i = 0; i < textSpans.Count; i++)
        {
            PptxPositionedTextSpan span = textSpans[i];
            if (span.FrameIndex != currentFrame)
            {
                currentFrame = span.FrameIndex;
                useHighlightContinuationSpacing = false;
            }

            bool hasZeroAuthoredCharacterSpacing =
                Math.Abs(span.GlyphSpan.CharacterSpacing) < PptxTextMetricRules.TextStateTolerance;
            if (hasZeroAuthoredCharacterSpacing && StartsOfficeHighlightContinuationTextState(span))
            {
                useHighlightContinuationSpacing = true;
            }

            bool eligibleNumberedProfileSpan =
                UsesNumberedTextStateProfile(span) &&
                hasZeroAuthoredCharacterSpacing;
            double? pdfCharacterSpacingOverride = null;
            if (hasZeroAuthoredCharacterSpacing && useHighlightContinuationSpacing)
            {
                pdfCharacterSpacingOverride = PptxTextMetricRules.OfficeHighlightContinuationCharacterSpacing(span.Run.FontSize);
            }
            else if (eligibleNumberedProfileSpan &&
                numberedFrameSpacing.TryGetValue(span.FrameIndex, out (double SpacingEm, int FirstNumberedParagraphIndex) frameSpacing) &&
                span.ParagraphIndex >= frameSpacing.FirstNumberedParagraphIndex)
            {
                pdfCharacterSpacingOverride = span.Run.FontSize * frameSpacing.SpacingEm;
            }

            adjusted[i] = pdfCharacterSpacingOverride is not null
                ? span with { PdfCharacterSpacingOverride = pdfCharacterSpacingOverride }
                : span;
        }

        return adjusted;
    }

    private static bool StartsOfficeHighlightContinuationTextState(PptxPositionedTextSpan span) =>
        span.Run.HighlightColor is not null &&
        span.FrameAutofitMode != "noAutofit" &&
        span.ParagraphIndex > 0;

    private static Dictionary<int, (double SpacingEm, int FirstNumberedParagraphIndex)> ReadOfficeNumberedFrameCharacterSpacing(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        var spacingByFrame = new Dictionary<int, (double SpacingEm, int FirstNumberedParagraphIndex)>();
        foreach (IGrouping<int, PptxPositionedTextSpan> frameGroup in textSpans.GroupBy(span => span.FrameIndex))
        {
            PptxPositionedTextSpan[] frame = frameGroup.ToArray();
            if (!frame.Any(UsesNumberedTextStateProfile) ||
                !frame.Any(span => string.Equals(span.ParagraphBulletKind, nameof(PptxParagraphBulletKind.AutoNumber), StringComparison.Ordinal)) ||
                frame.Any(span => Math.Abs(span.GlyphSpan.CharacterSpacing) >= PptxTextMetricRules.TextStateTolerance))
            {
                continue;
            }

            bool hasExplicitAutoNumberStart = frame.Any(span => span.ParagraphAutoNumberStartAt is not null);
            bool hasBodyContinuation = frame.Any(span =>
                !string.Equals(span.ParagraphBulletKind, nameof(PptxParagraphBulletKind.AutoNumber), StringComparison.Ordinal));
            double spacingEm = hasExplicitAutoNumberStart && hasBodyContinuation
                ? PptxTextMetricRules.OfficeAutofitNumberedDenseCharacterSpacingEm
                : PptxTextMetricRules.OfficeAutofitNumberedDefaultCharacterSpacingEm;
            int firstNumberedParagraphIndex = frame
                .Where(span => string.Equals(span.ParagraphBulletKind, nameof(PptxParagraphBulletKind.AutoNumber), StringComparison.Ordinal))
                .Min(span => span.ParagraphIndex);
            spacingByFrame[frameGroup.Key] = (spacingEm, firstNumberedParagraphIndex);
        }

        return spacingByFrame;
    }

    private static bool UsesNumberedTextStateProfile(PptxPositionedTextSpan span) =>
        string.Equals(span.FrameAutofitMode, "spAutoFit", StringComparison.Ordinal) ||
        string.Equals(span.FrameAutofitMode, "noAutofit", StringComparison.Ordinal);

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

    private static string FontKey(TextRun run)
    {
        string familyName = PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily);
        return FontKey(familyName, run.Bold, run.Italic);
    }

    private static string FontKey(string familyName, bool bold, bool italic)
    {
        return familyName + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture);
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
        IReadOnlyList<TextGlyphAtom> glyphs = BuildTextGlyphAtoms();
        double pdfFontSize = PptxPdfTextEmissionProfile.FontSize(run.FontSize);
        double pdfCharacterSpacing = run.CharacterSpacing;
        string? positioningArray = EncodeGlyphPositioningArray(embedded, glyphs, run.FontSize, pdfFontSize, pdfCharacterSpacing, forcePositioningArray: true);
        return new TextGlyphRun(run, resourceName, embedded, glyphHex, positioningArray, glyphs, x, baselineY, lineWidth, pdfFontSize, pdfCharacterSpacing, syntheticBold, syntheticItalic);

        IReadOnlyList<TextGlyphAtom> BuildTextGlyphAtoms()
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
                atoms.Add(new TextGlyphAtom(rune.Value, run.FontFamily, PptxGlyphTypefaceResolutionSource.Primary, glyph, advance, adjustmentBefore));
                previousGlyph = glyph;
            }

            return atoms;
        }
    }

    private static TextGlyphRun? BuildTextGlyphRun(string resourceName, PdfEmbeddedFont embedded, PptxPositionedTextSpan span, bool syntheticBold, bool syntheticItalic)
    {
        TextRun run = span.Run;
        string glyphHex = EncodeGlyphHex(embedded, span.GlyphSpan);
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
        PptxPdfTextEmissionContext emissionContext = CreatePdfTextEmissionContext();
        double pdfFontSize = PptxPdfTextEmissionProfile.FontSize(emissionContext);
        double pdfCharacterSpacing = span.PdfCharacterSpacingOverride
            ?? PptxPdfTextEmissionProfile.CharacterSpacing(emissionContext, span.GlyphSpan.CharacterSpacing);
        string? positioningArray = EncodeGlyphPositioningArray(embedded, span.GlyphSpan, pdfFontSize, pdfCharacterSpacing, forcePositioningArray: true);
        IReadOnlyList<TextGlyphAtom> glyphs = span.GlyphSpan.Glyphs
            .Select(glyph => new TextGlyphAtom(glyph.CodePoint, glyph.Typeface, glyph.TypefaceResolutionSource, glyph.GlyphId, glyph.Advance, glyph.AdjustmentBefore))
            .ToArray();
        return new TextGlyphRun(run, resourceName, embedded, glyphHex, positioningArray, glyphs, x, baselineY, lineWidth, pdfFontSize, pdfCharacterSpacing, syntheticBold, syntheticItalic);

        PptxPdfTextEmissionContext CreatePdfTextEmissionContext()
        {
            TextRun emissionRun = span.Run;
            return new PptxPdfTextEmissionContext(
                emissionRun.FontSize,
                emissionRun.Y + emissionRun.BaselineOffset,
                span.FrameIndex,
                span.ParagraphIndex,
                span.LineIndex,
                span.SpanIndex,
                span.LineSpanCount,
                span.FrameFontScale,
                span.FrameShapeX,
                span.FrameShapeTopY,
                span.FrameShapeWidth,
                span.FrameShapeHeight,
                span.TableRowIndex,
                span.TableColumnIndex,
                span.TableRowSpan,
                span.TableColumnSpan,
                span.FrameInsetLeft,
                span.FrameInsetRight,
                span.FrameInsetTop,
                span.FrameInsetBottom,
                span.FrameWrapMode,
                span.FrameWrapValue,
                span.FrameVerticalOverflowMode,
                span.FrameVerticalOverflowValue,
                span.FrameVerticalOverflowSource,
                span.FrameAutofitMode,
                span.FrameTextX,
                span.FrameTextWidth,
                span.FrameTextWrapWidth,
                span.FrameTextHeight,
                span.FrameClipX,
                span.FrameClipWidth,
                span.FrameClipY,
                span.FrameClipHeight,
                span.FrameColumnCount,
                span.FrameColumnSpacing,
                span.LineBox?.TopY ?? emissionRun.Y,
                span.LineBox?.Advance ?? 0d,
                span.LineBox?.MaxFontSize ?? emissionRun.FontSize);
        }
    }

    private static string EncodeGlyphHex(PdfEmbeddedFont embedded, PptxTextGlyphSpanLayout span)
    {
        var builder = new StringBuilder(span.Glyphs.Count * 4);
        foreach (PptxTextGlyphLayout glyph in span.Glyphs)
        {
            if (embedded.TryGetEncodedCid(glyph.GlyphId, out ushort cid))
            {
                builder.Append(cid.ToString("X4", CultureInfo.InvariantCulture));
            }
        }

        return builder.ToString();
    }

    private static string? EncodeGlyphPositioningArray(PdfEmbeddedFont embedded, PptxTextGlyphSpanLayout span, double pdfFontSize, double pdfCharacterSpacing, bool forcePositioningArray)
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
                    glyph.AdjustmentBefore - pdfCharacterSpacing,
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

            if (embedded.TryGetEncodedCid(glyph.GlyphId, out ushort cid))
            {
                builder.Append('<').Append(cid.ToString("X4", CultureInfo.InvariantCulture)).Append('>');
            }
        }

        builder.Append(']');
        return hasPositioning || forcePositioningArray ? builder.ToString() : null;
    }

    private static string? EncodeGlyphPositioningArray(PdfEmbeddedFont embedded, IReadOnlyList<TextGlyphAtom> glyphs, double layoutFontSize, double pdfFontSize, double characterSpacing, bool forcePositioningArray)
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
                    glyph.AdjustmentBefore - characterSpacing,
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

            if (embedded.TryGetEncodedCid(glyph.GlyphId, out ushort cid))
            {
                builder.Append('<').Append(cid.ToString("X4", CultureInfo.InvariantCulture)).Append('>');
            }
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

    private static void DrawHighlightRectangle(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run, RgbColor highlight, double baselineY, double lineWidth)
    {
        if (!TryGetHighlightRectangle(embedded, run, baselineY, lineWidth, out TextHighlightRectangle rectangle))
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
        graphics.FillRectangleEvenOdd(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        graphics.RestoreState();
    }

    private static bool TryGetHighlightRectangle(PdfEmbeddedFont embedded, TextRun run, double baselineY, double lineWidth, out TextHighlightRectangle rectangle)
    {
        rectangle = default;
        if (!BaselineIntersectsClip(run, baselineY))
        {
            return false;
        }

        double fontScale = run.FontSize / embedded.Font.UnitsPerEm;
        double highlightDescent = PptxTextMetricRules.HighlightDescent(embedded, run.FontSize, fontScale);
        double highlightHeight = PptxTextMetricRules.HighlightHeight(embedded, run.FontSize, fontScale);
        double highlightY = baselineY - highlightDescent;
        rectangle = new TextHighlightRectangle(run.X, highlightY, lineWidth, highlightHeight);
        return true;
    }

    private static bool TryGetUnderlineRectangle(PdfEmbeddedFont embedded, TextGlyphRun glyphRun, out TextDecorationRectangle rectangle)
    {
        TextRun run = glyphRun.Source;
        double underlineScale = run.FontSize / embedded.Font.UnitsPerEm;
        double underlineThickness = PptxTextMetricRules.UnderlineThickness(embedded, run.FontSize);
        double underlineTopY = glyphRun.BaselineY + embedded.Font.Post.UnderlinePosition * underlineScale;
        double underlineY = underlineTopY - underlineThickness;
        rectangle = new TextDecorationRectangle(glyphRun.X, underlineY, glyphRun.Width, underlineThickness);
        return glyphRun.Width > PptxTextMetricRules.TextStateTolerance && underlineThickness > 0d;
    }

    private static bool TryGetStrikeRectangle(PdfEmbeddedFont embedded, TextGlyphRun glyphRun, out TextDecorationRectangle rectangle)
    {
        TextRun run = glyphRun.Source;
        rectangle = new TextDecorationRectangle(
            glyphRun.X,
            PptxTextMetricRules.StrikeY(embedded, glyphRun.BaselineY, run.FontSize),
            glyphRun.Width,
            PptxTextMetricRules.StrikeThickness(embedded, run.FontSize));
        return glyphRun.Width > PptxTextMetricRules.TextStateTolerance && rectangle.Height > 0d;
    }

    private static void FillTextDecorationRectangleEvenOdd(PdfGraphicsBuilder graphics, TextDecorationRectangle rectangle)
    {
        double x = rectangle.X;
        double midX = rectangle.X + rectangle.Width / 2d;
        double right = rectangle.X + rectangle.Width;
        double bottom = rectangle.Y;
        double top = rectangle.Y + rectangle.Height;

        graphics.MoveTo(x, top);
        graphics.LineTo(midX, top);
        graphics.LineTo(right, top);
        graphics.LineTo(right, bottom);
        graphics.LineTo(midX, bottom);
        graphics.LineTo(x, bottom);
        graphics.ClosePath();
        graphics.FillCurrentPathEvenOdd();
    }

    private static bool BaselineIntersectsClip(TextRun run, double baselineY)
    {
        if (HasTextTransform(run))
        {
            return true;
        }

        if (!run.StrictClip)
        {
            return true;
        }

        return baselineY + run.FontSize >= run.ClipY - PptxTextMetricRules.TextStateTolerance &&
            baselineY - run.FontSize <= run.ClipY + run.ClipHeight + PptxTextMetricRules.TextStateTolerance;
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
        if (ShouldDrawGlyphOutlinePath(glyphRun))
        {
            DrawGlyphOutlinePath();
            return;
        }

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
                characterSpacing: glyphRun.PdfCharacterSpacing,
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
                characterSpacing: glyphRun.PdfCharacterSpacing,
                textRenderingMode: TextRenderingMode(glyphRun),
                strokeRed: TextStrokeColor(glyphRun).Red,
                strokeGreen: TextStrokeColor(glyphRun).Green,
                strokeBlue: TextStrokeColor(glyphRun).Blue,
                strokeWidth: TextStrokeWidth(glyphRun));
        }

        void DrawGlyphOutlinePath()
        {
            TextRun outlineRun = glyphRun.Source;
            graphics.SetFillRgb(outlineRun.Color.Red, outlineRun.Color.Green, outlineRun.Color.Blue);
            double shear = glyphRun.SyntheticItalic ? PdfGraphicsBuilder.SyntheticItalicShear : 0d;

            double cursorX = glyphRun.X;
            bool hasPath = false;
            for (int i = 0; i < glyphRun.Glyphs.Count; i++)
            {
                TextGlyphAtom glyph = glyphRun.Glyphs[i];
                if (i > 0)
                {
                    cursorX += glyph.AdjustmentBefore;
                }

                if (PdfGlyphOutlinePath.TryAppendGlyphPath(
                    graphics,
                    glyphRun.Font.Font,
                    glyph.GlyphId,
                    cursorX,
                    glyphRun.BaselineY,
                    glyphRun.PdfFontSize,
                    shear))
                {
                    hasPath = true;
                }

                cursorX += glyph.Advance;
            }

            if (hasPath)
            {
                graphics.FillCurrentPath();
            }
        }
    }

    private static bool ShouldDrawGlyphOutlinePath(TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        if (run.Alpha >= 1d - PptxTextMetricRules.TextStateTolerance ||
            run.Outline is not null ||
            glyphRun.Glyphs.Count == 0)
        {
            return false;
        }

        bool hasDrawableGlyph = false;
        foreach (TextGlyphAtom glyph in glyphRun.Glyphs)
        {
            if (!glyphRun.Font.Font.TryReadGlyphOutline(glyph.GlyphId, out var outline) ||
                outline.Contours.Count == 0)
            {
                if (IsAdvanceOnlyGlyph(glyph))
                {
                    continue;
                }

                return false;
            }

            hasDrawableGlyph = true;
        }

        return hasDrawableGlyph;

        bool IsAdvanceOnlyGlyph(TextGlyphAtom glyph)
        {
            return Rune.IsWhiteSpace(new Rune(glyph.CodePoint));
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
