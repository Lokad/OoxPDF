using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Chart text parts render through separate calls whose resources merge into

    // one page-level list, so per-call numbering would collide (duplicate /Font
    // keys, last wins). Callers pass that shared list for page-unique sequencing;
    // the appended slice is already in the list, so callers must not re-add it.
    private static RenderedFonts RenderChartTextRuns(
        List<TextRun> runs,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> chartFonts,
        string resourcePrefix,
        PresentationFontResolver? fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (runs.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }
        runs = SplitRunsByResolvedTypeface(runs, fontResolver ?? new PresentationFontResolver(null), diagnosticSink, CancellationToken.None);

        RenderedFonts renderedFonts = CreateRenderedFonts(runs, fontResolver ?? new PresentationFontResolver(null), resourcePrefix, CancellationToken.None, diagnosticSink, chartFonts);
        DrawTextRunsWithFonts(runs, graphics, renderedFonts.Fonts);
        chartFonts.AddRange(renderedFonts.Resources);
        return renderedFonts;
    }

    // Legacy runs address embedded fonts by requested family, but CFF substitution
    // embeds fallback faces under fallback keys: without rewriting, substituted runs
    // would miss the lookup and vanish. Rewrite CFF-family runs through the same
    // per-glyph fallback the estimator measured with (F03). Families resolving to
    // embeddable-or-missing fonts keep legacy behavior exactly (missing fonts stay
    // invisible, as before).
    private static List<TextRun> SplitRunsByResolvedTypeface(
        IReadOnlyList<TextRun> runs,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        var rewritten = new List<TextRun>(runs.Count);
        var verdicts = new Dictionary<string, OpenTypeFont?>(StringComparer.OrdinalIgnoreCase);
        var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        TextAdvanceEstimator? estimator = null;
        foreach (TextRun run in runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string requestedFamily = PptxFontFallbackRules.ResolveDefaultLatinTypeface(run.FontFamily);
            string familyKey = FontKey(requestedFamily, run.Bold, run.Italic);
            if (!verdicts.TryGetValue(familyKey, out OpenTypeFont? primary))
            {
                primary = fontResolver.ResolvePresentationOpenTypeFont(new FontRequest(requestedFamily, run.Bold, run.Italic), cancellationToken)?.Font;
                verdicts[familyKey] = primary;
            }

            if (primary is null || primary.HasTrueTypeOutlines)
            {
                rewritten.Add(run);
                continue;
            }

            estimator ??= new TextAdvanceEstimator(fontResolver, cancellationToken);
            List<(string Typeface, string Text)> segments = SplitRunByFallbackTypeface(run, estimator, cancellationToken);
            if (segments.Count == 1)
            {
                rewritten.Add(run with { FontFamily = segments[0].Typeface });
            }
            else
            {
                double cursorX = run.X;
                foreach ((string typeface, string text) in segments)
                {
                    double width = estimator.Measure(text, run.FontSize, typeface, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
                    rewritten.Add(run with
                    {
                        Text = text,
                        X = cursorX,
                        Width = width,
                        FontFamily = typeface,
                        Alignment = TextAlignment.Left,
                        PreventCoalesce = true
                    });
                    cursorX += width;
                }
            }

            if (reported.Add(familyKey))
            {
                diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                    "FONT_UNSUPPORTED_OUTLINES",
                    OoxPdfSeverity.Warning,
                    "Font '" + requestedFamily + "' has no embeddable TrueType outlines (CFF/OpenType-CFF); affected runs were substituted with fallback typefaces.",
                    PartName: null,
                    SlideIndex: null,
                    PageIndex: null,
                    Feature: requestedFamily,
                    Fallback: "Per-glyph fallback typeface"));
            }
        }

        return rewritten;
    }

    private static List<(string Typeface, string Text)> SplitRunByFallbackTypeface(TextRun run, TextAdvanceEstimator estimator, CancellationToken cancellationToken)
    {
        var segments = new List<(string Typeface, string Text)>();
        string? segmentTypeface = null;
        var segmentText = new StringBuilder();
        void FlushSegment()
        {
            if (segmentText.Length != 0 && segmentTypeface is not null)
            {
                segments.Add((segmentTypeface, segmentText.ToString()));
                segmentText.Clear();
            }
        }

        foreach (Rune rune in run.Text.EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? typeface = estimator.ResolveGlyphFont(run.FontFamily, run.Bold, run.Italic, rune.Value)?.Typeface;
            if (typeface is null)
            {
                continue;
            }

            if (segmentTypeface is not null && !typeface.Equals(segmentTypeface, StringComparison.OrdinalIgnoreCase))
            {
                FlushSegment();
                segmentTypeface = null;
            }

            segmentTypeface = typeface;
            segmentText.Append(char.ConvertFromUtf32(rune.Value));
        }

        FlushSegment();
        return segments;
    }

    // Counts page-level names of the form prefix + digits so chart text parts
    // that render through separate calls share one numbering sequence per page.
    private static int CountPrefixedResourceNames(List<PdfFontResource>? nameScope, string prefix)
    {
        if (nameScope is null)
        {
            return 0;
        }

        string sanitizedPrefix = PdfEmbeddedFont.SanitizeName(prefix);
        int count = 0;
        foreach (PdfFontResource resource in nameScope)
        {
            string name = PdfEmbeddedFont.SanitizeName(resource.ResourceName);
            if (name.Length > sanitizedPrefix.Length &&
                name.StartsWith(sanitizedPrefix, StringComparison.Ordinal))
            {
                bool digitsOnly = true;
                for (int i = sanitizedPrefix.Length; i < name.Length; i++)
                {
                    if (!char.IsAsciiDigit(name[i]))
                    {
                        digitsOnly = false;
                        break;
                    }
                }

                if (digitsOnly)
                {
                    count++;
                }
            }
        }

        return count;
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


    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextRun> textRuns, PresentationFontResolver fontResolver, string resourcePrefix, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null, List<PdfFontResource>? nameScope = null)
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
            .ToArray(), fontResolver, resourcePrefix, cancellationToken, diagnosticSink, nameScope);
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextFontUse> uses, PresentationFontResolver fontResolver, string resourcePrefix, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null, List<PdfFontResource>? nameScope = null)
    {
        if (uses.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }

        uses = SubstituteUnembeddableFontUses(uses, fontResolver, diagnosticSink, cancellationToken);

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
            if (!font.HasTrueTypeOutlines)
            {
                continue;
            }

            PdfEmbeddedFont embedded = fontResolver.GetOrCreateSubset(resolution, font, group.SelectMany(use => use.CodePoints).ToArray(), cancellationToken);
            string resourceName = resourcePrefix + (CountPrefixedResourceNames(nameScope, resourcePrefix) + resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            fonts[group.Key] = new RenderedFont(resourceName, embedded, resolution, first.Bold && !resolution.Bold, first.Italic && !resolution.Italic);
            resources.Add(new PdfFontResource(resourceName, embedded));
        }

        return new RenderedFonts(fonts, resources);
    }

    // CFF/OpenType-CFF fonts have valid metrics but no TrueType outlines, so the
    // PDF subsetter cannot embed them. Rewrite affected uses to the fallback faces
    // the measurement estimator resolves for the same code points, so measured and
    // emitted glyphs use the same face (F03); report each substituted group once.
    private static IReadOnlyList<TextFontUse> SubstituteUnembeddableFontUses(
        IReadOnlyList<TextFontUse> uses,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        var groupFonts = new Dictionary<string, OpenTypeFont?>(StringComparer.OrdinalIgnoreCase);
        bool needsSubstitution = false;
        foreach (IGrouping<string, TextFontUse> group in uses.GroupBy(use => FontKey(use.FamilyName, use.Bold, use.Italic), StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            TextFontUse first = group.First();
            OpenTypeFont? font = fontResolver.ResolvePresentationOpenTypeFont(new FontRequest(first.FamilyName, first.Bold, first.Italic), cancellationToken)?.Font;
            groupFonts[group.Key] = font;
            if (font is not null && !font.HasTrueTypeOutlines)
            {
                needsSubstitution = true;
            }
        }

        if (!needsSubstitution)
        {
            return uses;
        }

        var estimator = new TextAdvanceEstimator(fontResolver, cancellationToken);
        var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var substituted = new List<TextFontUse>(uses.Count);
        foreach (TextFontUse use in uses)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (groupFonts[FontKey(use.FamilyName, use.Bold, use.Italic)] is not { } font || font.HasTrueTypeOutlines)
            {
                substituted.Add(use);
                continue;
            }
            foreach (IGrouping<string, (int CodePoint, string? Typeface)> split in use.CodePoints
                         .Select(codePoint => (CodePoint: codePoint, Typeface: estimator.ResolveGlyphFont(use.FamilyName, use.Bold, use.Italic, codePoint)?.Typeface))
                         .Where(resolved => !string.IsNullOrEmpty(resolved.Typeface))
                         .GroupBy(resolved => resolved.Typeface!, StringComparer.OrdinalIgnoreCase))
            {
                cancellationToken.ThrowIfCancellationRequested();
                substituted.Add(new TextFontUse(split.Key, use.Bold, use.Italic, split.Select(resolved => resolved.CodePoint).ToArray()));
            }

            if (reported.Add(FontKey(use.FamilyName, use.Bold, use.Italic)))
            {
                diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                    "FONT_UNSUPPORTED_OUTLINES",
                    OoxPdfSeverity.Warning,
                    "Font '" + use.FamilyName + "' has no embeddable TrueType outlines (CFF/OpenType-CFF); affected runs were substituted with fallback typefaces.",
                    PartName: null,
                    SlideIndex: null,
                    PageIndex: null,
                    Feature: use.FamilyName,
                    Fallback: "Per-glyph fallback typeface"));
            }
        }

        return substituted;
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

    private static void DrawTextSpansWithFonts(IReadOnlyList<PptxPositionedTextSpan> textSpans, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts, PptxTextHyperlinkScope? hyperlinkScope = null)
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
                    // Unresolved-font spans stay link-free: no glyphs are painted and the font layer reports the miss.
                    hyperlinkScope?.CollectEmissionSpan(emissionSpan, rendered.Font);
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
