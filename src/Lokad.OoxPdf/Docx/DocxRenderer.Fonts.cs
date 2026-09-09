using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static DocxFontResources PrepareFontResources(DocxDocument document, IFontResolver fontResolver, Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxFontPlan plan = DocxFontPlan.Create(document, fontResolver, cancellationToken);
        var resources = new List<PdfFontResource>();
        var runResources = new Dictionary<DocxTextRun, DocxRunFontResource>();
        var fontCache = new Dictionary<(string StableId, int FaceIndex), OpenTypeFont?>();
        var reportedUnembeddableFaces = new HashSet<(string StableId, int FaceIndex)>();
        plan = SubstituteUnembeddableFonts(plan, fontCache, diagnosticSink, reportedUnembeddableFaces, cancellationToken);
        PrepareResolvedRunFontResources(plan, resources, runResources, fontCache, cancellationToken);
        DocxRunFontResource? fallback = PrepareFallbackFontResource(plan, fontResolver, resources, runResources, fontCache, diagnosticSink, reportedUnembeddableFaces, cancellationToken);
        IReadOnlyDictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>> fallbackChains = PreparePerCharacterFallbackResources(plan, fontResolver, fallback?.Resolution, resources, runResources, fontCache, diagnosticSink, reportedUnembeddableFaces, cancellationToken);
        IDocxTextMeasurer? measurer = plan.Runs.Any(run => LoadFont(run.Resolution, fontCache, cancellationToken) is not null) || fallback is not null
            ? new DocxFontPlanTextMeasurer(plan, fallback?.Resolution, cancellationToken, fontResolver, fontCache)
            : null;
        return new DocxFontResources(plan, measurer, resources, runResources, fallback, fallbackChains);
    }

    // CFF/OpenType-CFF fonts have valid metrics but no TrueType outlines, so the
    // PDF subsetter cannot embed them. Remap those runs to unresolved before any
    // measurement or emission so both paths share the document fallback face (F03:
    // measured and emitted glyphs use the same face) and report each substituted
    // typeface once through the diagnostic sink.
    private static DocxFontPlan SubstituteUnembeddableFonts(
        DocxFontPlan plan,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported,
        CancellationToken cancellationToken)
    {
        List<DocxResolvedRunTypeface>? remapped = null;
        for (int i = 0; i < plan.Runs.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxResolvedRunTypeface run = plan.Runs[i];
            if (run.Resolution is not { } resolution ||
                LoadFont(resolution, fontCache, cancellationToken) is not { } font ||
                font.HasTrueTypeOutlines)
            {
                if (remapped is not null)
                {
                    remapped.Add(run);
                }

                continue;
            }

            if (remapped is null)
            {
                remapped = new List<DocxResolvedRunTypeface>(plan.Runs.Count);
                for (int j = 0; j < i; j++)
                {
                    remapped.Add(plan.Runs[j]);
                }
            }

            ReportUnembeddableFace(resolution, "substituted with the document fallback typeface", "Document fallback typeface", diagnosticSink, reported);

            remapped.Add(run with { Resolution = null });
        }

        return remapped is null ? plan : new DocxFontPlan(remapped);
    }

    // Reports an unembeddable (CFF/OpenType-CFF) face once per conversion.
    // Callers describe what took its place.
    private static void ReportUnembeddableFace(
        FontFaceResolution resolution,
        string action,
        string fallback,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported)
    {
        if (!reported.Add((resolution.Source.StableId, resolution.FontFaceIndex)))
        {
            return;
        }

        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            "FONT_UNSUPPORTED_OUTLINES",
            OoxPdfSeverity.Warning,
            "Font '" + resolution.FamilyName + "' has no embeddable TrueType outlines (CFF/OpenType-CFF) and was " + action + ".",
            PartName: null,
            SlideIndex: null,
            PageIndex: null,
            Feature: resolution.FamilyName,
            Fallback: fallback));
    }

    private sealed class ScaledDocxTextMeasurer(IDocxTextMeasurer inner, double textScale, double lineMetricScale) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            return inner.MeasureText(run, text, fontSize) * textScale;
        }

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxLineMetricsProvider lineMetrics
                ? lineMetrics.MeasureSingleLineHeight(run, fontSize) * lineMetricScale
                : fontSize * 1.2d * lineMetricScale;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsAscender(run, fontSize) * lineMetricScale
                : fontSize * lineMetricScale;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsDescender(run, fontSize) * lineMetricScale
                : fontSize * 0.2d * lineMetricScale;
        }
    }

    private static void PrepareResolvedRunFontResources(
        DocxFontPlan plan,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        var resolvedRuns = new List<(DocxResolvedRunTypeface Run, FontFaceResolution Resolution)>();
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Resolution is not { } resolution || LoadFont(resolution, fontCache, cancellationToken) is null)
            {
                continue;
            }
            resolvedRuns.Add((run, resolution));
        }

        foreach (IGrouping<(string StableId, int FaceIndex), (DocxResolvedRunTypeface Run, FontFaceResolution Resolution)> group in resolvedRuns.GroupBy(item => (item.Resolution.Source.StableId, item.Resolution.FontFaceIndex)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            FontFaceResolution resolution = group.First().Resolution;
            IReadOnlyList<int> glyphs = CollectRunGlyphs(group.Select(item => item.Run), cancellationToken);
            if (glyphs.Count == 0)
            {
                continue;
            }

            OpenTypeFont? font = LoadFont(resolution, fontCache, cancellationToken);
            if (font is null)
            {
                continue;
            }

            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs, cancellationToken);
            string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            var runResource = new DocxRunFontResource(name, embedded, resolution);
            resources.Add(new PdfFontResource(name, embedded));
            foreach (DocxResolvedRunTypeface run in group.Select(item => item.Run))
            {
                cancellationToken.ThrowIfCancellationRequested();
                runResources[run.Run] = runResource;
            }
        }
    }

    private static DocxRunFontResource? PrepareFallbackFontResource(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported,
        CancellationToken cancellationToken)
    {
        FontFaceResolution resolution = ResolveDocumentBaseFont(plan, fontResolver, fontCache, cancellationToken);
        OpenTypeFont? font = LoadFont(resolution, fontCache, cancellationToken);
        if (font is null)
        {
            return null;
        }

        if (!font.HasTrueTypeOutlines)
        {
            ReportUnembeddableFace(resolution, "skipped as the document fallback typeface", "Missing-font path", diagnosticSink, reported);
            return null;
        }

        DocxResolvedRunTypeface[] fallbackRuns = plan.Runs
            .Where(run => !runResources.ContainsKey(run.Run))
            .ToArray();
        if (fallbackRuns.Length == 0)
        {
            return null;
        }

        IReadOnlyList<int> glyphs = CollectRunGlyphs(fallbackRuns, cancellationToken);
        if (glyphs.Count == 0)
        {
            return null;
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs, cancellationToken);
        string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
        var runResource = new DocxRunFontResource(name, embedded, resolution);
        resources.Add(new PdfFontResource(name, embedded));
        foreach (DocxResolvedRunTypeface run in fallbackRuns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            runResources[run.Run] = runResource;
        }

        return runResource;
    }

    private static IReadOnlyDictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>> PreparePerCharacterFallbackResources(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        FontFaceResolution? documentFallback,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<(string StableId, int FaceIndex)> reported,
        CancellationToken cancellationToken)
    {
        var chains = new Dictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>>();
        var fallbackCodepoints = new Dictionary<(string StableId, int FaceIndex), HashSet<int>>();
        var fallbackFonts = new Dictionary<(string StableId, int FaceIndex), OpenTypeFont>();
        var fallbackResolutions = new Dictionary<(string StableId, int FaceIndex), FontFaceResolution>();
        var runFallbacks = new Dictionary<DocxTextRun, (OpenTypeFont PrimaryFont, IReadOnlyList<OpenTypeFont> Fonts, IReadOnlyList<FontFaceResolution> Resolutions, IReadOnlyList<FontCoverageSpan> Spans)>();
        foreach (DocxResolvedRunTypeface resolved in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(resolved.Run.Text))
            {
                continue;
            }

            if (!runResources.TryGetValue(resolved.Run, out DocxRunFontResource? primaryResource))
            {
                continue;
            }

            OpenTypeFont? primaryFont = LoadFont(primaryResource.Resolution, fontCache, cancellationToken);
            if (primaryFont is null)
            {
                continue;
            }

            // Primary-covered runs need no fallback candidates; skipping their
            // program loads leaves split outcomes unchanged (G04).
            if (FontCoverageFallback.IsFullyCovered(resolved.Run.Text, primaryFont, cancellationToken))
            {
                continue;
            }

            var fonts = new List<OpenTypeFont>();
            var resolutions = new List<FontFaceResolution>();
            foreach (FontFaceResolution candidate in DocxFontFallbackRules.ResolveCandidateResolutions(fontResolver, resolved.Run, resolved.ResolvedFamily, primaryResource.Resolution, documentFallback))
            {
                cancellationToken.ThrowIfCancellationRequested();
                OpenTypeFont? candidateFont = fonts.Count == 0 ? primaryFont : LoadFont(candidate, fontCache, cancellationToken);
                if (candidateFont is null || fonts.Contains(candidateFont))
                {
                    continue;
                }

                if (fonts.Count != 0 && !candidateFont.HasTrueTypeOutlines)
                {
                    ReportUnembeddableFace(candidate, "skipped for per-character fallback", "Next covering typeface", diagnosticSink, reported);
                    continue;
                }

                fonts.Add(candidateFont);
                resolutions.Add(candidate);
            }

            if (fonts.Count <= 1)
            {
                continue;
            }

            IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(resolved.Run.Text, fonts, cancellationToken);
            bool needsFallback = false;
            foreach (FontCoverageSpan span in spans)
            {
                if (span.FontIndex > 0)
                {
                    needsFallback = true;
                    break;
                }
            }

            if (!needsFallback)
            {
                continue;
            }

            runFallbacks[resolved.Run] = (primaryFont, fonts, resolutions, spans);
            for (int i = 0; i < spans.Count; i++)
            {
                FontCoverageSpan span = spans[i];
                if (span.FontIndex <= 0)
                {
                    continue;
                }

                FontFaceResolution spanResolution = resolutions[span.FontIndex];
                var key = (spanResolution.Source.StableId, spanResolution.FontFaceIndex);
                if (!fallbackCodepoints.TryGetValue(key, out HashSet<int>? codepoints))
                {
                    codepoints = new HashSet<int>();
                    fallbackCodepoints[key] = codepoints;
                    fallbackFonts[key] = fonts[span.FontIndex];
                    fallbackResolutions[key] = spanResolution;
                }

                foreach (Rune rune in resolved.Run.Text.Substring(span.Start, span.Length).EnumerateRunes())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    codepoints.Add(rune.Value);
                }
            }
        }

        if (runFallbacks.Count == 0)
        {
            return chains;
        }

        var fallbackResources = new Dictionary<(string StableId, int FaceIndex), DocxRunFontResource>();
        foreach (var key in fallbackCodepoints.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var glyphs = new HashSet<int>(fallbackCodepoints[key]);
            foreach (Rune rune in " 0123456789".EnumerateRunes())
            {
                glyphs.Add(rune.Value);
            }

            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(fallbackFonts[key], glyphs, cancellationToken);
            string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            var resource = new DocxRunFontResource(name, embedded, fallbackResolutions[key]);
            resources.Add(new PdfFontResource(name, embedded));
            fallbackResources[key] = resource;
        }

        foreach (DocxTextRun run in runFallbacks.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var entry = runFallbacks[run];
            var chain = new List<DocxFallbackFontEntry> { new DocxFallbackFontEntry(entry.PrimaryFont, runResources[run]) };
            for (int i = 1; i < entry.Fonts.Count; i++)
            {
                bool used = false;
                foreach (FontCoverageSpan span in entry.Spans)
                {
                    if (span.FontIndex == i)
                    {
                        used = true;
                        break;
                    }
                }

                if (!used)
                {
                    continue;
                }

                FontFaceResolution spanResolution = entry.Resolutions[i];
                var key = (spanResolution.Source.StableId, spanResolution.FontFaceIndex);
                chain.Add(new DocxFallbackFontEntry(entry.Fonts[i], fallbackResources[key]));
            }

            chains[run] = chain;
        }

        return chains;
    }
    private static IReadOnlyList<int> CollectRunGlyphs(IEnumerable<DocxResolvedRunTypeface> runs, CancellationToken cancellationToken)
    {
        var glyphs = new HashSet<int>();
        foreach (DocxResolvedRunTypeface run in runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (Rune rune in run.Run.Text.EnumerateRunes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                glyphs.Add(rune.Value);
            }
        }

        foreach (Rune rune in " 0123456789".EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            glyphs.Add(rune.Value);
        }

        return glyphs.ToArray();
    }

    private static FontFaceResolution ResolveDocumentBaseFont(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Resolution is { } fontResolution && LoadFont(fontResolution, fontCache, cancellationToken) is not null)
            {
                return fontResolution;
            }
        }

        return DocxFontFallbackRules.ResolveDefaultDocumentTypeface(fontResolver, false, false);
    }

    private static OpenTypeFont? LoadFont(
        FontFaceResolution? resolution,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (resolution is null)
        {
            return null;
        }

        var key = (resolution.Source.StableId, resolution.FontFaceIndex);
        if (fontCache.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        cached = FontProgramLoader.Load(resolution, cancellationToken);
        fontCache[key] = cached;
        return cached;
    }
}
