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
    private static DocxFontResources PrepareFontResources(DocxDocument document, IFontResolver fontResolver, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxFontPlan plan = DocxFontPlan.Create(document, fontResolver, cancellationToken);
        var resources = new List<PdfFontResource>();
        var runResources = new Dictionary<DocxTextRun, DocxRunFontResource>();
        var fontCache = new Dictionary<(string StableId, int FaceIndex), OpenTypeFont?>();
        PrepareResolvedRunFontResources(plan, resources, runResources, fontCache, cancellationToken);
        DocxRunFontResource? fallback = PrepareFallbackFontResource(plan, fontResolver, resources, runResources, fontCache, cancellationToken);
        IDocxTextMeasurer? measurer = plan.Runs.Any(run => LoadFont(run.Resolution, fontCache, cancellationToken) is not null) || fallback is not null
            ? new DocxFontPlanTextMeasurer(plan, fallback?.Resolution, cancellationToken)
            : null;
        return new DocxFontResources(plan, measurer, resources, runResources, fallback);
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
        CancellationToken cancellationToken)
    {
        FontFaceResolution resolution = ResolveDocumentBaseFont(plan, fontResolver, fontCache, cancellationToken);
        OpenTypeFont? font = LoadFont(resolution, fontCache, cancellationToken);
        if (font is null)
        {
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
