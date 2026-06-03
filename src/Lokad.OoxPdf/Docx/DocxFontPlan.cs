using Lokad.OoxPdf.Fonts;
using System.Text;

namespace Lokad.OoxPdf.Docx;

internal enum DocxTypefaceResolutionSource
{
    Primary,
    FontTableAlternate,
    Theme,
    ResolverFallback,
    Missing
}

internal sealed record DocxResolvedRunTypeface(
    DocxTextRun Run,
    IReadOnlyList<string> CandidateFamilies,
    string? RequestedFamily,
    string? ResolvedFamily,
    DocxTypefaceResolutionSource Source,
    FontResolution? Resolution);

internal sealed record DocxFontPlan(IReadOnlyList<DocxResolvedRunTypeface> Runs)
{
    public static DocxFontPlan Create(DocxDocument document, IFontResolver fontResolver)
    {
        IReadOnlyList<DocxTextRun> runs = DocxBlockTraversal.EnumerateBodyParagraphs(document)
            .Concat(GetReferencedHeaderFooterParagraphs(document.HeaderBodyElementsByType, document.HeaderParagraphsByType, document.HeaderParagraphs))
            .Concat(GetReferencedHeaderFooterParagraphs(document.FooterBodyElementsByType, document.FooterParagraphsByType, document.FooterParagraphs))
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.PageSettings))
            .Concat(document.BodyElements
                .OfType<DocxSectionBreakElement>()
                .SelectMany(sectionBreak => DocxBlockTraversal.EnumerateStaticStoryParagraphs(sectionBreak.PageSettings)))
            .Concat(document.RelatedStories.SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs))
            .SelectMany(GetParagraphFontRuns)
            .ToArray();

        return new DocxFontPlan(runs
            .Select(run => ResolveRunTypeface(run, document.FontCatalog, fontResolver))
            .ToArray());
    }

    private static IEnumerable<DocxParagraph> GetReferencedHeaderFooterParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        IReadOnlyList<DocxParagraph> fallbackParagraphs)
    {
        if (bodyElementsByType.Count != 0)
        {
            return bodyElementsByType.Values.SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs);
        }

        return paragraphsByType.Count == 0
            ? fallbackParagraphs
            : paragraphsByType.Values.SelectMany(paragraphs => paragraphs);
    }

    private static DocxResolvedRunTypeface ResolveRunTypeface(DocxTextRun run, DocxFontCatalog fontCatalog, IFontResolver fontResolver)
    {
        DocxTypefaceCandidates candidates = DocxFontResolver.ResolveLatinTypeface(run, fontCatalog);
        IReadOnlyList<string> families = DistinctFamilies(candidates.Primary, candidates.Alternate, candidates.Theme);
        if (families.Count == 0)
        {
            return new DocxResolvedRunTypeface(run, families, null, null, DocxTypefaceResolutionSource.Missing, null);
        }

        for (int i = 0; i < families.Count; i++)
        {
            string family = families[i];
            FontResolution resolution = fontResolver.Resolve(new FontRequest(family, run.Bold, run.Italic));
            if (!resolution.IsFallback)
            {
                return new DocxResolvedRunTypeface(run, families, family, resolution.FamilyName, SourceForCandidate(candidates, family), resolution);
            }
        }

        string requested = families[0];
        FontResolution fallback = fontResolver.Resolve(new FontRequest(requested, run.Bold, run.Italic));
        return new DocxResolvedRunTypeface(run, families, requested, fallback.FamilyName, DocxTypefaceResolutionSource.ResolverFallback, fallback);
    }

    private static DocxTypefaceResolutionSource SourceForCandidate(DocxTypefaceCandidates candidates, string family)
    {
        if (EqualsCandidate(candidates.Primary, family))
        {
            return DocxTypefaceResolutionSource.Primary;
        }

        if (EqualsCandidate(candidates.Alternate, family))
        {
            return DocxTypefaceResolutionSource.FontTableAlternate;
        }

        return DocxTypefaceResolutionSource.Theme;
    }

    private static bool EqualsCandidate(string? candidate, string family)
    {
        return candidate is not null && candidate.Equals(family, StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<DocxTextRun> GetParagraphFontRuns(DocxParagraph paragraph)
    {
        foreach (DocxTextRun run in paragraph.Runs)
        {
            yield return run;
        }

        if (paragraph.ListLabel is not null)
        {
            DocxTextRun? firstRun = paragraph.Runs.FirstOrDefault();
            yield return DocxLayoutEngine.CreateListLabelRun(
                paragraph.ListLabel,
                firstRun,
                firstRun?.FontSize ?? 11d);
        }
    }

    private static IReadOnlyList<string> DistinctFamilies(params string?[] families)
    {
        return families
            .Where(family => !string.IsNullOrWhiteSpace(family))
            .Select(family => family!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }
}

internal sealed class DocxFontPlanTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
{
    private readonly IReadOnlyList<DocxResolvedRunTypeface> runs;
    private readonly FontResolution? fallbackResolution;
    private readonly Dictionary<(string Path, int FaceIndex), OpenTypeFont?> fonts = new();

    public DocxFontPlanTextMeasurer(DocxFontPlan plan, FontResolution? fallbackResolution = null)
    {
        runs = plan.Runs;
        this.fallbackResolution = fallbackResolution;
    }

    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontResolution resolution)
        {
            return 0d;
        }

        OpenTypeFont? font = LoadFont(resolution);
        if (font is null || font.UnitsPerEm == 0)
        {
            return 0d;
        }

        double units = 0d;
        ushort previousGlyph = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = font.MapCodePoint(rune.Value);
            if (previousGlyph != 0 && glyph != 0)
            {
                units += font.GetKerning(previousGlyph, glyph);
            }

            units += font.GetAdvanceWidth(glyph);
            previousGlyph = glyph;
        }

        return DocxTextSpacing.AddCharacterSpacing(units * fontSize / font.UnitsPerEm, run, text);
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if ((resolved?.Resolution ?? fallbackResolution) is not FontResolution resolution)
        {
            return fontSize;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? fontSize
            : DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font, fontSize);
    }

    public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
    {
        OpenTypeFont? font = ResolveFont(run);
        return font is null ? fontSize : DocxLineMetrics.MeasureWindowsAscender(font, fontSize);
    }

    public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
    {
        OpenTypeFont? font = ResolveFont(run);
        return font is null ? 0d : DocxLineMetrics.MeasureWindowsDescender(font, fontSize);
    }

    private DocxResolvedRunTypeface? ResolveRun(DocxTextRun? run)
    {
        if (run is null)
        {
            return null;
        }

        return runs.FirstOrDefault(resolved => resolved.Run.Equals(run));
    }

    private OpenTypeFont? ResolveFont(DocxTextRun? run)
    {
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        return (resolved?.Resolution ?? fallbackResolution) is FontResolution resolution
            ? LoadFont(resolution)
            : null;
    }

    private OpenTypeFont? LoadFont(FontResolution resolution)
    {
        if (string.IsNullOrWhiteSpace(resolution.FontFilePath))
        {
            return null;
        }

        var key = (resolution.FontFilePath, resolution.FontFaceIndex);
        if (fonts.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        OpenTypeFont? loaded = TryLoadFont(resolution.FontFilePath, resolution.FontFaceIndex);
        fonts[key] = loaded;
        return loaded;
    }

    private static OpenTypeFont? TryLoadFont(string path, int faceIndex)
    {
        try
        {
            return OpenTypeFont.Load(path, faceIndex);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
        {
            return null;
        }
    }
}

internal sealed record DocxFontPlanSnapshot(
    int RunCount,
    int PrimaryCount,
    int FontTableAlternateCount,
    int ThemeCount,
    int ResolverFallbackCount,
    int MissingCount,
    int DistinctCandidateFamilyCount,
    int DistinctResolvedFamilyCount,
    IReadOnlyList<DocxFontMetricBucketSnapshot> MetricBuckets)
{
    public static DocxFontPlanSnapshot FromPlan(DocxFontPlan plan)
    {
        return new DocxFontPlanSnapshot(
            plan.Runs.Count,
            Count(plan, DocxTypefaceResolutionSource.Primary),
            Count(plan, DocxTypefaceResolutionSource.FontTableAlternate),
            Count(plan, DocxTypefaceResolutionSource.Theme),
            Count(plan, DocxTypefaceResolutionSource.ResolverFallback),
            Count(plan, DocxTypefaceResolutionSource.Missing),
            plan.Runs
                .SelectMany(run => run.CandidateFamilies)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count(),
            plan.Runs
                .Select(run => run.ResolvedFamily)
                .Where(family => !string.IsNullOrWhiteSpace(family))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count(),
            ToMetricBuckets(plan));
    }

    private static int Count(DocxFontPlan plan, DocxTypefaceResolutionSource source)
    {
        return plan.Runs.Count(run => run.Source == source);
    }

    private static IReadOnlyList<DocxFontMetricBucketSnapshot> ToMetricBuckets(DocxFontPlan plan)
    {
        return plan.Runs
            .GroupBy(run => new
            {
                run.Source,
                FontSize = Math.Round(run.Run.FontSize, 3),
                ResolvedFamilyHash = HashFamily(run.ResolvedFamily)
            })
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key.Source.ToString(), StringComparer.Ordinal)
            .ThenBy(group => group.Key.FontSize)
            .Select(group => ToMetricBucketSnapshot(
                group.Key.Source.ToString(),
                group.Key.FontSize,
                group.Key.ResolvedFamilyHash,
                group.Count(),
                group.FirstOrDefault()?.Resolution))
            .ToArray();
    }

    private static DocxFontMetricBucketSnapshot ToMetricBucketSnapshot(
        string source,
        double fontSize,
        string? resolvedFamilyHash,
        int runCount,
        FontResolution? resolution)
    {
        OpenTypeFont? font = resolution is null ? null : TryLoadFont(resolution.FontFilePath, resolution.FontFaceIndex);
        if (font is null)
        {
            return new DocxFontMetricBucketSnapshot(
                source,
                fontSize,
                resolvedFamilyHash,
                runCount,
                UnitsPerEm: null,
                TypographicAscender: null,
                TypographicDescender: null,
                TypographicLineGap: null,
                WindowsAscender: null,
                WindowsDescender: null,
                SingleLineHeightPoints: null,
                WindowsAscenderPoints: null,
                WindowsDescenderPoints: null);
        }

        return new DocxFontMetricBucketSnapshot(
            source,
            fontSize,
            resolvedFamilyHash,
            runCount,
            font.UnitsPerEm,
            font.Os2.TypographicAscender,
            font.Os2.TypographicDescender,
            font.Os2.TypographicLineGap,
            font.Os2.WindowsAscender,
            font.Os2.WindowsDescender,
            DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font, fontSize),
            DocxLineMetrics.MeasureWindowsAscender(font, fontSize),
            DocxLineMetrics.MeasureWindowsDescender(font, fontSize));
    }

    private static OpenTypeFont? TryLoadFont(string? path, int faceIndex)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            return OpenTypeFont.Load(path, faceIndex);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    private static string? HashFamily(string? family)
    {
        if (string.IsNullOrWhiteSpace(family))
        {
            return null;
        }

        byte[] bytes = Encoding.UTF8.GetBytes(family.ToUpperInvariant());
        byte[] hash = System.Security.Cryptography.SHA256.HashData(bytes);
        return Convert.ToHexString(hash).Substring(0, 12).ToLowerInvariant();
    }
}

internal sealed record DocxFontMetricBucketSnapshot(
    string Source,
    double FontSize,
    string? ResolvedFamilyHash,
    int RunCount,
    int? UnitsPerEm,
    int? TypographicAscender,
    int? TypographicDescender,
    int? TypographicLineGap,
    int? WindowsAscender,
    int? WindowsDescender,
    double? SingleLineHeightPoints,
    double? WindowsAscenderPoints,
    double? WindowsDescenderPoints);
