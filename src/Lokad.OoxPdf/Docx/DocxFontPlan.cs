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
        IReadOnlyList<DocxTextRun> runs = document.Paragraphs
            .Concat(document.HeaderParagraphs)
            .Concat(document.FooterParagraphs)
            .Concat(document.Tables
                .SelectMany(table => table.Rows)
                .SelectMany(row => row.Cells)
                .SelectMany(GetCellParagraphs))
            .SelectMany(GetParagraphFontRuns)
            .ToArray();

        return new DocxFontPlan(runs
            .Select(run => ResolveRunTypeface(run, document.FontCatalog, fontResolver))
            .ToArray());
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

    private static IReadOnlyList<DocxParagraph> GetCellParagraphs(DocxTableCell cell)
    {
        return cell.Paragraphs.Count == 0 && cell.Text.Length != 0
            ? [new DocxParagraph([new DocxTextRun(cell.Text, 11d, null, false, false, false, null, null)], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null)]
            : cell.Paragraphs;
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

internal sealed class DocxFontPlanTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider
{
    private readonly IReadOnlyList<DocxResolvedRunTypeface> runs;
    private readonly Dictionary<(string Path, int FaceIndex), OpenTypeFont?> fonts = new();

    public DocxFontPlanTextMeasurer(DocxFontPlan plan)
    {
        runs = plan.Runs;
    }

    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if (resolved?.Resolution is not FontResolution resolution)
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

        return units * fontSize / font.UnitsPerEm;
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        DocxResolvedRunTypeface? resolved = ResolveRun(run);
        if (resolved?.Resolution is not FontResolution resolution)
        {
            return fontSize;
        }

        OpenTypeFont? font = LoadFont(resolution);
        return font is null
            ? fontSize
            : DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font, fontSize);
    }

    private DocxResolvedRunTypeface? ResolveRun(DocxTextRun? run)
    {
        if (run is null)
        {
            return null;
        }

        return runs.FirstOrDefault(resolved => resolved.Run.Equals(run));
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
    int DistinctResolvedFamilyCount)
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
                .Count());
    }

    private static int Count(DocxFontPlan plan, DocxTypefaceResolutionSource source)
    {
        return plan.Runs.Count(run => run.Source == source);
    }
}
