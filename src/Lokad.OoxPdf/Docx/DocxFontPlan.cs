using Lokad.OoxPdf.Fonts;

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
                .SelectMany(cell => cell.Paragraphs))
            .SelectMany(paragraph => paragraph.Runs)
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

    private static IReadOnlyList<string> DistinctFamilies(params string?[] families)
    {
        return families
            .Where(family => !string.IsNullOrWhiteSpace(family))
            .Select(family => family!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }
}
