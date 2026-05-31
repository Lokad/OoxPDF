namespace Lokad.OoxPdf.Docx;

internal sealed record DocxTypefaceCandidates(
    string? Primary,
    string? Alternate,
    string? Theme);

internal static class DocxFontResolver
{
    public static DocxTypefaceCandidates ResolveLatinTypeface(DocxTextRun run, DocxFontCatalog catalog)
    {
        string? primary = FirstNonEmpty(run.Fonts.Ascii, run.Fonts.HighAnsi, run.FontFamily);
        string? alternate = ResolveAlternate(primary, catalog);
        string? theme = ResolveThemeTypeface(FirstNonEmpty(run.Fonts.AsciiTheme, run.Fonts.HighAnsiTheme), catalog.ThemeFonts);
        return new DocxTypefaceCandidates(primary, alternate, theme);
    }

    private static string? ResolveAlternate(string? primary, DocxFontCatalog catalog)
    {
        if (string.IsNullOrWhiteSpace(primary))
        {
            return null;
        }

        return catalog.Entries
            .FirstOrDefault(entry => entry.Name.Equals(primary, StringComparison.OrdinalIgnoreCase))
            ?.AlternateName;
    }

    private static string? ResolveThemeTypeface(string? themeToken, DocxThemeFonts themeFonts)
    {
        return themeToken switch
        {
            "majorAscii" or "majorHAnsi" => themeFonts.MajorLatinTypeface,
            "minorAscii" or "minorHAnsi" => themeFonts.MinorLatinTypeface,
            _ => null
        };
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }
}
