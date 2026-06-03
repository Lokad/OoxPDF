namespace Lokad.OoxPdf.Docx;

internal sealed record DocxTypefaceCandidates(
    string? Primary,
    string? Alternate,
    string? Theme);

internal static class DocxFontResolver
{
    public static DocxTypefaceCandidates ResolveLatinTypeface(DocxTextRun run, DocxFontCatalog catalog)
    {
        DocxEffectiveRunProperties effective = run.EffectiveProperties;
        bool complexScript = DocxScriptClassifier.IsComplexScriptText(run.Text);
        string? primary = complexScript
            ? FirstNonEmpty(effective.Fonts.ComplexScript, effective.FontFamily, effective.Fonts.Ascii, effective.Fonts.HighAnsi)
            : FirstNonEmpty(effective.Fonts.Ascii, effective.Fonts.HighAnsi, effective.FontFamily);
        string? alternate = ResolveAlternate(primary, catalog);
        string? theme = complexScript
            ? ResolveThemeTypeface(FirstNonEmpty(effective.Fonts.ComplexScriptTheme, effective.Fonts.AsciiTheme, effective.Fonts.HighAnsiTheme), catalog.ThemeFonts)
            : ResolveThemeTypeface(FirstNonEmpty(effective.Fonts.AsciiTheme, effective.Fonts.HighAnsiTheme), catalog.ThemeFonts);
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
            "majorBidi" => themeFonts.MajorComplexScriptTypeface,
            "minorBidi" => themeFonts.MinorComplexScriptTypeface,
            _ => null
        };
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }
}
