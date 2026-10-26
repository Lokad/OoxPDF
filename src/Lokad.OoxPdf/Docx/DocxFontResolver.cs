namespace Lokad.OoxPdf.Docx;

internal sealed record DocxTypefaceCandidates(
    string? Primary,
    string? Alternate,
    string? Theme);

internal static class DocxFontResolver
{
    public static DocxTypefaceCandidates ResolveTypeface(DocxTextRun run, DocxFontCatalog catalog)
    {
        DocxEffectiveRunProperties effective = run.EffectiveProperties;
        DocxFontSlot slot = DocxScriptClassifier.FontSlot(run.Text);
        (string? primary, string? themeToken) = SelectSlot(effective.Fonts, effective.FontFamily, slot);
        string? alternate = ResolveAlternate(primary, catalog);
        return new DocxTypefaceCandidates(primary, alternate, ResolveThemeTypeface(themeToken, catalog.ThemeFonts));
    }

    public static bool SameTypefaceSelection(DocxRunFonts fonts, string? fontFamily, DocxFontSlot left, DocxFontSlot right)
    {
        // Complex-script fragments also own their bold/italic overrides. Other
        // slots need no artificial boundary when their declared choices agree.
        if ((left == DocxFontSlot.ComplexScript) != (right == DocxFontSlot.ComplexScript)) return false;
        var a = SelectSlot(fonts, fontFamily, left);
        var b = SelectSlot(fonts, fontFamily, right);
        return string.Equals(a.Family, b.Family, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(a.Theme, b.Theme, StringComparison.Ordinal);
    }

    private static (string? Family, string? Theme) SelectSlot(DocxRunFonts fonts, string? fontFamily, DocxFontSlot slot)
    {
        string? primary = slot switch {
            DocxFontSlot.ComplexScript => FirstNonEmpty(fonts.ComplexScript, fontFamily, fonts.Ascii, fonts.HighAnsi),
            DocxFontSlot.EastAsia => FirstNonEmpty(fonts.EastAsia, fontFamily, fonts.HighAnsi, fonts.Ascii),
            DocxFontSlot.HighAnsi => FirstNonEmpty(fonts.HighAnsi, fonts.Ascii, fontFamily),
            _ => FirstNonEmpty(fonts.Ascii, fonts.HighAnsi, fontFamily)
        };
        string? themeToken = slot switch {
            DocxFontSlot.ComplexScript => FirstNonEmpty(fonts.ComplexScriptTheme, fonts.AsciiTheme, fonts.HighAnsiTheme),
            DocxFontSlot.EastAsia => FirstNonEmpty(fonts.EastAsiaTheme, fonts.AsciiTheme, fonts.HighAnsiTheme),
            DocxFontSlot.HighAnsi => FirstNonEmpty(fonts.HighAnsiTheme, fonts.AsciiTheme),
            _ => FirstNonEmpty(fonts.AsciiTheme, fonts.HighAnsiTheme)
        };
        return (primary, themeToken);
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
            "majorEastAsia" => themeFonts.MajorEastAsiaTypeface,
            "minorEastAsia" => themeFonts.MinorEastAsiaTypeface,
            _ => null
        };
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }
}
