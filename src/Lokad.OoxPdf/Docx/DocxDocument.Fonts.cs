namespace Lokad.OoxPdf.Docx;

internal sealed record DocxFontCatalog(
    IReadOnlyList<DocxFontTableEntry> Entries,
    DocxThemeFonts ThemeFonts)
{
    public static DocxFontCatalog Empty { get; } = new([], DocxThemeFonts.Empty);
}

internal sealed record DocxFontTableEntry(
    string Name,
    string? AlternateName,
    string? FamilyValue,
    string? PitchValue,
    string? PanoseValue,
    string? CharsetValue);

internal sealed record DocxThemeFonts(
    string? MajorLatinTypeface,
    string? MinorLatinTypeface,
    string? MajorComplexScriptTypeface,
    string? MinorComplexScriptTypeface,
    string? MajorEastAsiaTypeface,
    string? MinorEastAsiaTypeface)
{
    public static DocxThemeFonts Empty { get; } = new(null, null, null, null, null, null);
}
