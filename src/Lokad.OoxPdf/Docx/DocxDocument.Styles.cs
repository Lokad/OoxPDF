namespace Lokad.OoxPdf.Docx;

internal sealed record DocxStyleCatalog(
    bool HasRunDefaults,
    bool HasParagraphDefaults,
    string? DefaultTableStyleId,
    IReadOnlyList<DocxStyleDefinitionSummary> ParagraphStyles,
    IReadOnlyList<DocxStyleDefinitionSummary> CharacterStyles,
    IReadOnlyList<DocxTableStyleDefinitionSummary> TableStyles)
{
    public static DocxStyleCatalog Empty { get; } = new(false, false, null, [], [], []);
}

internal sealed record DocxStyleDefinitionSummary(
    string StyleId,
    string? BasedOnStyleId,
    bool HasParagraphProperties,
    bool HasRunProperties);

internal sealed record DocxTableStyleDefinitionSummary(
    string StyleId,
    string? BasedOnStyleId,
    bool HasTableProperties,
    bool HasCellProperties,
    bool HasParagraphProperties,
    bool HasRunProperties,
    int BorderCount,
    int ConditionalRegionCount);
