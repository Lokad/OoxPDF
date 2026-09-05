namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed property family of a DOCX formatting revision: which property group changed.
/// Closed set — the six w:CT_TrackChange property-change elements (rPrChange through
/// sectPrChange); the reader assigns these at the boundary from element names (null when
/// not a property change), so
/// counters, balloon labels, and property priorities compare enum members instead of
/// raw strings. Formatting is the derived fallback for non-property revisions (never
/// stored on the model, only produced by family resolution for balloon labels). Snapshots stay string-mapped via ToValueString; open OOXML element-name
/// vocabularies (propertyElementNames, priority values) stay string.
/// </summary>
internal enum DocxRevisionPropertyFamily
{
    Run,
    Paragraph,
    Table,
    Row,
    Cell,
    Section,
    Formatting
}

internal static class DocxRevisionPropertyFamilyExtensions
{
    /// <summary>
    /// Stable spelling of a revision property family. Kept as the historical
    /// PascalCase strings so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxRevisionPropertyFamily family)
    {
        return family switch
        {
            DocxRevisionPropertyFamily.Run => "Run",
            DocxRevisionPropertyFamily.Paragraph => "Paragraph",
            DocxRevisionPropertyFamily.Table => "Table",
            DocxRevisionPropertyFamily.Row => "Row",
            DocxRevisionPropertyFamily.Cell => "Cell",
            DocxRevisionPropertyFamily.Section => "Section",
            DocxRevisionPropertyFamily.Formatting => "formatting",
            _ => throw new ArgumentOutOfRangeException(nameof(family), family, "Unknown revision property family.")
        };
    }
}