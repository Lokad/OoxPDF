namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed kind of a DOCX field reference: which field opcode produced it.
/// Closed set — the reader normalizes PAGE and NUMPAGES at the boundary and maps
/// everything else to Other, so downstream counters compare enum members
/// instead of raw strings.
/// </summary>
internal enum DocxFieldKind
{
    Page,
    NumPages,
    Other
}

internal static class DocxFieldKindExtensions
{
    /// <summary>
    /// Stable spelling of a field kind. Kept as the historical PascalCase
    /// strings so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxFieldKind kind)
    {
        return kind switch
        {
            DocxFieldKind.Page => "Page",
            DocxFieldKind.NumPages => "NumPages",
            DocxFieldKind.Other => "Other",
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown field kind.")
        };
    }

    /// <summary>
    /// Whether the field is dynamic (page-number family): PAGE or NUMPAGES.
    /// Centralizes the predicate previously repeated at each counter site.
    /// </summary>
    public static bool IsDynamic(this DocxFieldKind kind)
    {
        return kind is DocxFieldKind.Page or DocxFieldKind.NumPages;
    }
}
