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
    /// Whether the field is dynamic (page-number family): PAGE or NUMPAGES.
    /// Centralizes the predicate previously repeated at each counter site.
    /// </summary>
    public static bool IsDynamic(this DocxFieldKind kind)
    {
        return kind is DocxFieldKind.Page or DocxFieldKind.NumPages;
    }
}
