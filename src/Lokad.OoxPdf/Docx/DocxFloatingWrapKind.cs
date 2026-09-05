namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed wrap kind of a DOCX floating drawing: which wp:wrap element the anchor carries.
/// Closed set — the reader parses the anchor wrap element at the boundary, so layout
/// and snapshots compare enum members instead of raw element names.
/// </summary>
internal enum DocxFloatingWrapKind
{
    None,
    Square,
    Tight,
    Through,
    TopAndBottom
}

internal static class DocxFloatingWrapKindExtensions
{
    /// <summary>
    /// Stable spelling of a floating wrap kind. Kept as the historical OOXML element
    /// names so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxFloatingWrapKind wrapKind)
    {
        return wrapKind switch
        {
            DocxFloatingWrapKind.None => "wrapNone",
            DocxFloatingWrapKind.Square => "wrapSquare",
            DocxFloatingWrapKind.Tight => "wrapTight",
            DocxFloatingWrapKind.Through => "wrapThrough",
            DocxFloatingWrapKind.TopAndBottom => "wrapTopAndBottom",
            _ => throw new ArgumentOutOfRangeException(nameof(wrapKind), wrapKind, "Unknown floating wrap kind.")
        };
    }

    /// <summary>
    /// Parses an anchor wrap element local name (for example wrapSquare) into its kind.
    /// Accepts the bare kind spelling (for example square) for hand-built models.
    /// Returns null for absent or unrecognized wraps: unrecognized wraps never reach
    /// the model through the reader (unsupported anchors are filtered with a diagnostic),
    /// so downstream treats them as absent (no wrap exclusion).
    /// </summary>
    public static DocxFloatingWrapKind? FromLocalName(string? localName)
    {
        return localName?.ToLowerInvariant() switch
        {
            null => null,
            "wrapnone" or "none" => DocxFloatingWrapKind.None,
            "wrapsquare" or "square" => DocxFloatingWrapKind.Square,
            "wraptight" or "tight" => DocxFloatingWrapKind.Tight,
            "wrapthrough" or "through" => DocxFloatingWrapKind.Through,
            "wraptopandbottom" or "topandbottom" => DocxFloatingWrapKind.TopAndBottom,
            _ => null,
        };
    }
}
