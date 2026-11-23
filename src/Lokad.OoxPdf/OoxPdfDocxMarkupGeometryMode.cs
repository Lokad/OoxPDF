namespace Lokad.OoxPdf;

/// <summary>
/// Selects how DOCX review markup affects page layout geometry.
/// </summary>
public enum OoxPdfDocxMarkupGeometryMode
{
    /// <summary>
    /// Keeps the authored page media box and body text frame. This is the default compatibility mode.
    /// </summary>
    PreserveDocumentLayout = 0,

    /// <summary>
    /// Keeps the authored page media box but reserves a deterministic right-side review margin for all-markup balloons.
    /// </summary>
    ReserveMarkupMargin = 1,

    /// <summary>
    /// Uses an approximate Word print-review profile when supported balloons print in AllMarkup.
    /// The profile scales the body into the review frame. Otherwise, the authored body layout is preserved.
    /// </summary>
    WordCompatibleAllMarkup = 2
}
