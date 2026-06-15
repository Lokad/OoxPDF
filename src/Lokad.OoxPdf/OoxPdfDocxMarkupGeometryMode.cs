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
    /// Selects the Office-compatible all-markup print layout profile. Until the source package exposes enough
    /// Word print-view data, this profile falls back to the reserved markup-margin geometry.
    /// </summary>
    WordCompatibleAllMarkup = 2
}
