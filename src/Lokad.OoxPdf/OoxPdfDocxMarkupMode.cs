namespace Lokad.OoxPdf;

/// <summary>
/// Selects the WordprocessingML review view used when converting DOCX inputs.
/// </summary>
public enum OoxPdfDocxMarkupMode
{
    /// <summary>
    /// Renders the final document text without visible markup. This is the default compatibility mode.
    /// </summary>
    Final = 0,

    /// <summary>
    /// Renders the original document text by showing deleted and moved-from content while hiding inserted and moved-to content.
    /// </summary>
    Original = 1,

    /// <summary>
    /// Renders final text with lightweight change bars and comment markers.
    /// </summary>
    SimpleMarkup = 2,

    /// <summary>
    /// Renders visible tracked-change text with inline styling plus first-pass comment and revision balloons.
    /// </summary>
    AllMarkup = 3
}
