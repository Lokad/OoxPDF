namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed kind of a DOCX related story (footnotes, endnotes, comment threads) and of
/// the inline references that resolve against those stories by (Kind, Id).
/// Closed set — the reader assigns these at the boundary from part types and element
/// names (plus TextBox for layout-synthesized textbox stories), so resolution and
/// counters compare enum members instead of raw strings.
/// </summary>
internal enum DocxRelatedStoryKind
{
    Comment,
    Footnote,
    Endnote,
    TextBox
}

internal static class DocxRelatedStoryKindExtensions
{
    /// <summary>
    /// Stable spelling of a related-story kind. Kept as the historical PascalCase
    /// strings so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxRelatedStoryKind kind)
    {
        return kind switch
        {
            DocxRelatedStoryKind.Comment => "Comment",
            DocxRelatedStoryKind.Footnote => "Footnote",
            DocxRelatedStoryKind.Endnote => "Endnote",
            DocxRelatedStoryKind.TextBox => "TextBox",
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown related-story kind.")
        };
    }
}
