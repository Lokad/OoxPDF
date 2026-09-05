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

/// <summary>
/// Closed w:type of a footnote/endnote story element (separator and continuation
/// variants versus normal note bodies). The reader assigns this at the boundary from
/// the raw attribute; a missing attribute stays null (absent), which layout treats
/// like Normal. Unknown spellings map to Unknown rather than throwing, so the
/// reader stays total over invalid documents.
/// </summary>
internal enum DocxRelatedStoryType
{
    Normal,
    Separator,
    ContinuationSeparator,
    ContinuationNotice,
    Unknown
}

internal static class DocxRelatedStoryTypeExtensions
{
    /// <summary>
    /// Historical w:type spelling. Normal maps back to normal so inspection snapshots
    /// stay byte-comparable; a missing attribute is represented by null, never by this method.
    /// </summary>
    public static string ToValueString(this DocxRelatedStoryType type)
    {
        return type switch
        {
            DocxRelatedStoryType.Normal => "normal",
            DocxRelatedStoryType.Separator => "separator",
            DocxRelatedStoryType.ContinuationSeparator => "continuationSeparator",
            DocxRelatedStoryType.ContinuationNotice => "continuationNotice",
            _ => "unknown",
        };
    }
}
