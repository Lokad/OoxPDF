namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed kind of a story that layout items can belong to (T03). The body,
/// header/footer, table-cell, and related-story cases are the only values layout
/// assigns; emission-level groupings ("Static", "RelatedStory") stay renderer-local
/// fallback strings, never layout values.
/// </summary>
internal enum DocxStoryKind
{
    Body,
    Header,
    Footer,
    TableCell,
    Comment,
    Footnote,
    Endnote,
    TextBox
}

/// <summary>
/// Typed story identity carried by layout items (T03). Replaces the stringly-typed
/// StoryKind/StoryVariantType pair so misspellings fail at compile time and switches
/// are exhaustive. The historical spellings are preserved exactly for inspection
/// snapshots via <see cref="ToKindString"/>; snapshot JSON stays byte-identical.
/// Variant stays a raw string: header/footer w:type spellings ("default", "first",
/// "even") double as document dictionary keys, so unknown spellings flow through
/// rather than throwing.
/// </summary>
internal readonly record struct DocxStoryId(DocxStoryKind Kind, string? VariantType)
{
    public static DocxStoryId Body() => new(DocxStoryKind.Body, null);

    public static DocxStoryId TableCell() => new(DocxStoryKind.TableCell, null);

    public static DocxStoryId Header(string? variantType) => new(DocxStoryKind.Header, variantType);

    public static DocxStoryId Footer(string? variantType) => new(DocxStoryKind.Footer, variantType);

    /// <summary>
    /// Header/footer selection used by static story layout: the boolean chooses the
    /// kind, the raw w:type spelling ("default", "first", "even") rides along.
    /// </summary>
    public static DocxStoryId HeaderOrFooter(bool isHeader, string? variantType) =>
        isHeader ? Header(variantType) : Footer(variantType);

    public static DocxStoryId Related(DocxRelatedStoryKind kind) => new(ToStoryKind(kind), null);

    public bool IsHeaderOrFooter => Kind is DocxStoryKind.Header or DocxStoryKind.Footer;

    /// <summary>
    /// Historical spelling used by inspection snapshots. Every kind maps to the
    /// exact string the layout strings previously carried.
    /// </summary>
    public string ToKindString()
    {
        return Kind.ToKindString();
    }

    private static DocxStoryKind ToStoryKind(DocxRelatedStoryKind kind)
    {
        return kind switch
        {
            DocxRelatedStoryKind.Comment => DocxStoryKind.Comment,
            DocxRelatedStoryKind.Footnote => DocxStoryKind.Footnote,
            DocxRelatedStoryKind.Endnote => DocxStoryKind.Endnote,
            DocxRelatedStoryKind.TextBox => DocxStoryKind.TextBox,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown related-story kind.")
        };
    }
}

internal static class DocxStoryKindExtensions
{
    /// <summary>
    /// Historical spelling used by inspection snapshots. Shared with
    /// <see cref="DocxStoryId"/> so structure producers (which only know the kind,
    /// not a full identity) render byte-identical strings.
    /// </summary>
    public static string ToKindString(this DocxStoryKind kind)
    {
        return kind switch
        {
            DocxStoryKind.Body => "Body",
            DocxStoryKind.Header => "Header",
            DocxStoryKind.Footer => "Footer",
            DocxStoryKind.TableCell => "TableCell",
            DocxStoryKind.Comment => "Comment",
            DocxStoryKind.Footnote => "Footnote",
            DocxStoryKind.Endnote => "Endnote",
            DocxStoryKind.TextBox => "TextBox",
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown story kind.")
        };
    }
}
