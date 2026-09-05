namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed kind of a DOCX markup balloon candidate (and of the placement built from it).
/// Closed set — renderer-local: candidates are constructed at five sites and merged
/// by kind, so comparisons use enum members instead of raw strings.
/// Declaration order is the band ORDER (see OrderMarkupBalloonCandidateBand): Comment,
/// Markup, Revision, Overflow. Do not reorder without updating band layout.
/// </summary>
internal enum DocxMarkupBalloonKind
{
    Comment,
    Markup,
    Revision,
    Overflow
}

internal static class DocxMarkupBalloonKindExtensions
{
    /// <summary>
    /// Stable spelling of a markup balloon kind. Kept as the historical PascalCase
    /// strings so placement snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxMarkupBalloonKind kind)
    {
        return kind switch
        {
            DocxMarkupBalloonKind.Comment => "Comment",
            DocxMarkupBalloonKind.Markup => "Markup",
            DocxMarkupBalloonKind.Revision => "Revision",
            DocxMarkupBalloonKind.Overflow => "Overflow",
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown markup balloon kind.")
        };
    }
}
