namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Provenance of a block-level break element: which OOXML construct produced it.
/// Closed set — parser, layout, and snapshots share these members instead of
/// comparing raw strings, so a new break provenance fails loudly at parse time
/// rather than silently flowing downstream.
/// </summary>
internal enum DocxBreakSourceKind
{
    RunBreak,
    PageBreakBefore,
    TerminalTable
}

internal static class DocxBreakSourceKindExtensions
{
    /// <summary>
    /// Stable snapshot/diagnostic spelling of a break source kind. Kept as the
    /// historical lowercase strings so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxBreakSourceKind sourceKind)
    {
        return sourceKind switch
        {
            DocxBreakSourceKind.RunBreak => "runBreak",
            DocxBreakSourceKind.PageBreakBefore => "pageBreakBefore",
            DocxBreakSourceKind.TerminalTable => "terminalTable",
            _ => throw new ArgumentOutOfRangeException(nameof(sourceKind), sourceKind, "Unknown break source kind.")
        };
    }
}
