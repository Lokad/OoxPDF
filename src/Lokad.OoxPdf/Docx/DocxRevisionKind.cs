namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Closed kind of a DOCX tracked revision: which OOXML revision element produced it.
/// Closed set — the tracked-change descendants (ins/del/moveFrom/moveTo), the six
/// property-change elements, and the four move-range markers; the reader assigns
/// these at the boundary from element names, so counters, predicates, formatting,
/// and balloon labels compare enum members instead of raw strings. Snapshots stay
/// string-mapped via ToValueString.
/// </summary>
internal enum DocxRevisionKind
{
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
    RunPropertiesChange,
    ParagraphPropertiesChange,
    TablePropertiesChange,
    TableRowPropertiesChange,
    TableCellPropertiesChange,
    SectionPropertiesChange,
    MoveFromRangeStart,
    MoveFromRangeEnd,
    MoveToRangeStart,
    MoveToRangeEnd
}

internal static class DocxRevisionKindExtensions
{
    /// <summary>
    /// Stable spelling of a revision kind. Kept as the historical PascalCase
    /// strings so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxRevisionKind kind)
    {
        return kind switch
        {
            DocxRevisionKind.Insertion => "Insertion",
            DocxRevisionKind.Deletion => "Deletion",
            DocxRevisionKind.MoveFrom => "MoveFrom",
            DocxRevisionKind.MoveTo => "MoveTo",
            DocxRevisionKind.RunPropertiesChange => "RunPropertiesChange",
            DocxRevisionKind.ParagraphPropertiesChange => "ParagraphPropertiesChange",
            DocxRevisionKind.TablePropertiesChange => "TablePropertiesChange",
            DocxRevisionKind.TableRowPropertiesChange => "TableRowPropertiesChange",
            DocxRevisionKind.TableCellPropertiesChange => "TableCellPropertiesChange",
            DocxRevisionKind.SectionPropertiesChange => "SectionPropertiesChange",
            DocxRevisionKind.MoveFromRangeStart => "MoveFromRangeStart",
            DocxRevisionKind.MoveFromRangeEnd => "MoveFromRangeEnd",
            DocxRevisionKind.MoveToRangeStart => "MoveToRangeStart",
            DocxRevisionKind.MoveToRangeEnd => "MoveToRangeEnd",
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown revision kind.")
        };
    }
}