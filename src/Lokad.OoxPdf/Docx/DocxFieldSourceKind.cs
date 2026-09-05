namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Provenance of a DOCX field reference: which field syntax produced it.
/// Closed set — simple fields carry an instr attribute while complex fields use
/// separate instruction/result runs, so downstream counters compare enum members
/// instead of raw strings.
/// </summary>
internal enum DocxFieldSourceKind
{
    Simple,
    ComplexInstruction
}

internal static class DocxFieldSourceKindExtensions
{
    /// <summary>
    /// Stable spelling of a field source kind. Kept as the historical PascalCase
    /// strings so inspection snapshots stay comparable.
    /// </summary>
    public static string ToValueString(this DocxFieldSourceKind sourceKind)
    {
        return sourceKind switch
        {
            DocxFieldSourceKind.Simple => "Simple",
            DocxFieldSourceKind.ComplexInstruction => "ComplexInstruction",
            _ => throw new ArgumentOutOfRangeException(nameof(sourceKind), sourceKind, "Unknown field source kind.")
        };
    }
}
