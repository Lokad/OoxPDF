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
