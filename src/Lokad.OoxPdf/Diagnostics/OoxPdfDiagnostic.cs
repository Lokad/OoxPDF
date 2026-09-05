namespace Lokad.OoxPdf.Diagnostics;

public sealed record OoxPdfDiagnostic(
    string Id,
    OoxPdfSeverity Severity,
    string Message,
    string? PartName,
    int? SlideIndex,
    int? PageIndex,
    string? Feature,
    string? Fallback);
