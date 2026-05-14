namespace Lokad.OoxPdf.Diagnostics;

public sealed record OoxPdfDiagnostic(
    string Id,
    OoxPdfSeverity Severity,
    string Message,
    string? PartName = null,
    int? SlideIndex = null,
    int? PageIndex = null,
    string? Feature = null,
    string? Fallback = null);
