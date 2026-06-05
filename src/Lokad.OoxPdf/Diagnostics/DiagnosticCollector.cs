namespace Lokad.OoxPdf.Diagnostics;

public sealed class DiagnosticCollector
{
    private readonly List<OoxPdfDiagnostic> diagnostics = [];

    public IReadOnlyList<OoxPdfDiagnostic> Diagnostics => diagnostics;

    public bool HasWarningsOrErrors => diagnostics.Any(d => d.Severity is OoxPdfSeverity.Warning or OoxPdfSeverity.Error);

    public void Add(OoxPdfDiagnostic diagnostic)
    {
        diagnostics.Add(diagnostic);
    }
}
