namespace Lokad.OoxPdf.Diagnostics;

public sealed class DiagnosticCollector
{
    /// <summary>
    /// Maximum retained diagnostic samples per conversion (Q02). Occurrence totals
    /// are still counted past the cap; see <see cref="OccurrenceCounts"/> and
    /// <see cref="DroppedCount"/>.
    /// </summary>
    public const int MaxRetainedDiagnostics = 4096;

    private readonly List<OoxPdfDiagnostic> diagnostics = [];
    private readonly Dictionary<string, int> occurrenceCounts = new(StringComparer.Ordinal);
    private int droppedCount;

    public IReadOnlyList<OoxPdfDiagnostic> Diagnostics => diagnostics;

    public IReadOnlyDictionary<string, int> OccurrenceCounts => occurrenceCounts;

    public int DroppedCount => droppedCount;

    private bool sawWarningOrError;

    /// <summary>
    /// Whether any warning or error was ever observed (R20). Tracked independently of
    /// sampling: retained entries are capped, but severity must survive overflow so CLI
    /// strict mode cannot miss an error that was dropped from the retained samples.
    /// </summary>
    public bool HasWarningsOrErrors => sawWarningOrError;

    public void Add(OoxPdfDiagnostic diagnostic)
    {
        occurrenceCounts[diagnostic.Id] = occurrenceCounts.TryGetValue(diagnostic.Id, out int count) ? count + 1 : 1;
        if (diagnostic.Severity is OoxPdfSeverity.Warning or OoxPdfSeverity.Error)
        {
            sawWarningOrError = true;
        }
        if (diagnostics.Count < MaxRetainedDiagnostics)
        {
            diagnostics.Add(diagnostic);
            return;
        }

        droppedCount++;
    }

    /// <summary>
    /// Retained samples plus a terminal overflow marker when samples were dropped,
    /// so serialized diagnostics disclose the loss instead of truncating silently.
    /// </summary>
    public IReadOnlyList<OoxPdfDiagnostic> DiagnosticsWithOverflow()
    {
        if (droppedCount == 0)
        {
            return diagnostics;
        }

        return
        [
            .. diagnostics,
            new OoxPdfDiagnostic(
                "DIAGNOSTIC_OVERFLOW",
                OoxPdfSeverity.Error,
                $"Retained the first {MaxRetainedDiagnostics} diagnostics and dropped {droppedCount} further occurrences; per-ID totals are available on the collector.",
                PartName: null,
                SlideIndex: null,
                PageIndex: null,
                Feature: null,
                Fallback: "Truncated"),
        ];
    }
}
