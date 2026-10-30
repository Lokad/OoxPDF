namespace Lokad.OoxPdf.Fonts;

/// <summary>Acquires immutable font programs behind a resolver-scoped stable identity.</summary>
/// <remarks>
/// Implementations support concurrent callers and return byte-identical content
/// for a StableId over their lifetime. The identity is opaque and case-sensitive;
/// never normalize it as a path or compare it across resolver kinds.
/// Admit significant acquisition storage and construction overlap before allocation.
/// The operation budget is borrowed: never dispose it or retain it beyond the
/// acquisition/lease. Shared caches use a separate, longer-lived caller budget.
/// Caller-owned, caller-funded immutable memory may be borrowed without copying
/// or admitting it again. Cancellation must finish outstanding I/O users before
/// releasing their buffers or admission.
/// </remarks>
public interface IFontProgramSource
{
    string StableId { get; }
    ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default);
}
