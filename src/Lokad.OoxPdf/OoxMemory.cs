namespace Lokad.OoxPdf;

// Admission only: no private pool, heap ledger, retry or recovery policy.
internal static class OoxMemory
{
    public static IMemoryReservation? Reserve(IMemoryBudget? budget, long bytes, OoxPdfLimitScope scope = OoxPdfLimitScope.Conversion)
    {
        if (budget is null) return null;
        if (!budget.TryReserve(bytes, out var reservation)) throw Refused(scope);
        return reservation ?? throw new InvalidOperationException("Memory admission succeeded without a reservation.");
    }

    public static void Resize(IMemoryReservation? reservation, long bytes, OoxPdfLimitScope scope = OoxPdfLimitScope.Conversion)
    {
        if (reservation is not null && !reservation.TryResize(bytes)) throw Refused(scope);
    }

    private static OoxPdfLimitExceededException Refused(OoxPdfLimitScope scope) => new(
        "Host refused shared-memory capacity for conversion storage.",
        OoxPdfLimitKind.MemoryAdmission, scope);
}
