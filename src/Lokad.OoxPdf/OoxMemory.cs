using System.Xml.Linq;

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

    // Coarse family capacity from source shape, not an object ledger or heap
    // ceiling. Retained capacity includes ordinary placement/retry headroom;
    // a second window covers construction and refunds when the factory finishes.
    internal static IMemoryReservation? ReserveModel(XContainer source, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        IMemoryScope? owner = OoxConversionBudget.Current?.Models;
        if (owner is null) return null;
        IMemoryReservation? storage = Reserve(owner, 0);
        try { GrowModel(storage, source, token); return storage; }
        catch { storage?.Dispose(); throw; }
    }

    internal static void GrowModel(IMemoryReservation? storage, XContainer source, CancellationToken token)
    {
        if (storage is null) return;
        Resize(storage, checked(storage.Bytes + 2 * EstimateMarkupCapacity(source, token)));
        token.ThrowIfCancellationRequested();
    }

    // Shared source-shape summary; page painting adds its own expansion headroom.
    internal static long EstimateMarkupCapacity(XContainer source, CancellationToken token)
    {
        long capacity = 64L * 1024;
        foreach (XNode node in source.DescendantNodes())
        {
            token.ThrowIfCancellationRequested();
            capacity = checked(capacity + (node is XElement ? 1024L : node is XText text ? 64L * text.Value.Length : 0));
        }
        return capacity;
    }

    private static OoxPdfLimitExceededException Refused(OoxPdfLimitScope scope) => new(
        "Host refused shared-memory capacity for conversion storage.",
        OoxPdfLimitKind.MemoryAdmission, scope);
}
