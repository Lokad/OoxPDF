namespace Lokad.OoxPdf;

/// <summary>Caller-owned admission to a shared pool of memory capacity.</summary>
/// <remarks>
/// Implementations support concurrent callers. Admission never waits. Scopes and
/// reservations share this pool; creating a scope reserves nothing and introduces
/// no private allowance. This contract admits significant storage, not the exact
/// managed heap. Callers supply headroom for incidental storage and collection lag.
/// </remarks>
public interface IMemoryBudget
{
    /// <summary>Creates an independently owned group sharing this budget's pool.</summary>
    IMemoryScope CreateScope();

    /// <summary>
    /// Reserves capacity before acquisition. On success return a non-null owned
    /// reservation; on refusal return false and null, without changing capacity.
    /// A zero-byte reservation is permitted. Negative sizes are invalid.
    /// </summary>
    bool TryReserve(long bytes, out IMemoryReservation? reservation);
}

/// <summary>An owned group of reservations against the same shared pool.</summary>
/// <remarks>
/// Dispose releases this group's outstanding reservations exactly once and must
/// not throw. It does not dispose the caller's budget or independently created
/// scopes. The owner must finish all storage users before disposal. Reservations
/// disposed after their scope must not refund capacity a second time.
/// </remarks>
public interface IMemoryScope : IMemoryBudget, IDisposable { }

/// <summary>Owned capacity held until storage users finish.</summary>
/// <remarks>Dispose refunds exactly once after those users finish and must not throw.</remarks>
public interface IMemoryReservation : IDisposable
{
    /// <summary>The currently reserved capacity in bytes.</summary>
    long Bytes { get; }

    /// <summary>
    /// Changes capacity without waiting. Refusal preserves the previous amount;
    /// shrinking must succeed. Owners admit old/new construction overlap before
    /// replacing storage, then shrink after the old storage's users finish.
    /// </summary>
    bool TryResize(long bytes);
}
