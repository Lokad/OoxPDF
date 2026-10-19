namespace Lokad.OoxPdf;

/// <summary>
/// Host-supplied admission to shared memory capacity. Called once before a
/// conversion constructs retained storage or rendering intermediates.
/// </summary>
/// <remarks>
/// The host chooses a coarse reservation for the expected retained storage and
/// construction peak of this input, using its own shared pool. This is admission,
/// not measurement of the managed heap. Include headroom for the document model,
/// font preparation, decoded images and PDF resources; archive/image bounds still
/// apply independently. Caller-owned streams/resolvers and their bounded caches
/// remain the caller's responsibility. Existing bounded shared font caches are
/// not attributed to individual conversions; allow for that baseline in pool capacity.
/// Implementations must support concurrent conversions and nested calls. A null
/// result refuses the conversion. A successful lease must release its reservation
/// exactly once when disposed; Dispose must not throw. The library disposes it on
/// success, failure or cancellation after staged document ownership is released.
/// </remarks>
public interface IOoxPdfMemoryAdmission
{
    /// <summary>
    /// Attempts to admit one conversion without waiting for capacity. Observe the
    /// supplied cancellation token; return null when capacity is unavailable.
    /// </summary>
    IDisposable? TryAcquire(CancellationToken cancellationToken);
}
