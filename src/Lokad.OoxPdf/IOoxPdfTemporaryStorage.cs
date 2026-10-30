namespace Lokad.OoxPdf;

/// <summary>Creates host-selected temporary backing when staged PDF content exceeds its resident window.</summary>
/// <remarks>
/// Each call returns a fresh, empty, readable/writable/seekable stream positioned at zero.
/// The provider remains caller-owned. The converter owns each returned stream and
/// awaits its disposal. Implementations must support
/// asynchronous reads/writes and concurrent calls for independent conversions.
/// Fund significant retained storage and construction buffers from the supplied budget
/// before allocation, including replacement overlap, and refund after their last user.
/// The budget is borrowed: never dispose it or keep it after the returned stream closes.
/// Null means the caller did not configure shared memory admission.
/// </remarks>
public interface IOoxPdfTemporaryStorage
{
    ValueTask<Stream> CreateAsync(IMemoryBudget? memoryBudget, CancellationToken cancellationToken = default);
}
