namespace Lokad.OoxPdf.Fonts;

/// <summary>Borrows immutable programs held and funded by the caller.</summary>
/// <remarks>Keep the supplied storage and its funding alive for this source and its leases.</remarks>
public sealed class MemoryFontProgramSource : IFontProgramSource
{
    private readonly ReadOnlyMemory<byte> bytes;

    public MemoryFontProgramSource(string stableId, ReadOnlyMemory<byte> bytes)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(stableId);
        StableId = stableId;
        this.bytes = bytes;
    }

    public string StableId { get; }

    public ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();
        using var program = new FontProgramBuffer(bytes, null);
        return ValueTask.FromResult(program.Acquire());
    }
}
