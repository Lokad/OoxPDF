namespace Lokad.OoxPdf.Fonts;

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

    public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();
        return ValueTask.FromResult(bytes);
    }
}
