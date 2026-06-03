namespace Lokad.OoxPdf.Fonts;

public interface IFontProgramSource
{
    string StableId { get; }

    ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default);
}
