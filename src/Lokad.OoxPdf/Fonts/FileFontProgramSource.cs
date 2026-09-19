namespace Lokad.OoxPdf.Fonts;

public sealed class FileFontProgramSource : IFontProgramSource
{
    private readonly string path;
    private readonly object sync = new();
    private ReadOnlyMemory<byte>? cachedBytes;

    public FileFontProgramSource(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        this.path = System.IO.Path.GetFullPath(path);
    }

    public string Path => path;

    public string StableId => "file:" + path;

    public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default)
    {
        if (cachedBytes is ReadOnlyMemory<byte> bytes)
        {
            return ValueTask.FromResult(bytes);
        }

        ct.ThrowIfCancellationRequested();
        lock (sync)
        {
            if (cachedBytes is ReadOnlyMemory<byte> cached)
            {
                return ValueTask.FromResult(cached);
            }

            ct.ThrowIfCancellationRequested();
            byte[] loaded = File.ReadAllBytes(path);
            cachedBytes = loaded;
            return ValueTask.FromResult((ReadOnlyMemory<byte>)loaded);
        }
    }
}
