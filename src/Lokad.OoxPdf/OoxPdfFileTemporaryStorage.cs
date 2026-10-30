namespace Lokad.OoxPdf;

/// <summary>Creates owned, asynchronous temporary files in the supplied directory.</summary>
/// <remarks>Files use exclusive access and delete-on-close. The directory path is resolved at construction.</remarks>
public sealed class OoxPdfFileTemporaryStorage : IOoxPdfTemporaryStorage
{
    private readonly string directory;

    public OoxPdfFileTemporaryStorage(string directory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);
        this.directory = Path.GetFullPath(directory);
    }

    public ValueTask<Stream> CreateAsync(IMemoryBudget? memoryBudget, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "OoxPdfPages-" + Guid.NewGuid().ToString("N") + ".tmp");
        // The staging owner supplies its admitted encoding/readback buffer. Disable
        // FileStream's additional managed buffer rather than account for a duplicate.
        Stream stream = new FileStream(path, new FileStreamOptions {
            Mode = FileMode.CreateNew, Access = FileAccess.ReadWrite, Share = FileShare.None,
            BufferSize = 1, Options = FileOptions.Asynchronous | FileOptions.RandomAccess | FileOptions.DeleteOnClose });
        return ValueTask.FromResult(stream);
    }
}
