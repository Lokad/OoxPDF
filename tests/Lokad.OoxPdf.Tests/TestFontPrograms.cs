using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class TestFontPrograms
{
    // Inspectors that need bytes after acquisition keep an independent test copy.
    internal static async ValueTask<ReadOnlyMemory<byte>> CopyAsync(IFontProgramSource source, CancellationToken token = default)
    {
        using var lease = await source.AcquireAsync(null, token).ConfigureAwait(false);
        return lease.Bytes.ToArray();
    }

    internal static IOoxPdfFontProgramLease Borrow(ReadOnlyMemory<byte> bytes)
    {
        using var buffer = new FontProgramBuffer(bytes, null);
        return buffer.Acquire();
    }
}
