using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfObjectWriter
{
    private readonly Stream stream;
    private readonly CancellationToken cancellationToken;
    private readonly List<long> offsets = [];
    private readonly byte[] scratch;
    private long position;

    // The staged document owns/admitted this buffer. Every write completes before
    // formatting or readback reuses it; the writer never owns the caller's stream.
    public PdfObjectWriter(Stream stream, CancellationToken cancellationToken, byte[] scratch)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(scratch);
        if (!stream.CanWrite) throw new ArgumentException("Output stream must be writable.", nameof(stream));
        if (scratch.Length < 4096) throw new ArgumentException("Formatting scratch must contain at least 4096 bytes.", nameof(scratch));
        this.stream = stream;
        this.cancellationToken = cancellationToken;
        this.scratch = scratch;
    }

    public IReadOnlyList<long> Offsets => offsets;
    public long Position => position;
    internal Memory<byte> Scratch => scratch;

    public ValueTask WriteHeaderAsync() => WriteAsciiAsync("%PDF-1.7\n");

    public async ValueTask WriteObjectAsync(int objectNumber, string body)
    {
        await WriteObjectHeaderAsync(objectNumber).ConfigureAwait(false);
        await WriteAsciiAsync(body).ConfigureAwait(false);
        if (!body.EndsWith('\n')) await WriteAsciiAsync("\n").ConfigureAwait(false);
        await WriteObjectTrailerAsync().ConfigureAwait(false);
    }

    public async ValueTask WriteObjectHeaderAsync(int objectNumber)
    {
        cancellationToken.ThrowIfCancellationRequested();
        offsets.Add(position);
        await WriteAsciiAsync(FormattableString.Invariant($"{objectNumber} 0 obj\n")).ConfigureAwait(false);
    }

    public ValueTask WriteObjectTrailerAsync() => WriteAsciiAsync("endobj\n");

    public async ValueTask WriteContentStreamHeaderAsync(int objectNumber, int contentLength, string dictionaryEntries = "")
    {
        await WriteObjectHeaderAsync(objectNumber).ConfigureAwait(false);
        await WriteAsciiAsync(FormattableString.Invariant($"<<{(dictionaryEntries.Length == 0 ? string.Empty : " " + dictionaryEntries)} /Length {contentLength} >>\nstream\n")).ConfigureAwait(false);
    }

    public ValueTask WriteContentStreamTrailerAsync() => WriteAsciiAsync("endstream\nendobj\n");

    public async ValueTask WriteStreamObjectAsync(int objectNumber, string dictionaryEntries, ReadOnlyMemory<byte> bytes)
    {
        await WriteStreamObjectHeaderAsync(objectNumber, dictionaryEntries, bytes.Length).ConfigureAwait(false);
        await WriteRawBytesAsync(bytes).ConfigureAwait(false);
        await WriteStreamObjectTrailerAsync().ConfigureAwait(false);
    }

    public async ValueTask WriteStreamObjectHeaderAsync(int objectNumber, string dictionaryEntries, int length)
    {
        // Resource framing keeps its established trailing separator; staged page
        // content retains the distinct newline placement in its methods above.
        await WriteObjectHeaderAsync(objectNumber).ConfigureAwait(false);
        await WriteAsciiAsync(FormattableString.Invariant($"<< {dictionaryEntries} /Length {length} >>\nstream\n")).ConfigureAwait(false);
    }

    public ValueTask WriteStreamObjectTrailerAsync() => WriteAsciiAsync("\nendstream\nendobj\n");

    public async ValueTask WriteAsciiAsync(string text)
    {
        cancellationToken.ThrowIfCancellationRequested();
        for (int offset = 0; offset < text.Length;)
        {
            int count = Math.Min(scratch.Length, text.Length - offset);
            if (offset + count < text.Length && char.IsHighSurrogate(text[offset + count - 1]) && char.IsLowSurrogate(text[offset + count]))
                count--; // Keep ASCII's scalar fallback identical across chunk boundaries.
            int bytes = Encoding.ASCII.GetBytes(text, offset, count, scratch, 0);
            await WriteRawBytesAsync(scratch.AsMemory(0, bytes)).ConfigureAwait(false);
            offset += count;
        }
    }

    // All output crosses one admission boundary before the caller sees it.
    // A failed destination may retain a prefix; admitted bytes need not equal its
    // physical byte count. Cancellation never abandons an outstanding buffer borrow.
    public async ValueTask WriteRawBytesAsync(ReadOnlyMemory<byte> bytes)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxConversionBudget.Current?.ChargePdfOutputBytes(bytes.Length);
        await stream.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
        position += bytes.Length;
        cancellationToken.ThrowIfCancellationRequested();
    }
}
