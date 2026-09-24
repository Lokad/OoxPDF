using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfObjectWriter
{
    private readonly Stream stream;
    private readonly CancellationToken cancellationToken;
    private readonly List<long> offsets = [];
    private long position;

    public PdfObjectWriter(Stream stream, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanWrite)
        {
            throw new ArgumentException("Output stream must be writable.", nameof(stream));
        }

        this.stream = stream;
        this.cancellationToken = cancellationToken;
    }

    public IReadOnlyList<long> Offsets => offsets;

    public long Position => position;

    public void WriteHeader()
    {
        cancellationToken.ThrowIfCancellationRequested();
        WriteAscii("%PDF-1.7\n");
    }

    public void WriteObject(int objectNumber, string body)
    {
        cancellationToken.ThrowIfCancellationRequested();
        offsets.Add(position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
        WriteAscii(body);
        if (!body.EndsWith('\n'))
        {
            WriteAscii("\n");
        }

        WriteAscii("endobj\n");
    }

    // page content streams are written from a single ASCII encoding of the
    // content. The byte count comes from the encoded span, so no second whole-page
    // string interpolation plus re-encoding is needed merely to measure the length.
    // Output bytes match the previous WriteObject encoding exactly.
    public void WriteContentStreamObject(int objectNumber, ReadOnlySpan<byte> contentBytes)
    {
        WriteContentStreamHeader(objectNumber, contentBytes.Length);
        WriteContentStreamBytes(contentBytes);
        WriteContentStreamTrailer();
    }

    // RV11-P1: chunked content-stream emission. The header carries the known
    // total length; callers stream the payload in bounded chunks, so spilled
    // pages are never fully resident during emission.
    public void WriteContentStreamHeader(int objectNumber, int contentLength)
    {
        cancellationToken.ThrowIfCancellationRequested();
        offsets.Add(position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
        WriteAscii(FormattableString.Invariant($"<< /Length {contentLength} >>\nstream\n"));
    }

    public void WriteContentStreamBytes(ReadOnlySpan<byte> chunk)
    {
        cancellationToken.ThrowIfCancellationRequested();
        WriteRawBytes(chunk);
    }

    public void WriteContentStreamTrailer()
    {
        WriteAscii("endstream\nendobj\n");
    }

    public void WriteStreamObject(int objectNumber, string dictionaryEntries, ReadOnlySpan<byte> streamBytes)
    {
        cancellationToken.ThrowIfCancellationRequested();
        offsets.Add(position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
        WriteAscii(FormattableString.Invariant($"<< {dictionaryEntries} /Length {streamBytes.Length} >>\nstream\n"));
        cancellationToken.ThrowIfCancellationRequested();
        WriteRawBytes(streamBytes);
        WriteAscii("\nendstream\nendobj\n");
    }

    public void WriteAscii(string text)
    {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        WriteRawBytes(bytes);
    }

    // R06.1: every PDF byte crosses this single boundary. The output budget is
    // admitted before the destination sees the chunk, so a zero budget writes
    // nothing and an exact budget trips only when a further byte is attempted.
    // On success the admitted total equals Position. When the destination throws
    // mid-write, the budget holds admitted bytes, which may exceed the prefix the
    // destination actually kept; no exact recoverable count is invented. No
    // output-sized buffer is introduced here: ASCII payloads are encoded once by
    // WriteAscii and raw spans charge their length directly.
    private void WriteRawBytes(ReadOnlySpan<byte> bytes)
    {
        OoxConversionBudget.Current?.ChargePdfOutputBytes(bytes.Length);
        stream.Write(bytes);
        position += bytes.Length;
    }
}