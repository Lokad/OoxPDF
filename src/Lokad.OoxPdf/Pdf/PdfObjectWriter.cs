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
        WriteObjectHeader(objectNumber);
        WriteAscii(body);
        if (!body.EndsWith('\n'))
        {
            WriteAscii("\n");
        }

        WriteObjectTrailer();
    }

    public void WriteObjectHeader(int objectNumber)
    {
        cancellationToken.ThrowIfCancellationRequested();
        offsets.Add(position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
    }

    public void WriteObjectTrailer()
    {
        WriteAscii("endobj\n");
    }

    // RV11-P1: chunked content-stream emission. The header carries the known
    // total length; callers stream the payload in bounded chunks, so spilled
    // pages are never fully resident during emission.
    public void WriteContentStreamHeader(int objectNumber, int contentLength, string dictionaryEntries = "")
    {
        WriteObjectHeader(objectNumber);
        WriteAscii(FormattableString.Invariant($"<<{(dictionaryEntries.Length == 0 ? string.Empty : " " + dictionaryEntries)} /Length {contentLength} >>\nstream\n"));
    }

    public void WriteContentStreamBytes(ReadOnlySpan<byte> chunk)
    {
        WriteRawBytes(chunk);
    }

    public void WriteContentStreamTrailer()
    {
        WriteAscii("endstream\nendobj\n");
    }

    public void WriteStreamObject(int objectNumber, string dictionaryEntries, ReadOnlySpan<byte> streamBytes)
    {
        WriteStreamObjectHeader(objectNumber, dictionaryEntries, streamBytes.Length);
        WriteContentStreamBytes(streamBytes);
        WriteStreamObjectTrailer();
    }

    public void WriteStreamObjectHeader(int objectNumber, string dictionaryEntries, int length)
    {
        // Buffered and generated resources share their established framing;
        // page-content staging retains its own newline placement above.
        WriteObjectHeader(objectNumber);
        WriteAscii(FormattableString.Invariant($"<< {dictionaryEntries} /Length {length} >>\nstream\n"));
    }

    public void WriteStreamObjectTrailer()
    {
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
    public void WriteRawBytes(ReadOnlySpan<byte> bytes)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxConversionBudget.Current?.ChargePdfOutputBytes(bytes.Length);
        stream.Write(bytes);
        position += bytes.Length;
    }
}
