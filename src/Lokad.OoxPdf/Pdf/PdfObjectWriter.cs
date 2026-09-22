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
        cancellationToken.ThrowIfCancellationRequested();
        offsets.Add(position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
        WriteAscii(FormattableString.Invariant($"<< /Length {contentBytes.Length} >>\nstream\n"));
        cancellationToken.ThrowIfCancellationRequested();
        stream.Write(contentBytes);
        position += contentBytes.Length;
        WriteAscii("endstream\nendobj\n");
    }

    public void WriteStreamObject(int objectNumber, string dictionaryEntries, ReadOnlySpan<byte> streamBytes)
    {
        cancellationToken.ThrowIfCancellationRequested();
        offsets.Add(position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
        WriteAscii(FormattableString.Invariant($"<< {dictionaryEntries} /Length {streamBytes.Length} >>\nstream\n"));
        cancellationToken.ThrowIfCancellationRequested();
        stream.Write(streamBytes);
        position += streamBytes.Length;
        WriteAscii("\nendstream\nendobj\n");
    }

    public void WriteAscii(string text)
    {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        stream.Write(bytes);
        position += bytes.Length;
    }
}