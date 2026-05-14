using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfObjectWriter
{
    private readonly Stream stream;
    private readonly List<long> offsets = [];

    public PdfObjectWriter(Stream stream)
    {
        this.stream = stream;
    }

    public IReadOnlyList<long> Offsets => offsets;

    public long Position => stream.Position;

    public void WriteHeader()
    {
        WriteAscii("%PDF-1.7\n");
    }

    public void WriteObject(int objectNumber, string body)
    {
        offsets.Add(stream.Position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
        WriteAscii(body);
        if (!body.EndsWith('\n'))
        {
            WriteAscii("\n");
        }

        WriteAscii("endobj\n");
    }

    public void WriteStreamObject(int objectNumber, string dictionaryEntries, ReadOnlySpan<byte> streamBytes)
    {
        offsets.Add(stream.Position);
        WriteAscii(FormattableString.Invariant($"{objectNumber} 0 obj\n"));
        WriteAscii(FormattableString.Invariant($"<< {dictionaryEntries} /Length {streamBytes.Length} >>\nstream\n"));
        stream.Write(streamBytes);
        WriteAscii("\nendstream\nendobj\n");
    }

    public void WriteAscii(string text)
    {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        stream.Write(bytes);
    }
}
