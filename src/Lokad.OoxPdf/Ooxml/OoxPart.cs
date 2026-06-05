namespace Lokad.OoxPdf.Ooxml;

internal sealed record OoxPart(string Name, string ContentType, byte[] Bytes)
{
    public Stream OpenRead()
    {
        return new MemoryStream(Bytes, writable: false);
    }
}
