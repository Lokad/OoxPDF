namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfImageXObject
{
    private PdfImageXObject(int width, int height, byte[] bytes, string filter, byte[]? alpha)
    {
        Width = width;
        Height = height;
        Bytes = bytes;
        Filter = filter;
        Alpha = alpha;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Bytes { get; }

    public string Filter { get; }

    public byte[]? Alpha { get; }

    public string ResourceKey => $"{Width}x{Height}:{Filter}:{Bytes.Length}:{Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(Bytes))[..16]}";

    public static PdfImageXObject Jpeg(int width, int height, byte[] bytes)
    {
        return new PdfImageXObject(width, height, bytes, "/DCTDecode", null);
    }

    public static PdfImageXObject RgbPng(int width, int height, byte[] rgb, byte[]? alpha)
    {
        return new PdfImageXObject(width, height, Compress(rgb), "/FlateDecode", alpha is null ? null : Compress(alpha));
    }

    private static byte[] Compress(byte[] bytes)
    {
        using var output = new MemoryStream();
        using (var zlib = new System.IO.Compression.ZLibStream(output, System.IO.Compression.CompressionLevel.SmallestSize, leaveOpen: true))
        {
            zlib.Write(bytes);
        }

        return output.ToArray();
    }
}
