namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfImageXObject
{
    private string? resourceKey;

    private PdfImageXObject(int width, int height, byte[] bytes, string filter, string colorSpace, int bitsPerComponent, byte[]? alpha)
    {
        Width = width;
        Height = height;
        Bytes = bytes;
        Filter = filter;
        ColorSpace = colorSpace;
        BitsPerComponent = bitsPerComponent;
        Alpha = alpha;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Bytes { get; }

    public string Filter { get; }

    public string ColorSpace { get; }

    public int BitsPerComponent { get; }

    public byte[]? Alpha { get; }

    public string ResourceKey => resourceKey ??= $"{Width}x{Height}:{Filter}:{ColorSpace}:{BitsPerComponent}:{Bytes.Length}:{HashPrefix(Bytes)}:{Alpha?.Length ?? 0}:{(Alpha is null ? "none" : HashPrefix(Alpha))}";

    public static PdfImageXObject Jpeg(int width, int height, byte[] bytes, int componentCount = 3, int bitsPerComponent = 8)
    {
        string colorSpace = componentCount switch
        {
            1 => "/DeviceGray",
            4 => "/DeviceCMYK",
            _ => "/DeviceRGB"
        };
        return new PdfImageXObject(width, height, bytes, "/DCTDecode", colorSpace, bitsPerComponent, null);
    }

    public static PdfImageXObject RgbPng(int width, int height, byte[] rgb, byte[]? alpha)
    {
        return new PdfImageXObject(width, height, Compress(rgb), "/FlateDecode", "/DeviceRGB", 8, alpha is null ? null : Compress(alpha));
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

    private static string HashPrefix(byte[] bytes)
    {
        return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes))[..16];
    }
}
