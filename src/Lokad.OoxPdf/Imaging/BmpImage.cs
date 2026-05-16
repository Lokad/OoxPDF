namespace Lokad.OoxPdf.Imaging;

internal sealed class BmpImage
{
    private BmpImage(int width, int height, byte[] rgb, byte[]? alpha)
    {
        Width = width;
        Height = height;
        Rgb = rgb;
        Alpha = alpha;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Rgb { get; }

    public byte[]? Alpha { get; }

    public static BmpImage Read(byte[] bytes)
    {
        if (bytes.Length < 54 || bytes[0] != (byte)'B' || bytes[1] != (byte)'M')
        {
            throw new InvalidDataException("Data is not a BMP image.");
        }

        int pixelOffset = I32(bytes, 10);
        int dibSize = I32(bytes, 14);
        if (dibSize < 40 || pixelOffset < 14 + dibSize || pixelOffset >= bytes.Length)
        {
            throw new InvalidDataException("BMP header is invalid.");
        }

        int width = I32(bytes, 18);
        int signedHeight = I32(bytes, 22);
        ushort planes = U16(bytes, 26);
        ushort bitsPerPixel = U16(bytes, 28);
        int compression = I32(bytes, 30);
        if (width <= 0 || signedHeight == 0 || planes != 1 || compression != 0 || bitsPerPixel is not (24 or 32))
        {
            throw new NotSupportedException($"Unsupported BMP format: width={width}, height={signedHeight}, bitsPerPixel={bitsPerPixel}, compression={compression}.");
        }

        int height = Math.Abs(signedHeight);
        bool topDown = signedHeight < 0;
        int bytesPerPixel = bitsPerPixel / 8;
        int stride = ((width * bytesPerPixel + 3) / 4) * 4;
        if (pixelOffset + stride * height > bytes.Length)
        {
            throw new InvalidDataException("BMP pixel data is truncated.");
        }

        var rgb = new byte[width * height * 3];
        for (int y = 0; y < height; y++)
        {
            int sourceY = topDown ? y : height - 1 - y;
            int source = pixelOffset + sourceY * stride;
            int target = y * width * 3;
            for (int x = 0; x < width; x++)
            {
                rgb[target++] = bytes[source + 2];
                rgb[target++] = bytes[source + 1];
                rgb[target++] = bytes[source];
                source += bytesPerPixel;
            }
        }

        return new BmpImage(width, height, rgb, alpha: null);
    }

    private static ushort U16(byte[] bytes, int offset)
    {
        return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
    }

    private static int I32(byte[] bytes, int offset)
    {
        return bytes[offset] |
            (bytes[offset + 1] << 8) |
            (bytes[offset + 2] << 16) |
            (bytes[offset + 3] << 24);
    }
}
