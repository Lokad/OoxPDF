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

    // Q01: cooperative cancellation inside the pixel loop (see PngImage.Read).

    // Read keeps release-on-return lifetime for header probes, tests, and callers
    // without downstream transforms. Production image paths use ReadOwned and hold
    // the reservation across crop/recolor/compress work (R02).
    public static BmpImage Read(byte[] bytes, CancellationToken cancellationToken = default)
    {
        using DecodedPixels owned = ReadOwned(bytes, cancellationToken);
        return new BmpImage(owned.Width, owned.Height, owned.Rgb, owned.Alpha);
    }

    public static DecodedPixels ReadOwned(byte[] bytes, CancellationToken cancellationToken = default)
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

        if (signedHeight == int.MinValue)
        {
            throw new InvalidDataException("BMP header is invalid.");
        }

        int height = Math.Abs(signedHeight);
        ImagePixelBudget.Check(width, height, "BMP");
        bool topDown = signedHeight < 0;
        int bytesPerPixel = bitsPerPixel / 8;
        long stride = ((long)width * bytesPerPixel + 3L) / 4L * 4L;
        if ((long)pixelOffset + stride * height > bytes.Length)
        {
            throw new InvalidDataException("BMP pixel data is truncated.");
        }

        int stride32 = checked((int)stride);

        // R02: BMP decodes in place from caller-owned input, so the working set is
        // exactly the output plane. Ownership transfers to the caller (see PngImage).
        long liveEstimate = checked((long)width * height * 3L);
        OoxConversionBudget.LiveReservation? liveReservation = OoxConversionBudget.Current?.ReserveLiveImageBytes(liveEstimate);
        try
        {
            var rgb = new byte[width * height * 3];
            for (int y = 0; y < height; y++)
            {
                if ((y & 255) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }
                int sourceY = topDown ? y : height - 1 - y;
                int source = pixelOffset + sourceY * stride32;
                int target = y * width * 3;
                for (int x = 0; x < width; x++)
                {
                    rgb[target++] = bytes[source + 2];
                    rgb[target++] = bytes[source + 1];
                    rgb[target++] = bytes[source];
                    source += bytesPerPixel;
                }
            }
            return new DecodedPixels(width, height, rgb, alpha: null, liveReservation);
        }
        catch
        {
            liveReservation?.Dispose();
            throw;
        }
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
