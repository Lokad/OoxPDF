namespace Lokad.OoxPdf.Imaging;

internal static class ImagePixelBudget
{
    // The per-image cap bounds pixel planes independently of conversion-wide
    // admission limits. PNG keeps exact compressed IDAT storage and two scanlines
    // beside RGB/alpha planes; other decoders have their own scratch estimates.
    // Conversion-wide reservations and retained-resource caps bound covered
    // aggregate work; this check remains the early guard against lying headers.
    // Effect rasters are emitted once per shadow shape (never cached) and retained
    // only as their compressed PDF form.
    // A lying header (four bytes) must never turn into a giant allocation:
    // validate checked dimensions before any pixel buffer is sized.
    internal const int MaxDimension = 32768;
    internal const long MaxPixels = 67_108_864L;

    public static void Check(int width, int height, string format)
    {
        if (width <= 0 || height <= 0 || width > MaxDimension || height > MaxDimension)
        {
            throw new InvalidDataException($"{format} image dimensions are out of range: {width}x{height}.");
        }

        if ((long)width * height > MaxPixels)
        {
            throw new InvalidDataException($"{format} image exceeds the maximum supported pixel count of {MaxPixels}.");
        }
    }
}

