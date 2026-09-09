namespace Lokad.OoxPdf.Imaging;

internal static class ImagePixelBudget
{
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

