namespace Lokad.OoxPdf.Imaging;

internal static class ImagePixelBudget
{
    // PLAN M08 accounting status: the per-image pixel cap below bounds every single
    // image (single-image live set: compressed bytes, one inflated copy, RGB/alpha
    // planes, compression scratch; duplicate IDAT/inflated copies were removed).
    // What it does NOT bound is the aggregate: many individually legal images,
    // effect rasters (each capped per side), and retained recolor/crop variants can
    // still sum without a shared ceiling. That cumulative budget belongs to the
    // conversion-wide resource contract (Q01); per-image work here stays
    // document-proportional until then. Effect rasters are emitted once per shadow
    // shape (never cached) and retained only as their compressed PDF form.
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

