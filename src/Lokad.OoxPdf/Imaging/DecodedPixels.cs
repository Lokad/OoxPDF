namespace Lokad.OoxPdf.Imaging;

// R02: explicit decoded-pixel ownership. Pixel planes stay reserved from header
// validation through the final transform/compress step that consumes them; the
// reservation releases only when the owner disposes. Callers hold the instance
// across crop/recolor/compress work so the peak reflects the true simultaneous
// working set instead of dropping to zero at decoder return.
internal sealed class DecodedPixels : IDisposable
{
    private OoxConversionBudget.LiveReservation? reservation;
    private bool disposed;

    internal DecodedPixels(int width, int height, byte[] rgb, byte[]? alpha, OoxConversionBudget.LiveReservation? reservation)
    {
        Width = width;
        Height = height;
        Rgb = rgb;
        Alpha = alpha;
        this.reservation = reservation;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Rgb { get; }

    public byte[]? Alpha { get; }

    public void Dispose()
    {
        if (!disposed)
        {
            disposed = true;
            reservation?.Dispose();
            reservation = null;
        }
    }
}
