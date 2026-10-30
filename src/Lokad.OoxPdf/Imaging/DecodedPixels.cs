namespace Lokad.OoxPdf.Imaging;

// Local ownership of decoded planes. Callers may borrow them for transformation
// and compression. Capacity follows this existing owner until all borrows end.
internal sealed class DecodedPixels(int width, int height, byte[] rgb, byte[]? alpha,
    IMemoryReservation? storage = null) : IDisposable
{
    private byte[]? rgb = rgb;
    private byte[]? alpha = alpha;
    private IMemoryReservation? storage = storage;
    public int Width { get; } = width;
    public int Height { get; } = height;
    public byte[] Rgb => rgb ?? throw new ObjectDisposedException(nameof(DecodedPixels));
    public byte[]? Alpha
    {
        get { ObjectDisposedException.ThrowIf(rgb is null, this); return alpha; }
    }
    public void Dispose()
    {
        rgb = null;
        alpha = null;
        storage?.Dispose();
        storage = null;
    }
}
