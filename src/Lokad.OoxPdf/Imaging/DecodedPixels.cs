namespace Lokad.OoxPdf.Imaging;

// Local ownership of decoded planes. Callers may borrow them for transformation
// and compression, then release this owner's references without a memory ledger.
internal sealed class DecodedPixels(int width, int height, byte[] rgb, byte[]? alpha) : IDisposable
{
    private byte[]? rgb = rgb;
    private byte[]? alpha = alpha;
    public int Width { get; } = width;
    public int Height { get; } = height;
    public byte[] Rgb => rgb ?? throw new ObjectDisposedException(nameof(DecodedPixels));
    public byte[]? Alpha
    {
        get { ObjectDisposedException.ThrowIf(rgb is null, this); return alpha; }
    }
    public void Dispose() { rgb = null; alpha = null; }
}
