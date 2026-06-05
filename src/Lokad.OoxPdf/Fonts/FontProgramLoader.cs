namespace Lokad.OoxPdf.Fonts;

internal static class FontProgramLoader
{
    public static OpenTypeFont? Load(FontFaceResolution? resolution, CancellationToken cancellationToken = default)
    {
        if (resolution is null)
        {
            return null;
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            ReadOnlyMemory<byte> bytes = resolution.Source.GetBytesAsync(cancellationToken).AsTask().GetAwaiter().GetResult();
            cancellationToken.ThrowIfCancellationRequested();
            return OpenTypeFont.Load(bytes.ToArray(), resolution.FontFaceIndex);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
        {
            return null;
        }
    }
}
