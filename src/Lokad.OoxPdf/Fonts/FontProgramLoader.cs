namespace Lokad.OoxPdf.Fonts;

internal static class FontProgramLoader
{
    public static OpenTypeFont? Load(FontFaceResolution? resolution)
    {
        if (resolution is null)
        {
            return null;
        }

        try
        {
            ReadOnlyMemory<byte> bytes = resolution.Source.GetBytesAsync().AsTask().GetAwaiter().GetResult();
            return OpenTypeFont.Load(bytes.ToArray(), resolution.FontFaceIndex);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
        {
            return null;
        }
    }
}
