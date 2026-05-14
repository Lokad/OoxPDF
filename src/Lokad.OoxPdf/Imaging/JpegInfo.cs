namespace Lokad.OoxPdf.Imaging;

internal readonly record struct JpegInfo(int Width, int Height)
{
    public static JpegInfo Read(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
        {
            throw new InvalidDataException("Data is not a JPEG image.");
        }

        int offset = 2;
        while (offset + 4 < bytes.Length)
        {
            while (offset < bytes.Length && bytes[offset] == 0xFF)
            {
                offset++;
            }

            if (offset >= bytes.Length)
            {
                break;
            }

            byte marker = bytes[offset++];
            if (marker is 0xD8 or 0xD9)
            {
                continue;
            }

            if (offset + 2 > bytes.Length)
            {
                break;
            }

            int length = (bytes[offset] << 8) | bytes[offset + 1];
            if (length < 2 || offset + length > bytes.Length)
            {
                break;
            }

            if (marker is >= 0xC0 and <= 0xC3 or >= 0xC5 and <= 0xC7 or >= 0xC9 and <= 0xCB or >= 0xCD and <= 0xCF)
            {
                int height = (bytes[offset + 3] << 8) | bytes[offset + 4];
                int width = (bytes[offset + 5] << 8) | bytes[offset + 6];
                return new JpegInfo(width, height);
            }

            offset += length;
        }

        throw new InvalidDataException("JPEG dimensions were not found.");
    }
}
