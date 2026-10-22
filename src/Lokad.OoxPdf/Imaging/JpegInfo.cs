namespace Lokad.OoxPdf.Imaging;

internal readonly record struct JpegInfo(int Width, int Height, int BitsPerComponent, int ComponentCount, byte FrameMarker, bool InvertedCmyk = false)
{
    public bool IsBaselineDct => FrameMarker == 0xC0;

    public bool IsProgressiveDct => FrameMarker == 0xC2;

    public string FrameProfileName => FrameMarker switch
    {
        0xC0 => "baseline DCT",
        0xC1 => "extended sequential DCT",
        0xC2 => "progressive DCT",
        0xC3 => "lossless sequential",
        0xC5 => "differential sequential DCT",
        0xC6 => "differential progressive DCT",
        0xC7 => "differential lossless",
        0xC9 => "extended sequential arithmetic",
        0xCA => "progressive arithmetic",
        0xCB => "lossless arithmetic",
        0xCD => "differential sequential arithmetic",
        0xCE => "differential progressive arithmetic",
        0xCF => "differential lossless arithmetic",
        _ => "unknown"
    };

    public static JpegInfo Read(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
        {
            throw new InvalidDataException("Data is not a JPEG image.");
        }

        int offset = 2;
        JpegInfo? frame = null;
        byte? adobeTransform = null;
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
            if (marker is 0xD9 or 0xDA) { break; }
            if (marker == 0xD8)
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

            if (marker == 0xEE && length >= 7 && bytes.Slice(offset + 2, 5).SequenceEqual("Adobe"u8))
            {
                if (length < 14) { throw new InvalidDataException("JPEG Adobe marker is truncated."); }
                byte transform = bytes[offset + 13];
                if (adobeTransform is { } prior && prior != transform)
                {
                    throw new InvalidDataException("JPEG Adobe markers have conflicting color transforms.");
                }
                adobeTransform = transform;
            }

            if (marker is >= 0xC0 and <= 0xC3 or >= 0xC5 and <= 0xC7 or >= 0xC9 and <= 0xCB or >= 0xCD and <= 0xCF)
            {
                if (length < 8)
                {
                    break;
                }

                int bitsPerComponent = bytes[offset + 2];
                int height = (bytes[offset + 3] << 8) | bytes[offset + 4];
                int width = (bytes[offset + 5] << 8) | bytes[offset + 6];
                int componentCount = bytes[offset + 7];
                // Passthrough publishes these values directly into a PDF image.
                // Validate the frame header without decoding the entropy stream.
                if (width == 0 || height == 0)
                {
                    throw new InvalidDataException("JPEG dimensions must be positive.");
                }
                if (length != 8 + 3 * componentCount)
                {
                    throw new InvalidDataException("JPEG frame component table has an invalid length.");
                }
                if (bitsPerComponent != 8 || componentCount is not (1 or 3 or 4))
                {
                    throw new NotSupportedException("JPEG passthrough requires 8-bit grayscale, RGB or CMYK frame data.");
                }
                frame = new JpegInfo(width, height, bitsPerComponent, componentCount, marker);
            }

            offset += length;
        }

        if (frame is not { } info) { throw new InvalidDataException("JPEG dimensions were not found."); }
        if (info.ComponentCount == 4 && adobeTransform is not (null or 0 or 2))
        {
            throw new NotSupportedException("Unsupported four-component JPEG Adobe color transform.");
        }
        // Adobe CMYK and YCCK use inverted ink components. The PDF DCT filter
        // handles the encoded transform; Decode maps its samples to DeviceCMYK.
        return info with { InvertedCmyk = info.ComponentCount == 4 && adobeTransform is 0 or 2 };
    }
}
