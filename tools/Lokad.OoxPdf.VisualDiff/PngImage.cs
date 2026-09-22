using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Lokad.OoxPdf.VisualDiff;

internal sealed class PngImage
{
    private static readonly byte[] Signature = [137, 80, 78, 71, 13, 10, 26, 10];

    // PLAN Q06: shared pixel/dimension budget with the rasterizer so its output
    // always fits here (67M pixels, 32768 per side).
    private const int MaxDimension = 32768;
    private const long MaxPixels = 67_108_864L;

    private PngImage(int width, int height, byte[] rgba)
    {
        Width = width;
        Height = height;
        Rgba = rgba;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Rgba { get; }

    // R21: header-only dimensions so the paired quota can be enforced before
    // either RGBA buffer exists. Rejects the same malformed headers as Load.
    public static (int Width, int Height) ReadDimensions(string path)
    {
        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new InvalidDataException($"Cannot read PNG input '{path}': {ex.Message}", ex);
        }

        if (bytes.Length < Signature.Length || !bytes.AsSpan(0, Signature.Length).SequenceEqual(Signature))
        {
            throw new InvalidDataException("File is not a PNG image.");
        }

        if (bytes.Length < Signature.Length + 8 + 13)
        {
            throw new InvalidDataException("PNG chunk header is truncated.");
        }

        int length = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(Signature.Length, 4));
        string type = Encoding.ASCII.GetString(bytes, Signature.Length + 4, 4);
        if (length != 13 || !string.Equals(type, "IHDR", StringComparison.Ordinal))
        {
            throw new InvalidDataException("PNG image header is missing.");
        }

        ReadOnlySpan<byte> data = bytes.AsSpan(Signature.Length + 8, 13);
        int width = BinaryPrimitives.ReadInt32BigEndian(data[..4]);
        int height = BinaryPrimitives.ReadInt32BigEndian(data.Slice(4, 4));
        int bitDepth = data[8];
        int colorType = data[9];
        byte interlace = data[12];
        if (width <= 0 || height <= 0 || width > MaxDimension || height > MaxDimension)
        {
            throw new InvalidDataException($"PNG dimensions are out of range: {width}x{height}.");
        }

        if (!IsSupportedFormat(bitDepth, colorType) || interlace != 0)
        {
            throw new NotSupportedException("Only non-interlaced grayscale, indexed, truecolor, and truecolor-alpha PNGs with common bit depths are supported.");
        }

        return (width, height);
    }

    public static PngImage Load(string path)
    {
        // PLAN Q06: fail fast on absurd inputs before buffering the whole file.
        // Legitimate rasterizer/diff PNGs fit well under this (67M pixels cap).
        const long MaxPngInputBytes = 512L * 1024L * 1024L;
        if (new FileInfo(path).Length > MaxPngInputBytes)
        {
            throw new InvalidDataException($"PNG input exceeds the maximum inspectable size of {MaxPngInputBytes} bytes.");
        }

        byte[] bytes = File.ReadAllBytes(path);
        if (bytes.Length < Signature.Length || !bytes.AsSpan(0, Signature.Length).SequenceEqual(Signature))
        {
            throw new InvalidDataException("File is not a PNG image.");
        }

        int width = 0;
        int height = 0;
        int bitDepth = 0;
        int colorType = 0;
        var idat = new MemoryStream();
        byte[]? palette = null;
        byte[]? transparency = null;

        // PLAN Q06: chunk framing is untrusted. Bounds-check every slice so truncated
        // files fail with InvalidDataException instead of runtime slicing errors.
        int offset = Signature.Length;
        while (offset < bytes.Length)
        {
            if (offset + 8 > bytes.Length)
            {
                throw new InvalidDataException("PNG chunk header is truncated.");
            }

            int length = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset, 4));
            if (length < 0 || (long)offset + 8L + length + 4L > bytes.Length)
            {
                throw new InvalidDataException("PNG chunk extends beyond the end of the data.");
            }

            offset += 4;
            string type = Encoding.ASCII.GetString(bytes, offset, 4);
            offset += 4;
            ReadOnlySpan<byte> data = bytes.AsSpan(offset, length);
            offset += length + 4;

            if (type == "IHDR")
            {
                width = BinaryPrimitives.ReadInt32BigEndian(data[..4]);
                height = BinaryPrimitives.ReadInt32BigEndian(data.Slice(4, 4));
                bitDepth = data[8];
                colorType = data[9];
                byte interlace = data[12];
                if (!IsSupportedFormat(bitDepth, colorType) || interlace != 0)
                {
                    throw new NotSupportedException("Only non-interlaced grayscale, indexed, truecolor, and truecolor-alpha PNGs with common bit depths are supported.");
                }
            }
            else if (type == "PLTE")
            {
                palette = data.ToArray();
            }
            else if (type == "tRNS")
            {
                transparency = data.ToArray();
            }
            else if (type == "IDAT")
            {
                idat.Write(data);
            }
            else if (type == "IEND")
            {
                break;
            }
        }

        if (width <= 0 || height <= 0 || width > MaxDimension || height > MaxDimension)
        {
            throw new InvalidDataException($"PNG dimensions are out of range: {width}x{height}.");
        }

        long pixelCount;
        try
        {
            pixelCount = checked((long)width * height);
        }
        catch (OverflowException ex)
        {
            throw new InvalidDataException("PNG dimensions overflow.", ex);
        }

        if (pixelCount > MaxPixels)
        {
            throw new InvalidDataException($"PNG pixel count {pixelCount} exceeds the maximum of {MaxPixels}.");
        }

        byte[] decompressed = Inflate(idat.ToArray(), width, height, colorType == 6 ? 32 : colorType == 2 ? 24 : bitDepth);
        return new PngImage(width, height, DecodeScanlines(decompressed, width, height, bitDepth, colorType, palette, transparency));
    }

    private static bool IsSupportedFormat(int bitDepth, int colorType)
    {
        return colorType switch
        {
            0 => bitDepth is 1 or 2 or 4 or 8,
            2 => bitDepth == 8,
            3 => bitDepth is 1 or 2 or 4 or 8,
            6 => bitDepth == 8,
            _ => false
        };
    }

    private static byte[] Inflate(byte[] compressed, int width, int height, int bitsPerPixel)
    {
        // PLAN Q06: capped inflation instead of CopyTo: the exact inflated size is
        // stride+filter per row, so anything beyond it is hostile.
        long maxInflated;
        try
        {
            maxInflated = checked(((long)width * bitsPerPixel + 7L) / 8L + 1L) * height;
        }
        catch (OverflowException ex)
        {
            throw new InvalidDataException("PNG dimensions overflow.", ex);
        }

        using var input = new MemoryStream(compressed);
        using var zlib = new ZLibStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        byte[] buffer = new byte[81920];
        while (true)
        {
            int read = zlib.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return output.ToArray();
            }

            if (checked(output.Length + read) > maxInflated)
            {
                throw new InvalidDataException("PNG pixel data exceeds the size implied by its dimensions.");
            }

            output.Write(buffer, 0, read);
        }
    }

    private static byte[] DecodeScanlines(byte[] decompressed, int width, int height, int bitDepth, int colorType, byte[]? palette, byte[]? transparency)
    {
        int bitsPerPixel = colorType switch
        {
            0 => bitDepth,
            2 => 24,
            3 => bitDepth,
            6 => 32,
            _ => throw new NotSupportedException($"Unsupported PNG color type {colorType}.")
        };
        int filterBytesPerPixel = Math.Max(1, (bitsPerPixel + 7) / 8);
        int stride;
        long pixelBytes;
        try
        {
            stride = checked((width * bitsPerPixel + 7) / 8);
            pixelBytes = checked((long)width * height * 4L);
        }
        catch (OverflowException ex)
        {
            throw new InvalidDataException("PNG dimensions overflow.", ex);
        }

        var previous = new byte[stride];
        var current = new byte[stride];
        var rgba = new byte[pixelBytes];
        int sourceOffset = 0;
        int targetOffset = 0;

        for (int y = 0; y < height; y++)
        {
            if ((uint)sourceOffset >= (uint)decompressed.Length)
            {
                throw new InvalidDataException("PNG pixel data is truncated.");
            }

            byte filter = decompressed[sourceOffset++];
            if ((long)sourceOffset + stride > decompressed.Length)
            {
                throw new InvalidDataException("PNG pixel data is truncated.");
            }

            decompressed.AsSpan(sourceOffset, stride).CopyTo(current);
            sourceOffset += stride;
            Unfilter(filter, current, previous, filterBytesPerPixel);

            for (int x = 0; x < width; x++)
            {
                WritePixel(current, x, bitDepth, colorType, palette, transparency, rgba, ref targetOffset);
            }

            (previous, current) = (current, previous);
            Array.Clear(current);
        }

        return rgba;
    }

    private static void WritePixel(byte[] current, int x, int bitDepth, int colorType, byte[]? palette, byte[]? transparency, byte[] rgba, ref int targetOffset)
    {
        if (colorType == 6)
        {
            int pixel = x * 4;
            rgba[targetOffset++] = current[pixel];
            rgba[targetOffset++] = current[pixel + 1];
            rgba[targetOffset++] = current[pixel + 2];
            rgba[targetOffset++] = current[pixel + 3];
            return;
        }

        if (colorType == 2)
        {
            int pixel = x * 3;
            rgba[targetOffset++] = current[pixel];
            rgba[targetOffset++] = current[pixel + 1];
            rgba[targetOffset++] = current[pixel + 2];
            rgba[targetOffset++] = byte.MaxValue;
            return;
        }

        int sample = ReadPackedSample(current, x, bitDepth);
        if (colorType == 0)
        {
            byte gray = bitDepth == 8 ? (byte)sample : (byte)(sample * 255 / ((1 << bitDepth) - 1));
            rgba[targetOffset++] = gray;
            rgba[targetOffset++] = gray;
            rgba[targetOffset++] = gray;
            rgba[targetOffset++] = byte.MaxValue;
            return;
        }

        if (palette is null)
        {
            throw new InvalidDataException("Indexed PNG is missing a PLTE chunk.");
        }

        int paletteOffset = sample * 3;
        if (paletteOffset + 2 >= palette.Length)
        {
            throw new InvalidDataException("Indexed PNG refers to a missing palette entry.");
        }

        rgba[targetOffset++] = palette[paletteOffset];
        rgba[targetOffset++] = palette[paletteOffset + 1];
        rgba[targetOffset++] = palette[paletteOffset + 2];
        rgba[targetOffset++] = transparency is not null && sample < transparency.Length ? transparency[sample] : byte.MaxValue;
    }

    private static int ReadPackedSample(byte[] current, int x, int bitDepth)
    {
        if (bitDepth == 8)
        {
            return current[x];
        }

        int samplesPerByte = 8 / bitDepth;
        int byteIndex = x / samplesPerByte;
        int shift = (samplesPerByte - 1 - (x % samplesPerByte)) * bitDepth;
        int mask = (1 << bitDepth) - 1;
        return (current[byteIndex] >> shift) & mask;
    }

    private static void Unfilter(byte filter, byte[] current, byte[] previous, int bytesPerPixel)
    {
        for (int i = 0; i < current.Length; i++)
        {
            int left = i >= bytesPerPixel ? current[i - bytesPerPixel] : 0;
            int up = previous[i];
            int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;
            int predictor = filter switch
            {
                0 => 0,
                1 => left,
                2 => up,
                3 => (left + up) / 2,
                4 => Paeth(left, up, upLeft),
                _ => throw new InvalidDataException($"Unsupported PNG filter type {filter}.")
            };

            current[i] = unchecked((byte)(current[i] + predictor));
        }
    }

    private static int Paeth(int left, int up, int upLeft)
    {
        int p = left + up - upLeft;
        int pa = Math.Abs(p - left);
        int pb = Math.Abs(p - up);
        int pc = Math.Abs(p - upLeft);
        if (pa <= pb && pa <= pc)
        {
            return left;
        }

        return pb <= pc ? up : upLeft;
    }
}
