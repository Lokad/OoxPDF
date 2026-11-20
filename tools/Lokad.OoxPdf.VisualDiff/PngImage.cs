using System.Buffers.Binary;
using System.IO.Compression;

namespace Lokad.OoxPdf.VisualDiff;

internal sealed class PngImage
{
    private static readonly byte[] Signature = [137, 80, 78, 71, 13, 10, 26, 10];

    // shared pixel/dimension budget with the rasterizer so its output
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

    private const int HeaderLength = 8 + 8 + 13 + 4;
    private const long MaxPngInputBytes = 512L * 1024L * 1024L;

    public static (PngImage? Reference, PngImage? Candidate) LoadPair(
        string? referencePath, string? candidatePath, long maxCombinedPixels)
    {
        using FileStream? reference = referencePath is null ? null : File.OpenRead(referencePath);
        using FileStream? candidate = candidatePath is null ? null : File.OpenRead(candidatePath);
        Header referenceHeader = reference is null ? default : ReadHeader(reference);
        Header candidateHeader = candidate is null ? default : ReadHeader(candidate);

        // Admit the pair from IHDR before buffering the compressed data or decoding pixels.
        long combined = (long)referenceHeader.Width * referenceHeader.Height +
            (long)candidateHeader.Width * candidateHeader.Height;
        if (reference is not null && candidate is not null && combined > maxCombinedPixels)
        {
            throw new InvalidDataException($"PNG pair pixel count {combined} exceeds the maximum of {maxCombinedPixels} (--max-combined-pixels).");
        }

        return (reference is null ? null : Load(reference, referenceHeader),
            candidate is null ? null : Load(candidate, candidateHeader));
    }

    private static Header ReadHeader(FileStream input)
    {
        long length = input.Length;
        if (length > MaxPngInputBytes)
        {
            throw new InvalidDataException($"PNG input exceeds the maximum inspectable size of {MaxPngInputBytes} bytes.");
        }

        Span<byte> header = stackalloc byte[HeaderLength];
        int read = input.ReadAtLeast(header, HeaderLength, false);
        if (read < Signature.Length || !header[..Signature.Length].SequenceEqual(Signature))
        {
            throw new InvalidDataException("File is not a PNG image.");
        }
        if (length < HeaderLength || read < HeaderLength)
        {
            throw new InvalidDataException("PNG chunk header is truncated.");
        }
        if (BinaryPrimitives.ReadInt32BigEndian(header[8..12]) != 13 || !header[12..16].SequenceEqual("IHDR"u8))
        {
            throw new InvalidDataException("PNG image header is missing.");
        }

        int width = BinaryPrimitives.ReadInt32BigEndian(header[16..20]);
        int height = BinaryPrimitives.ReadInt32BigEndian(header[20..24]);
        int bitDepth = header[24];
        int colorType = header[25];
        if (width <= 0 || height <= 0 || width > MaxDimension || height > MaxDimension)
        {
            throw new InvalidDataException($"PNG dimensions are out of range: {width}x{height}.");
        }
        long pixelCount = (long)width * height;
        if (pixelCount > MaxPixels)
        {
            throw new InvalidDataException($"PNG pixel count {pixelCount} exceeds the maximum of {MaxPixels}.");
        }
        if (!IsSupportedFormat(bitDepth, colorType) || header[28] != 0)
        {
            throw new NotSupportedException("Only non-interlaced grayscale, indexed, truecolor, and truecolor-alpha PNGs with common bit depths are supported.");
        }

        return new Header(width, height, bitDepth, colorType, (int)(length - HeaderLength));
    }

    private static PngImage Load(Stream input, Header header)
    {
        var bytes = new byte[header.RemainingBytes];
        input.ReadExactly(bytes);
        using var idat = new MemoryStream();
        byte[]? palette = null;
        byte[]? transparency = null;

        int offset = 0;
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

            ReadOnlySpan<byte> type = bytes.AsSpan(offset + 4, 4);
            offset += 8;
            ReadOnlySpan<byte> data = bytes.AsSpan(offset, length);
            offset += length + 4;

            if (type.SequenceEqual("IHDR"u8))
            {
                throw new InvalidDataException("PNG contains more than one image header.");
            }
            if (type.SequenceEqual("PLTE"u8))
            {
                palette = data.ToArray();
            }
            else if (type.SequenceEqual("tRNS"u8))
            {
                transparency = data.ToArray();
            }
            else if (type.SequenceEqual("IDAT"u8))
            {
                idat.Write(data);
            }
            else if (type.SequenceEqual("IEND"u8))
            {
                break;
            }
        }

        int bitsPerPixel = header.ColorType == 6 ? 32 : header.ColorType == 2 ? 24 : header.BitDepth;
        byte[] decompressed = Inflate(idat.ToArray(), header.Width, header.Height, bitsPerPixel);
        return new PngImage(header.Width, header.Height,
            DecodeScanlines(decompressed, header.Width, header.Height, header.BitDepth, header.ColorType, palette, transparency));
    }

    private readonly record struct Header(int Width, int Height, int BitDepth, int ColorType, int RemainingBytes);

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
        // capped inflation instead of CopyTo: the exact inflated size is
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
