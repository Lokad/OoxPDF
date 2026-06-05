using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Lokad.OoxPdf.VisualDiff;

internal sealed class PngImage
{
    private static readonly byte[] Signature = [137, 80, 78, 71, 13, 10, 26, 10];

    private PngImage(int width, int height, byte[] rgba)
    {
        Width = width;
        Height = height;
        Rgba = rgba;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Rgba { get; }

    public static PngImage Load(string path)
    {
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

        int offset = Signature.Length;
        while (offset < bytes.Length)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset, 4));
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

        if (width <= 0 || height <= 0)
        {
            throw new InvalidDataException("PNG image is missing IHDR dimensions.");
        }

        byte[] decompressed = Inflate(idat.ToArray());
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

    private static byte[] Inflate(byte[] compressed)
    {
        using var input = new MemoryStream(compressed);
        using var zlib = new ZLibStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        zlib.CopyTo(output);
        return output.ToArray();
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
        int stride = (width * bitsPerPixel + 7) / 8;
        var previous = new byte[stride];
        var current = new byte[stride];
        var rgba = new byte[width * height * 4];
        int sourceOffset = 0;
        int targetOffset = 0;

        for (int y = 0; y < height; y++)
        {
            byte filter = decompressed[sourceOffset++];
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
