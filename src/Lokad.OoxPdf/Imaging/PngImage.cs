using System.Text;

namespace Lokad.OoxPdf.Imaging;

internal sealed class PngImage
{
    private static readonly byte[] Signature = [137, 80, 78, 71, 13, 10, 26, 10];
    private static readonly int[] Adam7StartX = [0, 4, 0, 2, 0, 1, 0];
    private static readonly int[] Adam7StartY = [0, 0, 4, 0, 2, 0, 1];
    private static readonly int[] Adam7StepX = [8, 8, 4, 4, 2, 2, 1];
    private static readonly int[] Adam7StepY = [8, 8, 8, 4, 4, 2, 2];

    private PngImage(int width, int height, byte[] rgb, byte[]? alpha)
    {
        Width = width;
        Height = height;
        Rgb = rgb;
        Alpha = alpha;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Rgb { get; }

    public byte[]? Alpha { get; }

    public static PngImage Read(byte[] bytes)
    {
        if (bytes.Length < Signature.Length || !bytes.AsSpan(0, Signature.Length).SequenceEqual(Signature))
        {
            throw new InvalidDataException("Data is not a PNG image.");
        }

        int width = 0;
        int height = 0;
        int bitDepth = 0;
        int colorType = 0;
        int interlace = 0;
        byte[]? palette = null;
        byte[]? transparency = null;
        using var idat = new MemoryStream();
        int offset = Signature.Length;
        while (offset + 8 <= bytes.Length)
        {
            int length = ReadInt32(bytes, offset);
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            offset += 8;
            ReadOnlySpan<byte> data = bytes.AsSpan(offset, length);
            offset += length + 4;

            if (type == "IHDR")
            {
                width = ReadInt32(data);
                height = ReadInt32(data[4..]);
                bitDepth = data[8];
                colorType = data[9];
                interlace = data[12];
                if (!IsSupportedFormat(bitDepth, colorType) || interlace is not (0 or 1))
                {
                    throw new NotSupportedException($"Unsupported PNG format: bitDepth={bitDepth}, colorType={colorType}, interlace={interlace}.");
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

        using var input = new MemoryStream(idat.ToArray());
        using var zlib = new System.IO.Compression.ZLibStream(input, System.IO.Compression.CompressionMode.Decompress);
        using var output = new MemoryStream();
        zlib.CopyTo(output);

        return Decode(output.ToArray(), width, height, bitDepth, colorType, interlace, palette, transparency);
    }

    private static bool IsSupportedFormat(int bitDepth, int colorType)
    {
        return colorType switch
        {
            0 => bitDepth is 1 or 2 or 4 or 8,
            2 => bitDepth == 8,
            3 => bitDepth is 1 or 2 or 4 or 8,
            4 => bitDepth == 8,
            6 => bitDepth == 8,
            _ => false
        };
    }

    private static PngImage Decode(byte[] decompressed, int width, int height, int bitDepth, int colorType, int interlace, byte[]? palette, byte[]? transparency)
    {
        if (colorType == 3 && (palette is null || palette.Length % 3 != 0))
        {
            throw new InvalidDataException("Indexed PNG is missing a valid palette.");
        }

        int bitsPerPixel = colorType switch
        {
            0 or 3 => bitDepth,
            2 => 24,
            4 => 16,
            6 => 32,
            _ => throw new NotSupportedException($"Unsupported PNG color type {colorType}.")
        };
        int filterBytesPerPixel = Math.Max(1, (bitsPerPixel + 7) / 8);
        var rgb = new byte[width * height * 3];
        byte[]? alpha = colorType is 3 or 4 or 6 || HasGrayscaleTransparency(colorType, transparency) ? new byte[width * height] : null;
        if (interlace == 1)
        {
            DecodeAdam7(decompressed, width, height, bitDepth, colorType, bitsPerPixel, filterBytesPerPixel, palette, transparency, rgb, alpha);
            return new PngImage(width, height, rgb, alpha);
        }

        int stride = (width * bitsPerPixel + 7) / 8;
        var previous = new byte[stride];
        var current = new byte[stride];
        int source = 0;
        for (int y = 0; y < height; y++)
        {
            byte filter = decompressed[source++];
            decompressed.AsSpan(source, stride).CopyTo(current);
            source += stride;
            Unfilter(filter, current, previous, filterBytesPerPixel);
            for (int x = 0; x < width; x++)
            {
                WritePixel(bitDepth, colorType, current, x, y * width + x, palette, transparency, rgb, alpha);
            }

            (previous, current) = (current, previous);
            Array.Clear(current);
        }

        return new PngImage(width, height, rgb, alpha);
    }

    private static void DecodeAdam7(byte[] decompressed, int width, int height, int bitDepth, int colorType, int bitsPerPixel, int filterBytesPerPixel, byte[]? palette, byte[]? transparency, byte[] rgb, byte[]? alpha)
    {
        int source = 0;
        for (int pass = 0; pass < 7; pass++)
        {
            int passWidth = Adam7Size(width, Adam7StartX[pass], Adam7StepX[pass]);
            int passHeight = Adam7Size(height, Adam7StartY[pass], Adam7StepY[pass]);
            if (passWidth == 0 || passHeight == 0)
            {
                continue;
            }

            int stride = (passWidth * bitsPerPixel + 7) / 8;
            var previous = new byte[stride];
            var current = new byte[stride];
            for (int row = 0; row < passHeight; row++)
            {
                byte filter = decompressed[source++];
                decompressed.AsSpan(source, stride).CopyTo(current);
                source += stride;
                Unfilter(filter, current, previous, filterBytesPerPixel);

                int y = Adam7StartY[pass] + row * Adam7StepY[pass];
                for (int x = 0; x < passWidth; x++)
                {
                    int finalX = Adam7StartX[pass] + x * Adam7StepX[pass];
                    WritePixel(bitDepth, colorType, current, x, y * width + finalX, palette, transparency, rgb, alpha);
                }

                (previous, current) = (current, previous);
                Array.Clear(current);
            }
        }
    }

    private static int Adam7Size(int size, int start, int step)
    {
        return size <= start ? 0 : (size - start + step - 1) / step;
    }

    private static void WritePixel(int bitDepth, int colorType, byte[] current, int x, int pixelIndex, byte[]? palette, byte[]? transparency, byte[] rgb, byte[]? alpha)
    {
        int rgbTarget = pixelIndex * 3;
        int alphaTarget = pixelIndex;
        switch (colorType)
        {
            case 0:
                int graySample = ReadPackedSample(current, x, bitDepth);
                byte gray = bitDepth == 8 ? (byte)graySample : (byte)(graySample * 255 / ((1 << bitDepth) - 1));
                rgb[rgbTarget++] = gray;
                rgb[rgbTarget++] = gray;
                rgb[rgbTarget++] = gray;
                if (alpha is not null)
                {
                    alpha[alphaTarget++] = MatchesTransparentGray(gray, transparency) ? (byte)0 : (byte)255;
                }

                break;

            case 2:
                int trueColor = x * 3;
                rgb[rgbTarget++] = current[trueColor];
                rgb[rgbTarget++] = current[trueColor + 1];
                rgb[rgbTarget++] = current[trueColor + 2];
                break;

            case 3:
                int paletteIndex = ReadPackedSample(current, x, bitDepth);
                int paletteOffset = paletteIndex * 3;
                if (palette is null || paletteOffset + 2 >= palette.Length)
                {
                    throw new InvalidDataException("Indexed PNG pixel references a missing palette entry.");
                }

                rgb[rgbTarget++] = palette[paletteOffset];
                rgb[rgbTarget++] = palette[paletteOffset + 1];
                rgb[rgbTarget++] = palette[paletteOffset + 2];
                if (alpha is not null)
                {
                    alpha[alphaTarget++] = transparency is not null && paletteIndex < transparency.Length ? transparency[paletteIndex] : (byte)255;
                }

                break;

            case 4:
                int grayAlphaOffset = x * 2;
                byte grayAlpha = current[grayAlphaOffset];
                rgb[rgbTarget++] = grayAlpha;
                rgb[rgbTarget++] = grayAlpha;
                rgb[rgbTarget++] = grayAlpha;
                if (alpha is not null)
                {
                    alpha[alphaTarget++] = current[grayAlphaOffset + 1];
                }

                break;

            case 6:
                int rgbaOffset = x * 4;
                rgb[rgbTarget++] = current[rgbaOffset];
                rgb[rgbTarget++] = current[rgbaOffset + 1];
                rgb[rgbTarget++] = current[rgbaOffset + 2];
                if (alpha is not null)
                {
                    alpha[alphaTarget++] = current[rgbaOffset + 3];
                }

                break;
        }
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

    private static bool HasGrayscaleTransparency(int colorType, byte[]? transparency)
    {
        return colorType == 0 && transparency is { Length: >= 2 };
    }

    private static bool MatchesTransparentGray(byte gray, byte[]? transparency)
    {
        return transparency is { Length: >= 2 } && transparency[0] == 0 && transparency[1] == gray;
    }

    private static void Unfilter(byte filter, byte[] current, byte[] previous, int bpp)
    {
        for (int i = 0; i < current.Length; i++)
        {
            int left = i >= bpp ? current[i - bpp] : 0;
            int up = previous[i];
            int upLeft = i >= bpp ? previous[i - bpp] : 0;
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

    private static int Paeth(int a, int b, int c)
    {
        int p = a + b - c;
        int pa = Math.Abs(p - a);
        int pb = Math.Abs(p - b);
        int pc = Math.Abs(p - c);
        return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
    }

    private static int ReadInt32(ReadOnlySpan<byte> bytes)
    {
        return (bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3];
    }

    private static int ReadInt32(byte[] bytes, int offset)
    {
        return ReadInt32(bytes.AsSpan(offset, 4));
    }
}
