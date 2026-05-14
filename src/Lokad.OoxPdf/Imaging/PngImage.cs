using System.Text;

namespace Lokad.OoxPdf.Imaging;

internal sealed class PngImage
{
    private static readonly byte[] Signature = [137, 80, 78, 71, 13, 10, 26, 10];

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
        int colorType = 0;
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
                byte bitDepth = data[8];
                colorType = data[9];
                byte interlace = data[12];
                if (bitDepth != 8 || colorType is not (2 or 6) || interlace != 0)
                {
                    throw new NotSupportedException("Only non-interlaced 8-bit truecolor and truecolor-alpha PNGs are supported.");
                }
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

        return Decode(output.ToArray(), width, height, colorType);
    }

    private static PngImage Decode(byte[] decompressed, int width, int height, int colorType)
    {
        int bpp = colorType == 6 ? 4 : 3;
        int stride = width * bpp;
        var previous = new byte[stride];
        var current = new byte[stride];
        var rgb = new byte[width * height * 3];
        byte[]? alpha = colorType == 6 ? new byte[width * height] : null;
        int source = 0;
        int rgbTarget = 0;
        int alphaTarget = 0;
        for (int y = 0; y < height; y++)
        {
            byte filter = decompressed[source++];
            decompressed.AsSpan(source, stride).CopyTo(current);
            source += stride;
            Unfilter(filter, current, previous, bpp);
            for (int x = 0; x < width; x++)
            {
                int p = x * bpp;
                rgb[rgbTarget++] = current[p];
                rgb[rgbTarget++] = current[p + 1];
                rgb[rgbTarget++] = current[p + 2];
                if (alpha is not null)
                {
                    alpha[alphaTarget++] = current[p + 3];
                }
            }

            (previous, current) = (current, previous);
            Array.Clear(current);
        }

        return new PngImage(width, height, rgb, alpha);
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
