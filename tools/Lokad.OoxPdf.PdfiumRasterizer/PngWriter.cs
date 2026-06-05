using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Lokad.OoxPdf.PdfiumRasterizer;

internal static class PngWriter
{
    private static readonly byte[] Signature = [137, 80, 78, 71, 13, 10, 26, 10];

    public static void WriteBgra(string path, int width, int height, byte[] bgra)
    {
        using FileStream output = File.Create(path);
        output.Write(Signature);

        byte[] ihdr = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(0, 4), width);
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4, 4), height);
        ihdr[8] = 8;
        ihdr[9] = 6;
        WriteChunk(output, "IHDR", ihdr);

        byte[] raw = new byte[height * (1 + width * 4)];
        int source = 0;
        int target = 0;
        for (int y = 0; y < height; y++)
        {
            raw[target++] = 0;
            for (int x = 0; x < width; x++)
            {
                byte blue = bgra[source++];
                byte green = bgra[source++];
                byte red = bgra[source++];
                byte alpha = bgra[source++];
                raw[target++] = red;
                raw[target++] = green;
                raw[target++] = blue;
                raw[target++] = alpha;
            }
        }

        using var compressed = new MemoryStream();
        using (var zlib = new ZLibStream(compressed, CompressionLevel.Fastest, leaveOpen: true))
        {
            zlib.Write(raw);
        }

        WriteChunk(output, "IDAT", compressed.ToArray());
        WriteChunk(output, "IEND", []);
    }

    private static void WriteChunk(Stream output, string type, byte[] data)
    {
        Span<byte> header = stackalloc byte[8];
        BinaryPrimitives.WriteInt32BigEndian(header[..4], data.Length);
        Encoding.ASCII.GetBytes(type, header[4..]);
        output.Write(header);
        output.Write(data);

        var crcInput = new byte[4 + data.Length];
        Encoding.ASCII.GetBytes(type, crcInput.AsSpan(0, 4));
        data.CopyTo(crcInput.AsSpan(4));
        Span<byte> crcBytes = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crcBytes, Crc32.Compute(crcInput));
        output.Write(crcBytes);
    }
}
