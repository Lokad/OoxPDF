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

        // stream rows into the compressor instead of staging a second
        // full-size raw buffer next to the bitmap.
        using var compressed = new MemoryStream();
        using (var zlib = new ZLibStream(compressed, CompressionLevel.Fastest, leaveOpen: true))
        {
            var row = new byte[1 + checked(width * 4)];
            int source = 0;
            for (int y = 0; y < height; y++)
            {
                row[0] = 0;
                for (int x = 0; x < width; x++)
                {
                    byte blue = bgra[source++];
                    byte green = bgra[source++];
                    byte red = bgra[source++];
                    byte alpha = bgra[source++];
                    int target = 1 + x * 4;
                    row[target] = red;
                    row[target + 1] = green;
                    row[target + 2] = blue;
                    row[target + 3] = alpha;
                }

                zlib.Write(row);
            }
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

        // incremental CRC instead of concatenating a second full-size copy.
        uint crc = Crc32.Init();
        Span<byte> typeBytes = stackalloc byte[4];
        Encoding.ASCII.GetBytes(type, typeBytes);
        crc = Crc32.Update(crc, typeBytes);
        crc = Crc32.Update(crc, data);
        Span<byte> crcBytes = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crcBytes, Crc32.Finish(crc));
        output.Write(crcBytes);
    }
}
