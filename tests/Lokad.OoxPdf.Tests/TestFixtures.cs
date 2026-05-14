using System.IO.Compression;
using System.Text;

namespace Lokad.OoxPdf.Tests;

internal static class TestFixtures
{
    public static MemoryStream CreateZipPackage(IReadOnlyDictionary<string, string> entries)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach ((string name, string content) in entries)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using Stream entryStream = entry.Open();
                byte[] bytes = Encoding.UTF8.GetBytes(content);
                entryStream.Write(bytes);
            }
        }

        stream.Position = 0;
        return stream;
    }

    public static MemoryStream CreateZipPackage(IReadOnlyDictionary<string, byte[]> entries)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach ((string name, byte[] content) in entries)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using Stream entryStream = entry.Open();
                entryStream.Write(content);
            }
        }

        stream.Position = 0;
        return stream;
    }

    public static string WriteTempPackage(string extension, IReadOnlyDictionary<string, string> entries)
    {
        using MemoryStream stream = CreateZipPackage(entries);
        string path = Path.ChangeExtension(Path.GetTempFileName(), extension);
        File.WriteAllBytes(path, stream.ToArray());
        return path;
    }

    public static string WriteTempPackage(string extension, IReadOnlyDictionary<string, byte[]> entries)
    {
        using MemoryStream stream = CreateZipPackage(entries);
        string path = Path.ChangeExtension(Path.GetTempFileName(), extension);
        File.WriteAllBytes(path, stream.ToArray());
        return path;
    }

    public static byte[] Utf8(string text)
    {
        return Encoding.UTF8.GetBytes(text);
    }

    public static byte[] CreateRgbPng(int width, int height, byte[] rgb)
    {
        using var stream = new MemoryStream();
        stream.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        Span<byte> ihdr = stackalloc byte[13];
        WriteInt32(ihdr[..4], width);
        WriteInt32(ihdr.Slice(4, 4), height);
        ihdr[8] = 8;
        ihdr[9] = 2;
        WritePngChunk(stream, "IHDR", ihdr);

        byte[] raw = new byte[height * (1 + width * 3)];
        int source = 0;
        int target = 0;
        for (int y = 0; y < height; y++)
        {
            raw[target++] = 0;
            for (int x = 0; x < width * 3; x++)
            {
                raw[target++] = rgb[source++];
            }
        }

        using var compressed = new MemoryStream();
        using (var zlib = new ZLibStream(compressed, CompressionLevel.Fastest, leaveOpen: true))
        {
            zlib.Write(raw);
        }

        WritePngChunk(stream, "IDAT", compressed.ToArray());
        WritePngChunk(stream, "IEND", []);
        return stream.ToArray();
    }

    private static void WritePngChunk(Stream stream, string type, ReadOnlySpan<byte> data)
    {
        Span<byte> header = stackalloc byte[8];
        WriteInt32(header[..4], data.Length);
        Encoding.ASCII.GetBytes(type, header[4..]);
        stream.Write(header);
        stream.Write(data);

        byte[] crcInput = new byte[4 + data.Length];
        Encoding.ASCII.GetBytes(type, crcInput.AsSpan(0, 4));
        data.CopyTo(crcInput.AsSpan(4));
        Span<byte> crc = stackalloc byte[4];
        WriteUInt32(crc, ComputeCrc32(crcInput));
        stream.Write(crc);
    }

    private static uint ComputeCrc32(ReadOnlySpan<byte> bytes)
    {
        uint crc = 0xFFFFFFFFu;
        foreach (byte value in bytes)
        {
            crc ^= value;
            for (int i = 0; i < 8; i++)
            {
                crc = (crc & 1) != 0 ? 0xEDB88320u ^ (crc >> 1) : crc >> 1;
            }
        }

        return crc ^ 0xFFFFFFFFu;
    }

    private static void WriteInt32(Span<byte> target, int value)
    {
        target[0] = (byte)(value >> 24);
        target[1] = (byte)(value >> 16);
        target[2] = (byte)(value >> 8);
        target[3] = (byte)value;
    }

    private static void WriteUInt32(Span<byte> target, uint value)
    {
        target[0] = (byte)(value >> 24);
        target[1] = (byte)(value >> 16);
        target[2] = (byte)(value >> 8);
        target[3] = (byte)value;
    }
}
