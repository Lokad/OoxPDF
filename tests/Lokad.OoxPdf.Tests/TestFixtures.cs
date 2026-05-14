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

    public static string WriteTempPackage(string extension, IReadOnlyDictionary<string, string> entries)
    {
        using MemoryStream stream = CreateZipPackage(entries);
        string path = Path.ChangeExtension(Path.GetTempFileName(), extension);
        File.WriteAllBytes(path, stream.ToArray());
        return path;
    }
}
