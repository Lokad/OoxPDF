using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxInheritanceFixture
{
    public static Dictionary<string, byte[]> ReadParts(string path, string? expectedSha256 = null)
    {
        using Stream input = expectedSha256 is null
            ? File.OpenRead(path) : new MemoryStream(File.ReadAllBytes(path), writable: false);
        if (expectedSha256 is not null)
        {
            TestAssert.Equal(expectedSha256, Convert.ToHexString(SHA256.HashData(input)));
            input.Position = 0;
        }

        using var zip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        {
            using var source = entry.Open();
            using var buffer = new MemoryStream();
            source.CopyTo(buffer);
            parts.Add(entry.FullName, buffer.ToArray());
        }
        return parts;
    }

    public static void Edit(Dictionary<string, byte[]> parts, string name, Action<XDocument> edit)
    {
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[name]));
        edit(xml);
        parts[name] = Encoding.UTF8.GetBytes(xml.ToString());
    }

    public static void WithDeck(Dictionary<string, byte[]> parts, Action<PptxDocument, OoxPackage> check)
    {
        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        check(new PptxReader().Read(package, default), package);
    }
}
