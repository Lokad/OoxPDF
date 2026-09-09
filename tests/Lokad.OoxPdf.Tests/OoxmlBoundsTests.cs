using System.IO.Compression;
using System.Text;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class OoxmlBoundsTests
{

    public static void RejectsDuplicateNormalizedPartNames()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipText(archive, "[Content_Types].xml", ContentTypesWithXmlDefault());
            WriteZipText(archive, "word/document.xml", "<w:document/>");
            WriteZipText(archive, "WORD\\document.xml", "<w:document/>");
        }

        stream.Position = 0;
        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void RejectsExactDuplicatePartNames()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipText(archive, "[Content_Types].xml", ContentTypesWithXmlDefault());
            WriteZipText(archive, "word/document.xml", "<w:document/>");
            WriteZipText(archive, "word/document.xml", "<w:document/>");
        }

        stream.Position = 0;
        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void RejectsOversizedContentTypes()
    {
        string padding = new string((char)97, 5 * 1024 * 1024);
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipText(archive, "[Content_Types].xml", "<Types><!--" + padding + "--></Types>");
        }

        stream.Position = 0;
        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void RejectsTruncatedArchives()
    {
        using var stream = new MemoryStream(new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x00 });
        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void RejectsConflictingContentTypeOverrides()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = TestContentTypes(
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/a\"/>",
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/b\"/>"),
        });

        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(packageStream, CancellationToken.None));
    }

    public static void ToleratesIdenticalContentTypeOverrides()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = TestContentTypes(
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/xml\"/>",
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/xml\"/>"),
            ["word/document.xml"] = "<w:document/>",
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        TestAssert.NotNull(package.GetPart("/word/document.xml"));
    }

    public static void RejectsDuplicateRelationshipIds()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = TestContentTypes(
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"),
            ["_rels/.rels"] = TestRelationships(
                "<Relationship Id=\"rId1\" Type=\"officeDocument\" Target=\"word/document.xml\"/>",
                "<Relationship Id=\"rId1\" Type=\"officeDocument\" Target=\"word/other.xml\"/>"),
            ["word/document.xml"] = "<w:document/>",
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        TestAssert.Throws<InvalidDataException>(() => package.GetRelationships("/", CancellationToken.None));
    }

    public static void RejectsDeeplyNestedXml()
    {
        string nested = "<r>" + string.Concat(Enumerable.Repeat("<a>", 40)) + "x" + string.Concat(Enumerable.Repeat("</a>", 40)) + "</r>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(nested));
        TestAssert.Throws<InvalidDataException>(() => SafeXml.Load(stream, CancellationToken.None, 1024 * 1024, 32));
    }

    public static void RejectsDefaultDeepXmlNesting()
    {
        string nested = "<r>" + string.Concat(Enumerable.Repeat("<a>", 300)) + "x" + string.Concat(Enumerable.Repeat("</a>", 300)) + "</r>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(nested));
        TestAssert.Throws<InvalidDataException>(() => SafeXml.Load(stream, CancellationToken.None));
    }

    public static void RejectsXmlBeyondCharacterBudget()
    {
        string document = "<r>" + new string((char)120, 200) + "</r>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(document));
        TestAssert.Throws<InvalidDataException>(() => SafeXml.Load(stream, CancellationToken.None, 100, 256));
    }

    public static void RejectsMalformedXmlAsInvalidData()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<a><b></a>"));
        TestAssert.Throws<InvalidDataException>(() => SafeXml.Load(stream, CancellationToken.None));
    }

    private static string TestContentTypes(params string[] entries)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            string.Concat(entries) +
            "</Types>";
    }

    private static string TestRelationships(params string[] entries)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            string.Concat(entries) +
            "</Relationships>";
    }

    private static string ContentTypesWithXmlDefault()
    {
        return TestContentTypes(
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
    }

    private static void WriteZipText(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        using Stream entryStream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        entryStream.Write(bytes);
    }
}

