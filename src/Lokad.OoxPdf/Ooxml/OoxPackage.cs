using System.IO.Compression;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal sealed class OoxPackage
{
    private const int MaxEntryCount = 10_000;
    private const long MaxPartBytes = 64L * 1024L * 1024L;
    private const long MaxTotalBytes = 256L * 1024L * 1024L;

    private readonly Dictionary<string, OoxPart> parts;

    private OoxPackage(Dictionary<string, OoxPart> parts, OoxContentTypes contentTypes)
    {
        this.parts = parts;
        ContentTypes = contentTypes;
    }

    public OoxContentTypes ContentTypes { get; }

    public IReadOnlyCollection<OoxPart> Parts => parts.Values;

    public static OoxPackage Open(string path)
    {
        using FileStream stream = File.OpenRead(path);
        return Open(stream);
    }

    public static OoxPackage Open(Stream stream)
    {
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        if (archive.Entries.Count > MaxEntryCount)
        {
            throw new InvalidDataException($"OOXML package has too many ZIP entries: {archive.Entries.Count}.");
        }

        ZipArchiveEntry? contentTypesEntry = archive.GetEntry("[Content_Types].xml")
            ?? throw new InvalidDataException("OOXML package is missing [Content_Types].xml.");

        using Stream contentTypesStream = contentTypesEntry.Open();
        OoxContentTypes contentTypes = OoxContentTypes.Parse(contentTypesStream);

        long totalBytes = 0;
        var parts = new Dictionary<string, OoxPart>(StringComparer.OrdinalIgnoreCase);

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (entry.FullName.EndsWith('/'))
            {
                continue;
            }

            string partName = OoxPath.NormalizePartName(entry.FullName);
            if (entry.Length > MaxPartBytes)
            {
                throw new InvalidDataException($"OOXML part '{partName}' exceeds the maximum supported size.");
            }

            totalBytes += entry.Length;
            if (totalBytes > MaxTotalBytes)
            {
                throw new InvalidDataException("OOXML package exceeds the maximum supported uncompressed size.");
            }

            string? contentType = partName == "/[Content_Types].xml"
                ? "application/xml"
                : contentTypes.GetContentType(partName);
            if (contentType is null)
            {
                continue;
            }

            using Stream entryStream = entry.Open();
            using var memory = new MemoryStream((int)entry.Length);
            entryStream.CopyTo(memory);
            parts[partName] = new OoxPart(partName, contentType, memory.ToArray());
        }

        return new OoxPackage(parts, contentTypes);
    }

    public OoxPart? GetPart(string partName)
    {
        return parts.TryGetValue(OoxPath.NormalizePartName(partName), out OoxPart? part) ? part : null;
    }

    public IReadOnlyList<OoxRelationship> GetRelationships(string sourcePartName)
    {
        string relationshipPartName = OoxPath.GetRelationshipPartName(sourcePartName);
        OoxPart? relationshipPart = GetPart(relationshipPartName);
        if (relationshipPart is null)
        {
            return [];
        }

        using Stream stream = relationshipPart.OpenRead();
        return ParseRelationships(stream, sourcePartName);
    }

    public static IReadOnlyList<OoxRelationship> ParseRelationships(Stream stream, string sourcePartName)
    {
        XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        XDocument document = SafeXml.Load(stream);
        var relationships = new List<OoxRelationship>();

        foreach (XElement element in document.Root?.Elements(relationshipsNamespace + "Relationship") ?? [])
        {
            string id = RequiredAttribute(element, "Id");
            string type = RequiredAttribute(element, "Type");
            string target = RequiredAttribute(element, "Target");
            string? targetMode = (string?)element.Attribute("TargetMode");
            string? resolvedTarget = targetMode?.Equals("External", StringComparison.OrdinalIgnoreCase) == true
                ? null
                : OoxPath.ResolveRelationshipTarget(sourcePartName, target);

            relationships.Add(new OoxRelationship(id, type, target, targetMode, resolvedTarget));
        }

        return relationships;
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required Relationship attribute '{name}'.");
    }
}
