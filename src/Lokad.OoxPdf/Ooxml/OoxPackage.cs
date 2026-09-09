using System.IO.Compression;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal sealed class OoxPackage
{
    private const int MaxEntryCount = 10_000;
    internal const long MaxPartBytes = 64L * 1024L * 1024L;
    internal const long MaxTotalBytes = 256L * 1024L * 1024L;
    internal const long MaxContentTypesBytes = 4L * 1024L * 1024L;

    private readonly Dictionary<string, OoxPart> parts;

    private OoxPackage(Dictionary<string, OoxPart> parts, OoxContentTypes contentTypes)
    {
        this.parts = parts;
        ContentTypes = contentTypes;
    }

    public OoxContentTypes ContentTypes { get; }

    public IReadOnlyCollection<OoxPart> Parts => parts.Values;

    public static OoxPackage Open(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using FileStream stream = File.OpenRead(path);
        return Open(stream, cancellationToken);
    }

    public static OoxPackage Open(Stream stream, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        if (archive.Entries.Count > MaxEntryCount)
        {
            throw new InvalidDataException($"OOXML package has too many ZIP entries: {archive.Entries.Count}.");
        }

        ZipArchiveEntry? contentTypesEntry = archive.GetEntry("[Content_Types].xml")
            ?? throw new InvalidDataException("OOXML package is missing [Content_Types].xml.");

        // The declared entry length is untrusted: copy through a bounded buffer first
        // so an oversized Content_Types part fails before XML parsing amplifies it.
        byte[] contentTypesBytes;
        using (Stream contentTypesStream = contentTypesEntry.Open())
        {
            contentTypesBytes = CopyBounded(contentTypesStream, MaxContentTypesBytes, "[Content_Types].xml", cancellationToken);
        }

        OoxContentTypes contentTypes;
        using (var contentTypesBuffer = new MemoryStream(contentTypesBytes, writable: false))
        {
            contentTypes = OoxContentTypes.Parse(contentTypesBuffer, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        long totalBytes = 0;
        var parts = new Dictionary<string, OoxPart>(StringComparer.OrdinalIgnoreCase);

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.FullName.EndsWith('/'))
            {
                continue;
            }

            string partName = OoxPath.NormalizePartName(entry.FullName);
            if (entry.Length < 0 || entry.Length > MaxPartBytes)
            {
                throw new InvalidDataException($"OOXML part '{partName}' exceeds the maximum supported size.");
            }


            string? contentType = partName == "/[Content_Types].xml"
                ? "application/xml"
                : contentTypes.GetContentType(partName);
            if (contentType is null)
            {
                continue;
            }

            // Two archive entries normalizing to one part name are ambiguous: keep neither guess.
            if (parts.ContainsKey(partName))
            {
                throw new InvalidDataException($"OOXML package contains a duplicate part {partName}.");
            }

            // Declared lengths are untrusted: enforce the budgets on actual copied bytes.
            byte[] partBytes;
            using (Stream entryStream = entry.Open())
            {
                partBytes = CopyBounded(entryStream, MaxPartBytes, partName, cancellationToken);
            }

            totalBytes = checked(totalBytes + partBytes.LongLength);
            if (totalBytes > MaxTotalBytes)
            {
                throw new InvalidDataException("OOXML package exceeds the maximum supported uncompressed size.");
            }

            parts.Add(partName, new OoxPart(partName, contentType, partBytes));
        }

        return new OoxPackage(parts, contentTypes);
    }

    public OoxPart? GetPart(string partName)
    {
        return parts.TryGetValue(OoxPath.NormalizePartName(partName), out OoxPart? part) ? part : null;
    }

    public IReadOnlyList<OoxRelationship> GetRelationships(string sourcePartName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string relationshipPartName = OoxPath.GetRelationshipPartName(sourcePartName);
        OoxPart? relationshipPart = GetPart(relationshipPartName);
        if (relationshipPart is null)
        {
            return [];
        }

        using Stream stream = relationshipPart.OpenRead();
        return ParseRelationships(stream, sourcePartName, cancellationToken);
    }

    public static IReadOnlyList<OoxRelationship> ParseRelationships(Stream stream, string sourcePartName, CancellationToken cancellationToken)
    {

        XDocument document = SafeXml.Load(stream, cancellationToken);
        var relationships = new List<OoxRelationship>();
        var seenIds = new HashSet<string>(StringComparer.Ordinal);

        foreach (XElement element in document.Root?.Elements(OoxNamespaces.PackageRelationshipsNamespace + "Relationship") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string id = RequiredAttribute(element, "Id");
            if (!seenIds.Add(id))
            {
                throw new InvalidDataException("OOXML package contains a duplicate relationship id.");
            }
            string type = RequiredAttribute(element, "Type");
            type = OoxNamespaces.NormalizeRelationshipType(type);
            string target = RequiredAttribute(element, "Target");
            string? targetMode = (string?)element.Attribute("TargetMode");
            string? resolvedTarget = targetMode?.Equals("External", StringComparison.OrdinalIgnoreCase) == true
                ? null
                : OoxPath.ResolveRelationshipTarget(sourcePartName, target);

            relationships.Add(new OoxRelationship(id, type, target, targetMode, resolvedTarget));
        }

        return relationships;
    }

    private static byte[] CopyBounded(Stream source, long maxBytes, string what, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81920];
        using var destination = new MemoryStream();
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return destination.ToArray();
            }

            if (checked(destination.Length + read) > maxBytes)
            {
                throw new InvalidDataException($"OOXML part {what} exceeds the maximum supported size.");
            }

            destination.Write(buffer, 0, read);
        }
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required Relationship attribute '{name}'.");
    }
}
