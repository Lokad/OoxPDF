namespace Lokad.OoxPdf.Ooxml;

internal static class OoxPath
{
    public static string NormalizePartName(string partName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(partName);

        string normalized = partName.Replace('\\', '/');
        if (!normalized.StartsWith('/'))
        {
            normalized = "/" + normalized;
        }

        string[] segments = normalized.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var cleanSegments = new List<string>(segments.Length);
        foreach (string segment in segments)
        {
            if (segment is "." or ".." || segment.Contains(':', StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsafe OOXML part path '{partName}'.");
            }

            cleanSegments.Add(segment);
        }

        if (cleanSegments.Count == 0)
        {
            throw new InvalidDataException("OOXML part path cannot refer to the package root.");
        }

        return "/" + string.Join('/', cleanSegments);
    }

    public static string ResolveRelationshipTarget(string sourcePartName, string target)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(target);

        if (target.Contains(':', StringComparison.Ordinal))
        {
            throw new InvalidDataException($"External or absolute relationship target '{target}' is not a package part.");
        }

        if (target.StartsWith('/'))
        {
            return NormalizePartName(target);
        }

        string normalizedSource = sourcePartName == "/" ? "/" : NormalizePartName(sourcePartName);
        string baseDirectory = normalizedSource == "/"
            ? "/"
            : normalizedSource[..(normalizedSource.LastIndexOf('/') + 1)];

        var segments = new List<string>();
        if (baseDirectory.Length > 1)
        {
            segments.AddRange(baseDirectory.Split('/', StringSplitOptions.RemoveEmptyEntries));
        }

        foreach (string segment in target.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count == 0)
                {
                    throw new InvalidDataException($"Relationship target '{target}' escapes the OOXML package root.");
                }

                segments.RemoveAt(segments.Count - 1);
                continue;
            }

            if (segment.Contains(':', StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsafe relationship target '{target}'.");
            }

            segments.Add(segment);
        }

        if (segments.Count == 0)
        {
            throw new InvalidDataException($"Relationship target '{target}' does not name a package part.");
        }

        return "/" + string.Join('/', segments);
    }

    public static string GetRelationshipPartName(string partName)
    {
        if (partName == "/")
        {
            return "/_rels/.rels";
        }

        string normalized = NormalizePartName(partName);
        int slash = normalized.LastIndexOf('/');
        string directory = normalized[..(slash + 1)];
        string fileName = normalized[(slash + 1)..];
        return $"{directory}_rels/{fileName}.rels";
    }
}
