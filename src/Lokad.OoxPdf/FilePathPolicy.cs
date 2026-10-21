namespace Lokad.OoxPdf;

// A process OS does not determine the case behavior of its mounted directories.
// Resolve directory links and compare case-only spellings using existing entries.
// Leave final file links alone: publication replaces their separate entry.
internal static class FilePathPolicy
{
    internal static bool MayOverwriteInput(string input, string destination)
    {
        if (MayReferToSameEntry(input, destination)) { return true; }
        try
        {
            var file = new FileInfo(Path.GetFullPath(input));
            // Opening input follows its final link. Publication replaces only
            // its destination entry, so resolve this side of the comparison.
            if (file.LinkTarget is not null)
            {
                input = (file.ResolveLinkTarget(returnFinalTarget: true)
                    ?? throw new IOException("Input file link could not be resolved.")).FullName;
            }
            return MayReferToSameEntry(input, destination);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return true; }
    }

    internal static bool MayReferToSameEntry(string first, string second)
    {
        try
        {
            first = ResolveDirectoryLinks(Path.GetFullPath(first));
            second = ResolveDirectoryLinks(Path.GetFullPath(second));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // An inaccessible or cyclic directory link cannot establish separation.
            return true;
        }
        if (!string.Equals(first, second, StringComparison.OrdinalIgnoreCase)) { return false; }
        if (string.Equals(first, second, StringComparison.Ordinal)) { return true; }

        string? directory = Path.GetDirectoryName(first);
        if (directory is null || !string.Equals(directory, Path.GetDirectoryName(second), StringComparison.Ordinal))
        {
            // Resolving differently spelled ancestors would require a separate
            // filesystem identity contract. Case-only parent ambiguity refuses.
            return true;
        }

        bool firstExists = Path.Exists(first);
        bool secondExists = Path.Exists(second);
        if (firstExists != secondExists) { return false; }
        if (!firstExists) { return true; }

        // Both queries can succeed for one case-insensitive entry. Only two
        // separately enumerated spellings establish distinct existing entries.
        string firstName = Path.GetFileName(first);
        string secondName = Path.GetFileName(second);
        bool foundFirst = false, foundSecond = false;
        try
        {
            foreach (string entry in Directory.EnumerateFileSystemEntries(directory))
            {
                string name = Path.GetFileName(entry);
                foundFirst |= string.Equals(name, firstName, StringComparison.Ordinal);
                foundSecond |= string.Equals(name, secondName, StringComparison.Ordinal);
                if (foundFirst && foundSecond) { return false; }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        return true;
    }

    private static string ResolveDirectoryLinks(string path)
    {
        var activeLinks = new HashSet<string>(StringComparer.Ordinal);
        return Path.Combine(ResolveDirectory(Path.GetDirectoryName(path)!), Path.GetFileName(path));

        string ResolveDirectory(string directoryPath)
        {
            var ancestors = new Stack<string>();
            var directory = new DirectoryInfo(directoryPath);
            while (directory.Parent is { } parent)
            {
                ancestors.Push(directory.Name);
                directory = parent;
            }
            string resolved = directory.FullName;
            while (ancestors.TryPop(out string? name))
            {
                directory = new DirectoryInfo(Path.Combine(resolved, name));
                resolved = directory.FullName;
                if (directory.LinkTarget is null) { continue; }
                if (!activeLinks.Add(resolved)) { throw new IOException("Cyclic directory link."); }
                string linkPath = resolved;
                FileSystemInfo target = directory.ResolveLinkTarget(returnFinalTarget: true)
                    ?? throw new IOException("Directory link could not be resolved.");
                // Unix final-target resolution can retain links in target ancestors.
                resolved = ResolveDirectory(target.FullName);
                activeLinks.Remove(linkPath);
            }
            return resolved;
        }
    }
}
