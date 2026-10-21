namespace Lokad.OoxPdf;

// A process OS does not determine the case behavior of its mounted directories.
// Compare case-only spellings using existing entries; refuse ambiguity without
// probing by creating files or depending on native file identity APIs.
internal static class FilePathPolicy
{
    internal static bool MayReferToSameEntry(string first, string second)
    {
        first = Path.GetFullPath(first);
        second = Path.GetFullPath(second);
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
}
