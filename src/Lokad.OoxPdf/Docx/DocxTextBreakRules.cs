namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Character break rules shared by layout, its counters, and source navigation.
/// </summary>
internal static class DocxTextBreakRules
{
    internal static bool IsOptionalBreakMarker(char value) => value is '\u00AD' or '\u200B';

    internal static bool IsWordJoiner(char value) => value is '\u2060' or '\uFEFF';

    internal static int EndOfBreakableSpaces(string text, int start, int end)
    {
        int next = start + 1;
        while (next < end && IsBreakableWhitespaceChar(text[next])) next++;
        return next;
    }

    /// <summary>
    /// Returns true for whitespace characters that may end a line.
    /// </summary>
    internal static bool IsBreakableWhitespaceChar(char value)
    {
        return char.IsWhiteSpace(value) &&
            !IsNoBreakWhitespaceChar(value);
    }

    /// <summary>
    /// Returns true for whitespace characters that must not end a line.
    /// </summary>
    internal static bool IsNoBreakWhitespaceChar(char value)
    {
        return value is '\u00A0' or '\u202F' or '\u2007';
    }
}
