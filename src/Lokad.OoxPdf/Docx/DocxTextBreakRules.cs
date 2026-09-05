namespace Lokad.OoxPdf.Docx;

/// <summary>
/// Character break rules shared by the wrap engine and the layout snapshot counters.
/// </summary>
internal static class DocxTextBreakRules
{
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
