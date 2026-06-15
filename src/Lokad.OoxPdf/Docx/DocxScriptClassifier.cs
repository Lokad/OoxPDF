using System.Text;

namespace Lokad.OoxPdf.Docx;

internal static class DocxScriptClassifier
{
    public static bool IsComplexScriptText(string text)
    {
        foreach (Rune rune in text.EnumerateRunes())
        {
            if (IsComplexScriptRune(rune.Value))
            {
                return true;
            }
        }

        return false;
    }

    public static bool IsEastAsianText(string text)
    {
        foreach (Rune rune in text.EnumerateRunes())
        {
            if (IsEastAsianRune(rune.Value))
            {
                return true;
            }
        }

        return false;
    }

    public static bool IsComplexScriptRune(int codePoint)
    {
        return codePoint is >= 0x0590 and <= 0x08FF or
            >= 0x0900 and <= 0x0D7F or
            >= 0x0E00 and <= 0x0E7F or
            >= 0x0F00 and <= 0x109F or
            >= 0x1780 and <= 0x17FF or
            >= 0xA840 and <= 0xA87F or
            >= 0xFB1D and <= 0xFDFF or
            >= 0xFE70 and <= 0xFEFF;
    }

    public static bool IsEastAsianRune(int codePoint)
    {
        return codePoint is >= 0x2E80 and <= 0xA4CF or
            >= 0xAC00 and <= 0xD7AF or
            >= 0xF900 and <= 0xFAFF or
            >= 0xFE30 and <= 0xFE4F or
            >= 0xFF00 and <= 0xFFEF or
            >= 0x20000 and <= 0x2FA1F or
            >= 0x30000 and <= 0x323AF;
    }
}
