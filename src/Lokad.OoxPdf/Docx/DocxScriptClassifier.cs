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
}
