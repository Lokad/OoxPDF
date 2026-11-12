namespace Lokad.OoxPdf.Docx;

internal static class DocxParagraphFragmentPolicy
{
    internal static bool HasWidowControl(DocxParagraph paragraph) =>
        paragraph.KeepRules.WidowControl != false;

    internal static bool HasOneLineHead(DocxParagraph paragraph, int lineCount, int fittingLineCount) =>
        HasWidowControl(paragraph) && lineCount > 1 && fittingLineCount == 1;

    internal static bool HasOneLineTail(DocxParagraph paragraph, int lineCount, int fittingLineCount) =>
        HasWidowControl(paragraph) && fittingLineCount > 0 && fittingLineCount < lineCount &&
        lineCount - fittingLineCount == 1;

    internal static int ResolveLineCount(DocxParagraph paragraph, int remainingLineCount,
        int fittingLineCount, bool deferSingleHead)
    {
        if (fittingLineCount > 1 && HasOneLineTail(paragraph, remainingLineCount, fittingLineCount))
            fittingLineCount--;
        return deferSingleHead && HasOneLineHead(paragraph, remainingLineCount, fittingLineCount)
            ? 0 : fittingLineCount;
    }
}
