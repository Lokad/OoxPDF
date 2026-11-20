using Lokad.OoxPdf;

internal static class DocxMarkupArguments
{
    internal static bool TryParseMode(string value, out OoxPdfDocxMarkupMode mode)
    {
        switch (value.Trim().ToLowerInvariant())
        {
            case "final":
                mode = OoxPdfDocxMarkupMode.Final;
                return true;
            case "original":
                mode = OoxPdfDocxMarkupMode.Original;
                return true;
            case "simple":
            case "simple-markup":
                mode = OoxPdfDocxMarkupMode.SimpleMarkup;
                return true;
            case "all":
            case "all-markup":
                mode = OoxPdfDocxMarkupMode.AllMarkup;
                return true;
            default:
                mode = OoxPdfDocxMarkupMode.Final;
                return false;
        }
    }

    internal static bool TryParseGeometryMode(string value, out OoxPdfDocxMarkupGeometryMode mode)
    {
        switch (value.Trim().ToLowerInvariant())
        {
            case "preserve":
            case "preserve-layout":
            case "preserve-document-layout":
                mode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
                return true;
            case "reserve":
            case "reserve-margin":
            case "markup-margin":
            case "reserve-markup-margin":
                mode = OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin;
                return true;
            case "word":
            case "word-compatible":
            case "word-compatible-all-markup":
            case "office":
            case "office-compatible":
            case "office-compatible-all-markup":
                mode = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                return true;
            default:
                mode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
                return false;
        }
    }
}
