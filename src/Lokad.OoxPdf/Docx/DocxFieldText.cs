using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxFieldText
{
    public static string? Placeholder(DocxFieldKind kind) => kind switch
    {
        DocxFieldKind.Page => "{PAGE}",
        DocxFieldKind.NumPages => "{NUMPAGES}",
        _ => null
    };

    private static int? Value(DocxFieldKind? kind, int? pageNumber, int? pageCount) => kind switch
    {
        DocxFieldKind.Page => pageNumber,
        DocxFieldKind.NumPages => pageCount,
        _ => null
    };

    public static bool CanResolve(DocxFieldKind kind, int? pageNumber, int? pageCount) =>
        Value(kind, pageNumber, pageCount) is not null;

    public static bool TryResolve(DocxTextRun run, string text, int? pageNumber, int? pageCount, out string resolved)
    {
        resolved = text;
        if (run.FieldKind is not { } kind || Value(kind, pageNumber, pageCount) is not { } value ||
            Placeholder(kind) is not { } placeholder) return false;
        resolved = text.Replace(placeholder, value.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
        return true;
    }

    public static string Resolve(DocxTextRun run, string text, int? pageNumber, int? pageCount) =>
        TryResolve(run, text, pageNumber, pageCount, out string resolved) ? resolved : text;
}
