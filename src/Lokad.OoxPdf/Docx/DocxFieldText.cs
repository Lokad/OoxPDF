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

    public static string Resolve(DocxTextRun run, string text, int? pageNumber, int? pageCount)
    {
        int? value = run.FieldKind switch
        {
            DocxFieldKind.Page => pageNumber,
            DocxFieldKind.NumPages => pageCount,
            _ => null
        };
        return value is not null && run.FieldKind is { } kind && Placeholder(kind) is { } placeholder
            ? text.Replace(placeholder, value.Value.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
            : text;
    }
}
