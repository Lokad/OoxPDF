using System.Globalization;

namespace Lokad.OoxPdf.Docx;

[Flags]
internal enum DocxCounterDependencies { None = 0, PageNumber = 1, PageCount = 2 }

internal static class DocxFieldText
{
    public static DocxCounterDependencies Dependencies(DocxFieldKind? kind) => kind switch
    {
        DocxFieldKind.Page => DocxCounterDependencies.PageNumber,
        DocxFieldKind.NumPages => DocxCounterDependencies.PageCount,
        _ => DocxCounterDependencies.None
    };

    public static bool HasInputs(DocxCounterDependencies dependencies, int? pageNumber, int? pageCount) =>
        ((dependencies & DocxCounterDependencies.PageNumber) == 0 || pageNumber is not null) &&
        ((dependencies & DocxCounterDependencies.PageCount) == 0 || pageCount is not null);

    // Source ownership follows content through cells and inline boxes. Layout
    // caches and finalized-count retries consume the same dependency contract.
    public static DocxCounterDependencies GetDependencies(IReadOnlyList<DocxBodyElement> elements, CancellationToken token = default)
    {
        var dependencies = DocxCounterDependencies.None;
        foreach (var element in elements)
        {
            token.ThrowIfCancellationRequested();
            if (element is DocxParagraphElement paragraph)
            {
                foreach (var run in paragraph.Paragraph.Runs)
                {
                    token.ThrowIfCancellationRequested();
                    dependencies |= Dependencies(run.FieldKind);
                }
                foreach (var box in paragraph.Paragraph.InlineTextBoxes)
                    dependencies |= GetDependencies(box.BodyElements, token);
            }
            else if (element is DocxTableElement table)
                foreach (var row in table.Table.Rows)
                {
                    token.ThrowIfCancellationRequested();
                    foreach (var cell in row.Cells)
                        dependencies |= GetDependencies(DocxTableCellContent.GetBodyElements(cell), token);
                }
        }
        return dependencies;
    }

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
