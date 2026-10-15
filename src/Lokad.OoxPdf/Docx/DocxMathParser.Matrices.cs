using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathParser
{
    private static DocxMathMatrix ReadMatrix(XElement element, IReadOnlyList<DocxMathMatrixRow> rows,
        Func<XElement, string, DocxTextRun> readRun, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, "mPr", out XElement? owner, out XElement[] properties);
        foreach (XElement property in properties)
        { token.ThrowIfCancellationRequested(); valid &= property.Name == MathNamespace + "ctrlPr" && HasKnownControl(property); }
        int columns = rows[0].Cells.Count;
        foreach (DocxMathMatrixRow row in rows)
        { token.ThrowIfCancellationRequested(); valid &= !row.HasUnsupportedProperties && row.Cells.Count == columns; }
        foreach (XElement row in ContentChildren(element))
        foreach (XElement cell in ContentChildren(row))
        {
            token.ThrowIfCancellationRequested();
            valid &= cell.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(cell);
        }
        XElement? control = owner?.Element(MathNamespace + "ctrlPr");
        // Keep absent/authored controls distinct from painted cell formatting.
        DocxTextRun? text = control is null ? null : readRun(control, string.Empty);
        return new(rows)
        { HasUnsupportedProperties = !valid, ControlText = text, ControlFontFamily = text is null ? null : ControlFontFamily(control, text) };
    }
}
