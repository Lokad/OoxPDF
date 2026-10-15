using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathParser
{
    private static DocxMathMatrix ReadMatrix(XElement element, IReadOnlyList<DocxMathMatrixRow> rows,
        Func<XElement, string, DocxTextRun> readRun, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, "mPr", out XElement? owner, out XElement[] properties);
        IReadOnlyList<DocxMathMatrixColumn>? columnGroups = null;
        string? baseJustification = null;
        foreach (XElement property in properties)
        {
            token.ThrowIfCancellationRequested();
            if (property.Name == MathNamespace + "mcs")
            {
                valid &= property.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(property) && property.HasElements;
                var groups = new List<DocxMathMatrixColumn>();
                foreach (XElement column in property.Elements())
                { token.ThrowIfCancellationRequested(); var group = ReadMatrixColumn(column, token); groups.Add(group); valid &= !group.HasUnsupportedProperties; }
                columnGroups = groups;
            }
            else if (property.Name == MathNamespace + "baseJc")
            {
                baseJustification = (string?)property.Attribute(MathNamespace + "val");
                valid &= HasKnownValueProperty(property) && baseJustification is null or "center" or "top" or "bot";
            }
            else if (property.Name.LocalName is "cGpRule" or "cGp" or "cSp" or "rSpRule" or "rSp")
            {
                // All five omitted/val-less scalar defaults are zero. Only the
                // existing default spacing profile is admitted here.
                string? value = (string?)property.Attribute(MathNamespace + "val");
                valid &= HasKnownValueProperty(property) && (value is null ||
                    uint.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out uint number) && number == 0);
            }
            else if (property.Name == MathNamespace + "plcHide")
            { valid &= HasKnownValueProperty(property) && HasBooleanValue((string?)property.Attribute(MathNamespace + "val")); }
            else { valid &= property.Name == MathNamespace + "ctrlPr" && HasKnownControl(property); }
        }
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
        { ColumnGroups = columnGroups, BaseJustificationValue = baseJustification, HasUnsupportedProperties = !valid,
            HidePlaceholdersValue = ReadBoolean(owner, "plcHide"),
            ControlText = text, ControlFontFamily = text is null ? null : ControlFontFamily(control, text) };
    }

    private static DocxMathMatrixColumn ReadMatrixColumn(XElement element, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, "mcPr", out XElement? owner, out XElement[] properties) &&
            element.Name == MathNamespace + "mc" && !ContentChildren(element).Any();
        foreach (XElement property in properties)
        {
            token.ThrowIfCancellationRequested();
            valid &= HasKnownValueProperty(property) && property.Name.LocalName is "count" or "mcJc";
        }
        string? countText = (string?)owner?.Element(MathNamespace + "count")?.Attribute(MathNamespace + "val");
        int? count = null;
        if (countText is not null)
        {
            bool parsed = int.TryParse(countText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int value);
            count = value; valid &= parsed && value is >= 1 and <= 255;
        }
        string? justification = (string?)owner?.Element(MathNamespace + "mcJc")?.Attribute(MathNamespace + "val");
        valid &= justification is null or "center" or "left" or "right";
        return new(count, justification, !valid);
    }
}
