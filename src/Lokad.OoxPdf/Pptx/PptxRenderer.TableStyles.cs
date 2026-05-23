using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryReadBuiltInTableStyleCellFill(XElement table, int rowIndex, PptxTheme theme, out RgbColor color, out double alpha)
    {
        alpha = 1d;
        if (!TryReadBuiltInTableStyle(table, out BuiltInTableStyle style) ||
            !string.Equals(style.Name, "Medium-Style-2", StringComparison.Ordinal) ||
            !theme.TryResolveColor(style.Accent, out RgbColor accent))
        {
            color = default;
            return false;
        }

        XElement? tableProperties = table.Element(DrawingNamespace + "tblPr");
        bool firstRow = ParseOptionalBoolAttribute(tableProperties, "firstRow");
        if (firstRow && rowIndex == 0)
        {
            color = accent;
            return true;
        }

        bool bandRow = ParseOptionalBoolAttribute(tableProperties, "bandRow");
        int bodyRowIndex = rowIndex - (firstRow ? 1 : 0);
        color = bandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0
            ? TintColor(accent, 0.4d)
            : TintColor(accent, 0.2d);
        return true;
    }

    private static bool TryReadBuiltInTableStyleTextColor(XElement table, int rowIndex, PptxTheme theme, out RgbColor color)
    {
        XElement? tableProperties = table.Element(DrawingNamespace + "tblPr");
        if (TryReadBuiltInTableStyle(table, out BuiltInTableStyle style) &&
            string.Equals(style.Name, "Medium-Style-2", StringComparison.Ordinal) &&
            rowIndex == 0 &&
            ParseOptionalBoolAttribute(tableProperties, "firstRow") &&
            theme.TryResolveColor("lt1", out color))
        {
            return true;
        }

        color = default;
        return false;
    }

    private static bool TryReadBuiltInTableStyle(XElement table, out BuiltInTableStyle style)
    {
        string? styleId = (string?)table
            .Element(DrawingNamespace + "tblPr")
            ?.Element(DrawingNamespace + "tableStyleId");
        return BuiltInTableStyles.TryGetValue(styleId ?? string.Empty, out style);
    }

    private static IReadOnlyDictionary<string, BuiltInTableStyle> BuiltInTableStyles { get; } =
        new Dictionary<string, BuiltInTableStyle>(StringComparer.OrdinalIgnoreCase)
        {
            ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = new("Medium-Style-2", "tx1"),
            ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"] = new("Medium-Style-2", "accent1"),
            ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"] = new("Medium-Style-2", "accent2"),
            ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = new("Medium-Style-2", "accent3"),
            ["{00A15C55-8517-42AA-B614-E9B94910E393}"] = new("Medium-Style-2", "accent4"),
            ["{7DF18680-E054-41AD-8BC1-D1AEF772440D}"] = new("Medium-Style-2", "accent5"),
            ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}"] = new("Medium-Style-2", "accent6")
        };

    private static RgbColor TintColor(RgbColor color, double tint)
    {
        return new RgbColor(
            ToByte(color.Red + (255d - color.Red) * tint),
            ToByte(color.Green + (255d - color.Green) * tint),
            ToByte(color.Blue + (255d - color.Blue) * tint));
    }

    private readonly record struct BuiltInTableStyle(string Name, string Accent);
}
