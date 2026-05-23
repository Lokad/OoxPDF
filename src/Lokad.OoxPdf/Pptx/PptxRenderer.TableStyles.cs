using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryReadBuiltInTableStyleCellFill(XElement table, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, out RgbColor color, out double alpha)
    {
        alpha = 1d;
        if (!TryReadBuiltInTableStyle(table, out BuiltInTableStyle style) ||
            !theme.TryResolveColor(style.Accent, out RgbColor accent))
        {
            color = default;
            return false;
        }

        XElement? tableProperties = table.Element(DrawingNamespace + "tblPr");
        bool firstRow = ReadTablePropertyFlag(tableProperties, "firstRow");
        bool lastRow = ReadTablePropertyFlag(tableProperties, "lastRow");
        bool firstCol = ReadTablePropertyFlag(tableProperties, "firstCol");
        bool lastCol = ReadTablePropertyFlag(tableProperties, "lastCol");
        bool bandCol = ReadTablePropertyFlag(tableProperties, "bandCol");
        int bodyColumnIndex = columnIndex - (firstCol ? 1 : 0);
        if (string.Equals(style.Name, "Medium-Style-2", StringComparison.Ordinal) &&
            ((firstRow && rowIndex == 0) ||
                (lastRow && rowIndex == rowCount - 1) ||
                (firstCol && columnIndex == 0) ||
                (lastCol && columnIndex == columnCount - 1)))
        {
            color = accent;
            return true;
        }

        if (string.Equals(style.Name, "Light-Style-1", StringComparison.Ordinal) &&
            firstRow &&
            rowIndex == 0)
        {
            color = accent;
            return true;
        }

        if (string.Equals(style.Name, "Dark-Style-1", StringComparison.Ordinal))
        {
            if (firstRow && rowIndex == 0 && theme.TryResolveColor("dk1", out RgbColor dark))
            {
                color = dark;
                return true;
            }

            if ((firstCol && columnIndex == 0) ||
                (lastCol && columnIndex == columnCount - 1))
            {
                color = ShadeColor(accent, 0.6d);
                return true;
            }

            if (lastRow && rowIndex == rowCount - 1)
            {
                color = accent;
                return true;
            }
        }

        bool bandRow = ReadTablePropertyFlag(tableProperties, "bandRow");
        int bodyRowIndex = rowIndex - (firstRow ? 1 : 0);
        if (string.Equals(style.Name, "Light-Style-1", StringComparison.Ordinal))
        {
            if (bandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0)
            {
                color = accent;
                alpha = 0.4d;
                return true;
            }

            if (bandCol && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0)
            {
                color = accent;
                alpha = 0.4d;
                return true;
            }

            color = default;
            return false;
        }

        if (string.Equals(style.Name, "Medium-Style-2", StringComparison.Ordinal))
        {
            bool banded = (bandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (bandCol && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            color = banded
                ? TintColor(accent, 0.4d)
                : TintColor(accent, 0.2d);
            return true;
        }

        if (string.Equals(style.Name, "Dark-Style-1", StringComparison.Ordinal))
        {
            bool banded = (bandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (bandCol && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            color = banded
                ? ShadeColor(accent, 0.4d)
                : ShadeColor(accent, 0.2d);
            return true;
        }

        color = default;
        return false;
    }

    private static TableCellTextStyle ReadBuiltInTableStyleTextStyle(XElement table, int rowIndex, int columnIndex, PptxTheme theme)
    {
        XElement? tableProperties = table.Element(DrawingNamespace + "tblPr");
        bool bold = false;
        RgbColor? color = null;
        bool supportedStyle = TryReadBuiltInTableStyle(table, out BuiltInTableStyle style) &&
            (string.Equals(style.Name, "Medium-Style-2", StringComparison.Ordinal) ||
                string.Equals(style.Name, "Light-Style-1", StringComparison.Ordinal) ||
                string.Equals(style.Name, "Dark-Style-1", StringComparison.Ordinal));
        bool firstRow = ReadTablePropertyFlag(tableProperties, "firstRow") && rowIndex == 0;
        bool firstCol = ReadTablePropertyFlag(tableProperties, "firstCol") && columnIndex == 0;
        bool lastRow = ReadTablePropertyFlag(tableProperties, "lastRow") &&
            rowIndex == table.Elements(DrawingNamespace + "tr").Count() - 1;
        int columnCount = table
            .Element(DrawingNamespace + "tblGrid")
            ?.Elements(DrawingNamespace + "gridCol")
            .Count() ?? 0;
        bool lastCol = ReadTablePropertyFlag(tableProperties, "lastCol") &&
            columnCount > 0 &&
            columnIndex == columnCount - 1;
        if (supportedStyle &&
            firstRow &&
            theme.TryResolveColor("lt1", out RgbColor firstRowColor))
        {
            color = firstRowColor;
            bold = true;
        }

        if (supportedStyle && (firstCol || lastRow || lastCol))
        {
            bold = true;
            if ((string.Equals(style.Name, "Medium-Style-2", StringComparison.Ordinal) ||
                    string.Equals(style.Name, "Dark-Style-1", StringComparison.Ordinal)) &&
                (lastRow || lastCol) &&
                theme.TryResolveColor("lt1", out RgbColor conditionalColor))
            {
                color = conditionalColor;
            }
        }

        return new TableCellTextStyle(color, bold);
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
            ["{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"] = new("Light-Style-1", "tx1"),
            ["{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"] = new("Light-Style-1", "accent1"),
            ["{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}"] = new("Light-Style-1", "accent2"),
            ["{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}"] = new("Light-Style-1", "accent3"),
            ["{D27102A9-8310-4765-A935-A1911B00CA55}"] = new("Light-Style-1", "accent4"),
            ["{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}"] = new("Light-Style-1", "accent5"),
            ["{68D230F3-CF80-4859-8CE7-A43EE81993B5}"] = new("Light-Style-1", "accent6"),
            ["{E8034E78-7F5D-4C2E-B375-FC64B27BC917}"] = new("Dark-Style-1", "dk1"),
            ["{125E5076-3810-47DD-B79F-674D7AD40C01}"] = new("Dark-Style-1", "accent1"),
            ["{37CE84F3-28C3-443E-9E96-99CF82512B78}"] = new("Dark-Style-1", "accent2"),
            ["{D03447BB-5D67-496B-8E87-E561075AD55C}"] = new("Dark-Style-1", "accent3"),
            ["{E929F9F4-4A8F-4326-A1B4-22849713DDAB}"] = new("Dark-Style-1", "accent4"),
            ["{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}"] = new("Dark-Style-1", "accent5"),
            ["{AF606853-7671-496A-8E4F-DF71F8EC918B}"] = new("Dark-Style-1", "accent6"),
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

    private static RgbColor ShadeColor(RgbColor color, double shade)
    {
        return new RgbColor(
            ToByte(color.Red * shade),
            ToByte(color.Green * shade),
            ToByte(color.Blue * shade));
    }

    private readonly record struct BuiltInTableStyle(string Name, string Accent);

    private readonly record struct TableCellTextStyle(RgbColor? Color, bool Bold);
}
