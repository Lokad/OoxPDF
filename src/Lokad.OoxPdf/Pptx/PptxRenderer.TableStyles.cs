using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryReadBuiltInTableStyleCellFill(XElement table, int rowIndex, PptxTheme theme, out RgbColor color, out double alpha)
    {
        alpha = 1d;
        if (!IsMediumStyle2Accent1Table(table) || !theme.TryResolveColor("accent1", out RgbColor accent))
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
        if (IsMediumStyle2Accent1Table(table) &&
            rowIndex == 0 &&
            ParseOptionalBoolAttribute(tableProperties, "firstRow") &&
            theme.TryResolveColor("lt1", out color))
        {
            return true;
        }

        color = default;
        return false;
    }

    private static bool IsMediumStyle2Accent1Table(XElement table)
    {
        string? styleId = (string?)table
            .Element(DrawingNamespace + "tblPr")
            ?.Element(DrawingNamespace + "tableStyleId");
        return string.Equals(styleId, "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", StringComparison.OrdinalIgnoreCase);
    }

    private static RgbColor TintColor(RgbColor color, double tint)
    {
        return new RgbColor(
            ToByte(color.Red + (255d - color.Red) * tint),
            ToByte(color.Green + (255d - color.Green) * tint),
            ToByte(color.Blue + (255d - color.Blue) * tint));
    }
}
