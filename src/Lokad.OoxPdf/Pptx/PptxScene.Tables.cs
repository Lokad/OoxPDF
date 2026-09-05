using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneTable ReadTable(XElement frame, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? table = ReadTableElement(frame);
        IReadOnlyList<double> columnWidths = ReadTableColumnWidths(table);
        IReadOnlyList<double> rowHeights = ReadTableRowHeights(table);
        PptxSceneTableStyle style = ReadTableStyle(table);
        return new PptxSceneTable(
            columnWidths,
            rowHeights,
            ReadTableRows(table, theme, colorMap, style, rowHeights.Count, columnWidths.Count),
            style,
            table);
    }

    internal static XElement? ReadTableElement(XElement frame)
    {
        return frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl");
    }

    internal static IReadOnlyList<double> ReadTableColumnWidths(XElement? table)
    {
        return table
            ?.Element(DrawingNamespace + "tblGrid")
            ?.Elements(DrawingNamespace + "gridCol")
            .Select(column => Math.Max(1d, OoxXml.ParseOptionalLong(column, "w", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<double> ReadTableRowHeights(XElement? table)
    {
        return table
            ?.Elements(DrawingNamespace + "tr")
            .Select(row => Math.Max(1d, OoxXml.ParseOptionalLong(row, "h", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<PptxSceneTableRow> ReadTableRows(XElement? table, PptxTheme theme, PptxSceneTableStyle tableStyle, int rowCount, int columnCount)
    {
        return ReadTableRows(table, theme, PptxColorMap.Default, tableStyle, rowCount, columnCount);
    }

    internal static IReadOnlyList<PptxSceneTableRow> ReadTableRows(XElement? table, PptxTheme theme, PptxColorMap colorMap, PptxSceneTableStyle tableStyle, int rowCount, int columnCount)
    {
        if (table is null)
        {
            return [];
        }

        var rows = new List<PptxSceneTableRow>();
        int rowIndex = 0;
        foreach (XElement row in table.Elements(DrawingNamespace + "tr"))
        {
            var cells = new List<PptxSceneTableCell>();
            int columnIndex = 0;
            foreach (XElement cell in row.Elements(DrawingNamespace + "tc"))
            {
                PptxSceneTableCell sceneCell = ReadTableCell(cell, theme, colorMap, tableStyle, rowIndex, columnIndex, rowCount, columnCount);
                cells.Add(sceneCell);
                columnIndex += sceneCell.IsMergedContinuation ? 1 : sceneCell.ColumnSpan;
            }

            rows.Add(new PptxSceneTableRow(cells));
            rowIndex++;
        }

        return rows;
    }

    internal static PptxSceneTableStyle ReadTableStyle(XElement? table)
    {
        XElement? tableProperties = table?.Element(DrawingNamespace + "tblPr");
        string? styleId = (string?)tableProperties?.Element(DrawingNamespace + "tableStyleId");
        bool supported = PptxBuiltInTableStyles.TryGet(styleId, out PptxBuiltInTableStyle style);
        return new PptxSceneTableStyle(
            styleId,
            supported ? style.Name : string.Empty,
            supported ? style.Kind : PptxBuiltInTableStyleKind.Unknown,
            supported ? style.Accent : string.Empty,
            supported,
            ReadTablePropertyFlag(tableProperties, "firstRow"),
            ReadTablePropertyFlagValue(tableProperties, "firstRow"),
            ReadTablePropertyFlag(tableProperties, "lastRow"),
            ReadTablePropertyFlagValue(tableProperties, "lastRow"),
            ReadTablePropertyFlag(tableProperties, "firstCol"),
            ReadTablePropertyFlagValue(tableProperties, "firstCol"),
            ReadTablePropertyFlag(tableProperties, "lastCol"),
            ReadTablePropertyFlagValue(tableProperties, "lastCol"),
            ReadTablePropertyFlag(tableProperties, "bandRow"),
            ReadTablePropertyFlagValue(tableProperties, "bandRow"),
            ReadTablePropertyFlag(tableProperties, "bandCol"),
            ReadTablePropertyFlagValue(tableProperties, "bandCol"));
    }

    private static bool ReadTablePropertyFlag(XElement? tableProperties, string name)
    {
        if (tableProperties is null)
        {
            return false;
        }

        if (tableProperties.Attribute(name) is { } attribute)
        {
            return OoxBoolean.IsTrue(attribute.Value);
        }

        return tableProperties.Element(DrawingNamespace + name) is not null;
    }

    private static string ReadTablePropertyFlagValue(XElement? tableProperties, string name)
    {
        return (string?)tableProperties?.Attribute(name) ?? string.Empty;
    }

    internal static PptxSceneTableCell ReadTableCell(XElement cell, PptxTheme theme, PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount)
    {
        return ReadTableCell(cell, theme, PptxColorMap.Default, tableStyle, rowIndex, columnIndex, rowCount, columnCount);
    }

    internal static PptxSceneTableCell ReadTableCell(XElement cell, PptxTheme theme, PptxColorMap colorMap, PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount)
    {
        (PptxSceneTextInsets textInsets, PptxSceneTableCellTextInsetSources textInsetSources, PptxSceneTextInsetValues textInsetValues) = ReadTableCellTextInsetInfo(cell);
        (PptxSceneTableCellVerticalAnchor verticalAnchor, string? verticalAnchorValue, PptxSceneTableCellVerticalAnchorSource verticalAnchorSource) = ReadTableCellVerticalAnchorInfo(cell);
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        XElement? bodyProperties = textBody?.Element(DrawingNamespace + "bodyPr");
        int leadingEmptyParagraphCount = CountLeadingEmptyTableCellTextParagraphs(textBody);
        return new PptxSceneTableCell(
            ReadTableCellColumnSpan(cell),
            ReadTableCellRowSpan(cell),
            IsMergedTableCellContinuation(cell),
            textInsets,
            textInsetSources,
            textInsetValues,
            verticalAnchor,
            verticalAnchorValue,
            verticalAnchorSource,
            ReadTableCellFill(cell, theme, colorMap),
            ReadTableCellBorders(cell, theme, colorMap),
            PptxTableStyleResolver.ReadCellFill(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap),
            PptxTableStyleResolver.ReadCellTextStyle(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap),
            textBody,
            BuildTableCellLayoutTextBody(textBody, leadingEmptyParagraphCount),
            HasUnsupportedTextOrientation(bodyProperties),
            HasUnsupportedTextVerticalOverflow(bodyProperties),
            leadingEmptyParagraphCount);
    }

    private static XElement? BuildTableCellLayoutTextBody(XElement? textBody, int leadingEmptyParagraphCount)
    {
        if (textBody is null)
        {
            return null;
        }

        var textBodyCopy = new XElement(textBody.Name, textBody.Attributes(), textBody.Elements().Select(element => new XElement(element)));
        foreach (XElement paragraph in textBodyCopy.Elements(DrawingNamespace + "p").Take(leadingEmptyParagraphCount).ToArray())
        {
            paragraph.Remove();
        }

        return textBodyCopy;
    }

    private static int CountLeadingEmptyTableCellTextParagraphs(XElement? textBody)
    {
        if (textBody is null)
        {
            return 0;
        }

        XElement[] paragraphs = textBody.Elements(DrawingNamespace + "p").ToArray();
        if (!paragraphs.Any(ParagraphHasVisibleTextContent))
        {
            return 0;
        }

        int count = 0;
        foreach (XElement paragraph in paragraphs)
        {
            if (ParagraphHasVisibleTextContent(paragraph))
            {
                return count;
            }

            count++;
        }

        return 0;
    }

    private static bool ParagraphHasVisibleTextContent(XElement paragraph)
    {
        return paragraph.Elements().Any(child =>
            child.Name == DrawingNamespace + "r" ||
            child.Name == DrawingNamespace + "fld" ||
            child.Name == DrawingNamespace + "br");
    }

    internal static bool IsMergedTableCellContinuation(XElement cell)
    {
        return OoxXml.ReadBool(cell, "hMerge") || OoxXml.ReadBool(cell, "vMerge");
    }

    internal static int ReadTableCellColumnSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "gridSpan");
    }

    internal static int ReadTableCellRowSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "rowSpan");
    }

    private static int ReadTableCellSpan(XElement cell, string attributeName)
    {
        return cell.Attribute(attributeName) is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    internal static PptxSceneTextInsets ReadTableCellTextInsets(XElement cell)
    {
        return ReadTableCellTextInsetInfo(cell).Insets;
    }

    private static (PptxSceneTextInsets Insets, PptxSceneTableCellTextInsetSources Sources, PptxSceneTextInsetValues Values) ReadTableCellTextInsetInfo(XElement cell)
    {
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        XElement? bodyProperties = textBody?.Element(DrawingNamespace + "bodyPr");
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        (double left, PptxSceneTableCellTextInsetSource leftSource, string? leftValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "lIns", "marL", 91440);
        (double right, PptxSceneTableCellTextInsetSource rightSource, string? rightValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "rIns", "marR", 91440);
        (double top, PptxSceneTableCellTextInsetSource topSource, string? topValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "tIns", "marT", 45720);
        (double bottom, PptxSceneTableCellTextInsetSource bottomSource, string? bottomValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "bIns", "marB", 45720);
        return (
            new PptxSceneTextInsets(left, right, top, bottom),
            new PptxSceneTableCellTextInsetSources(leftSource, rightSource, topSource, bottomSource),
            new PptxSceneTextInsetValues(leftValue, rightValue, topValue, bottomValue));
    }

    private static (double Value, PptxSceneTableCellTextInsetSource Source, string? RawValue) ReadTableCellTextInset(
        XElement? bodyProperties,
        XElement? cellProperties,
        string bodyAttributeName,
        string cellAttributeName,
        long defaultEmu)
    {
        (double bodyValue, PptxSceneTableCellTextInsetSource bodySource, string? bodyRawValue) =
            ReadTableCellBodyInset(bodyProperties, bodyAttributeName, defaultEmu);
        if (cellProperties?.Attribute(cellAttributeName) is { } cellAttribute)
        {
            double cellValue = long.TryParse(cellAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long cellEmus)
                ? OoxUnits.EmuToPoints(cellEmus)
                : bodyValue;
            return (cellValue, PptxSceneTableCellTextInsetSource.CellProperties, cellAttribute.Value);
        }

        return (bodyValue, bodySource, bodyRawValue);
    }

    private static (double Value, PptxSceneTableCellTextInsetSource Source, string? RawValue) ReadTableCellBodyInset(
        XElement? bodyProperties,
        string attributeName,
        long defaultEmu)
    {
        double defaultValue = OoxUnits.EmuToPoints(defaultEmu);
        if (bodyProperties?.Attribute(attributeName) is { } bodyAttribute)
        {
            double value = long.TryParse(bodyAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long bodyEmus)
                ? OoxUnits.EmuToPoints(bodyEmus)
                : defaultValue;
            return (value, PptxSceneTableCellTextInsetSource.BodyProperties, bodyAttribute.Value);
        }

        return (defaultValue, PptxSceneTableCellTextInsetSource.Default, null);
    }

    private static double ReadInset(XElement? element, string attributeName, long defaultEmu)
    {
        long emu = element?.Attribute(attributeName) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultEmu;
        return OoxUnits.EmuToPoints(emu);
    }

    internal static PptxSceneTableCellVerticalAnchor ReadTableCellVerticalAnchor(XElement cell)
    {
        return ReadTableCellVerticalAnchorInfo(cell).Anchor;
    }

    private static (PptxSceneTableCellVerticalAnchor Anchor, string? Value, PptxSceneTableCellVerticalAnchorSource Source) ReadTableCellVerticalAnchorInfo(XElement cell)
    {
        string? anchor = ReadTableCellVerticalAnchorValue(cell);
        return (
            ParseTableCellVerticalAnchor(anchor),
            anchor,
            anchor is null
                ? PptxSceneTableCellVerticalAnchorSource.Default
                : PptxSceneTableCellVerticalAnchorSource.CellProperties);
    }

    private static PptxSceneTableCellVerticalAnchor ParseTableCellVerticalAnchor(string? anchor)
    {
        return anchor switch
        {
            "ctr" => PptxSceneTableCellVerticalAnchor.Middle,
            "b" => PptxSceneTableCellVerticalAnchor.Bottom,
            _ => PptxSceneTableCellVerticalAnchor.Top
        };
    }

    internal static string? ReadTableCellVerticalAnchorValue(XElement cell)
    {
        return (string?)cell
            .Element(DrawingNamespace + "tcPr")
            ?.Attribute("anchor");
    }

    internal static PptxSceneFillStyle ReadTableCellFill(XElement cell, PptxTheme theme)
    {
        return ReadTableCellFill(cell, theme, PptxColorMap.Default);
    }

    internal static PptxSceneFillStyle ReadTableCellFill(XElement cell, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return TryReadSolidColorWithAlpha(cellProperties, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    internal static PptxSceneTableCellBorders ReadTableCellBorders(XElement cell, PptxTheme theme)
    {
        return ReadTableCellBorders(cell, theme, PptxColorMap.Default);
    }

    internal static PptxSceneTableCellBorders ReadTableCellBorders(XElement cell, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return new PptxSceneTableCellBorders(
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnL"), theme, colorMap),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnR"), theme, colorMap),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnT"), theme, colorMap),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnB"), theme, colorMap));
    }

    private static PptxSceneTableCellBorder ReadTableCellBorder(XElement? line, PptxTheme theme, PptxColorMap colorMap)
    {
        if (line is null)
        {
            return default;
        }

        if (line.Element(DrawingNamespace + "noFill") is not null ||
            !TryReadSolidColorWithAlpha(line, theme, colorMap, out RgbColor color, out double alpha))
        {
            return new PptxSceneTableCellBorder(IsSpecified: true, default);
        }

        double lineWidth = line.Attribute("w") is { } widthAttribute
            ? Math.Max(1d, OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture)) / 2d)
            : 0.75d;
        XElement shapeProperties = WrapTableCellBorderLine(line);
        IReadOnlyList<double> dashPattern = TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> parsedDashPattern)
            ? parsedDashPattern
            : [];
        string? lineCap = ReadLineCap(shapeProperties);
        return new PptxSceneTableCellBorder(
            IsSpecified: true,
            new PptxSceneLineStyle(
                true,
                color,
                lineWidth,
                alpha,
                dashPattern,
                ReadPresetDashValue(shapeProperties),
                ReadLineCompound(shapeProperties),
                ReadLineCompoundValue(shapeProperties),
                lineCap switch
                {
                    "rnd" => 1,
                    "sq" => 2,
                    _ => null
                },
                lineCap,
                ReadLineJoin(shapeProperties),
                ReadLineJoinValue(shapeProperties), true));
    }

    private static XElement WrapTableCellBorderLine(XElement line)
    {
        return new XElement(
            DrawingNamespace + "spPr",
            new XElement(DrawingNamespace + "ln", line.Attributes(), line.Nodes()));
    }
}
