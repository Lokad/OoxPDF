using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneTable ReadTable(XElement frame, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? table = GetGraphicData(frame)?.Element(DrawingNamespace + "tbl");
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
            int nextColumn = 0;
            int continuationColumn = 0;
            foreach (XElement cell in row.Elements(DrawingNamespace + "tc"))
            {
                // A horizontal continuation occupies an already declared span.
                // Missing continuation placeholders retain compact-row recovery;
                // vertical continuations still own their declared grid footprint.
                bool coveredContinuation = OoxXml.ReadBool(cell, "hMerge") && continuationColumn < nextColumn;
                int columnIndex = coveredContinuation ? continuationColumn++ : nextColumn;
                PptxSceneTableCell sceneCell = ReadTableCell(cell, theme, colorMap, tableStyle, rowIndex, columnIndex, rowCount, columnCount);
                cells.Add(sceneCell);
                if (!coveredContinuation)
                {
                    continuationColumn = columnIndex + 1;
                    nextColumn = columnIndex + Math.Min(sceneCell.ColumnSpan, Math.Max(0, columnCount - columnIndex));
                }
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
            supported ? style.Kind : PptxBuiltInTableStyleKind.Unknown,
            supported ? style.Accent : string.Empty,
            supported,
            ReadTablePropertyFlag(tableProperties, "firstRow"),
            ReadTablePropertyFlag(tableProperties, "lastRow"),
            ReadTablePropertyFlag(tableProperties, "firstCol"),
            ReadTablePropertyFlag(tableProperties, "lastCol"),
            ReadTablePropertyFlag(tableProperties, "bandRow"),
            ReadTablePropertyFlag(tableProperties, "bandCol"));
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

    internal static PptxSceneTableCell ReadTableCell(XElement cell, PptxTheme theme, PptxColorMap colorMap, PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount)
    {
        (PptxSceneTextInsets textInsets, PptxSceneTableCellTextInsetSources textInsetSources, PptxSceneTextInsetValues textInsetValues) = ReadTableCellTextInsetInfo(cell);
        (PptxSceneTableCellVerticalAnchor verticalAnchor, string? verticalAnchorValue, PptxSceneTableCellVerticalAnchorSource verticalAnchorSource) = ReadTableCellVerticalAnchorInfo(cell);
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        XElement? bodyProperties = textBody?.Element(DrawingNamespace + "bodyPr");
        return new PptxSceneTableCell(
            columnIndex,
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
            // Explicit no-fill blocks the inherited fill for both placement and paint.
            cell.Element(DrawingNamespace + "tcPr")?.Element(DrawingNamespace + "noFill") is not null
                ? default
                : PptxTableStyleResolver.ReadCellFill(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap),
            PptxTableStyleResolver.ReadCellTextStyle(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap),
            textBody,
            TrimLeadingEmptyTableCellParagraphs(textBody),
            PptxRenderer.HasUnsupportedTextOrientation((string?)bodyProperties?.Attribute("vert")),
            HasUnsupportedTextVerticalOverflow(bodyProperties));
    }

    private static XElement? TrimLeadingEmptyTableCellParagraphs(XElement? textBody)
    {
        if (textBody is null)
        {
            return null;
        }

        XName paragraphName = DrawingNamespace + "p";
        XElement? firstContent = textBody.Elements(paragraphName).FirstOrDefault(ParagraphHasVisibleTextContent);
        // Layout treats this XML as read-only. Entirely empty cells keep their
        // paragraph formatting, and ordinary cells can share the source body.
        if (firstContent is null || ReferenceEquals(firstContent, textBody.Element(paragraphName)))
        {
            return textBody;
        }

        var trimmed = new XElement(textBody.Name, textBody.Attributes());
        bool leading = true;
        foreach (XElement element in textBody.Elements())
        {
            if (ReferenceEquals(element, firstContent))
            {
                leading = false;
            }
            if (!leading || element.Name != paragraphName)
            {
                // LINQ to XML copies children that already have a parent.
                trimmed.Add(element);
            }
        }
        return trimmed;
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
        string? cellValue = (string?)cellProperties?.Attribute(cellAttributeName);
        if (long.TryParse(cellValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long cellEmu))
        {
            return (OoxUnits.EmuToPoints(cellEmu), PptxSceneTableCellTextInsetSource.CellProperties, cellValue);
        }

        string? bodyValue = (string?)bodyProperties?.Attribute(bodyAttributeName);
        double points = OoxUnits.EmuToPoints(long.TryParse(bodyValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long bodyEmu)
            ? bodyEmu : defaultEmu);
        if (cellValue is not null)
        {
            return (points, PptxSceneTableCellTextInsetSource.CellProperties, cellValue);
        }

        return (points, bodyValue is null ? PptxSceneTableCellTextInsetSource.Default : PptxSceneTableCellTextInsetSource.BodyProperties, bodyValue);
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

    internal static PptxSceneFillStyle ReadTableCellFill(XElement cell, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return PptxColorResolver.TryReadSolidColorWithAlpha(cellProperties, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
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
            !PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, out RgbColor color, out double alpha))
        {
            return new PptxSceneTableCellBorder(IsSpecified: true, default);
        }

        double lineWidth = line.Attribute("w") is { } widthAttribute
            ? Math.Max(1d, OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture)) / 2d)
            : 0.75d;
        return new PptxSceneTableCellBorder(
            IsSpecified: true,
            ReadSceneLineStyle(line, color, lineWidth, alpha, widthSpecified: true));
    }
}
